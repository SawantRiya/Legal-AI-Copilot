import math
import re
import uuid
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional

from pypdf import PdfReader


RISK_KEYWORDS = {
    "confidentiality": "Confidentiality and information-use obligations",
    "governing law": "Governing law and jurisdiction",
    "termination": "Termination and exit mechanics",
    "indemnity": "Indemnity / risk allocation",
    "liability": "Limitation of liability",
    "non-compete": "Non-compete or post-employment restrictions",
    "privacy": "Privacy and data-handling requirements",
    "ip": "Intellectual property ownership",
    "dispute": "Dispute resolution",
}


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _extract_key_clauses(document: str) -> List[str]:
    text = document.lower()
    matches: List[str] = []
    for keyword, label in RISK_KEYWORDS.items():
        if keyword in text:
            matches.append(label)
    if not matches:
        matches = ["General commercial terms"]
    return matches[:6]


MONTH_PATTERN = r"Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?"
DATE_PATTERN = re.compile(
    rf"\b(?:\d{{1,2}}(?:st|nd|rd|th)?\s+(?:{MONTH_PATTERN})\.?[,]?\s+\d{{4}}|"
    rf"(?:{MONTH_PATTERN})\.?\s+\d{{1,2}}(?:st|nd|rd|th)?[,]?\s+\d{{4}}|"
    r"\d{4}-\d{2}-\d{2}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b",
    re.IGNORECASE,
)
STOP_WORDS = {
    "a", "an", "and", "are", "as", "at", "be", "by", "for", "from", "in", "is", "it",
    "of", "on", "or", "the", "this", "to", "was", "what", "when", "where", "who", "with",
}
NO_ANSWER = "The current document does not contain enough information to answer this question."


def clean_extracted_text(text: str) -> str:
    text = (text or "").replace("\u00ad", "").replace("\ufb01", "fi").replace("\ufb02", "fl")
    text = re.sub(r"(?<=\w)-\s*\n\s*(?=\w)", "", text)
    text = re.sub(r"[\t\f\v ]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def _sentences_by_page(pages: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    sentences = []
    for page in pages:
        page_number = int(page.get("page", 1))
        lines = clean_extracted_text(str(page.get("text", ""))).splitlines()
        current = ""

        def append_current() -> None:
            nonlocal current
            for sentence in re.split(r"(?<=[.!?])\s+", current):
                sentence = re.sub(r"\s+", " ", sentence).strip()
                if sentence:
                    sentences.append({"page": page_number, "text": sentence})
            current = ""

        for line in lines:
            line = line.strip()
            if not line:
                append_current()
                continue
            if re.match(r"^\d+(?:\.\d+)*\.\s+[A-Z][A-Za-z0-9 &/,'-]{1,70}$", line):
                append_current()
                sentences.append({"page": page_number, "text": line})
                continue
            current = f"{current} {line}".strip()
            if re.search(r"[.!?]['\"]?$", line):
                append_current()
        append_current()
    return sentences


def _normalized_date(value: str) -> Optional[str]:
    from datetime import datetime

    candidate = re.sub(r"(\d)(st|nd|rd|th)\b", r"\1", value, flags=re.IGNORECASE).replace(",", "").strip()
    for date_format in ("%d %B %Y", "%d %b %Y", "%B %d %Y", "%b %d %Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(candidate, date_format).date().isoformat()
        except ValueError:
            continue
    return None


def _tokenize(text: str) -> List[str]:
    tokens = re.findall(r"[a-z0-9]+", (text or "").lower())
    normalized = []
    for token in tokens:
        if token in STOP_WORDS:
            continue
        if token.endswith("ies") and len(token) > 4:
            token = token[:-3] + "y"
        elif token.endswith("s") and len(token) > 3 and not token.endswith("ss"):
            token = token[:-1]
        normalized.append(token)
    return normalized


def create_document_chunks(pages: List[Dict[str, Any]], document_id: str, max_chars: int = 1200) -> List[Dict[str, Any]]:
    chunks: List[Dict[str, Any]] = []
    for page_item in pages:
        page_number = int(page_item.get("page", 1))
        lines = clean_extracted_text(str(page_item.get("text", ""))).splitlines()
        section = "Document text"
        for line in lines:
            line = line.strip()
            if not line:
                continue
            heading = re.match(r"^(?:\d+(?:\.\d+)*[.)]?\s+)?([^:]{2,100}):\s*(.*)$", line)
            if heading:
                section = heading.group(1).strip()
                body = heading.group(2).strip()
                text = line if body else section
            else:
                text = line
                if len(line) <= 100 and (line.isupper() or re.match(r"^\d+(?:\.\d+)*[.)]?\s+\S.{0,70}$", line)):
                    section = line
            if len(text) <= max_chars:
                chunks.append({"document_id": document_id, "page": page_number, "text": text, "section": section})
                continue
            for offset in range(0, len(text), max_chars):
                chunks.append({"document_id": document_id, "page": page_number, "text": text[offset:offset + max_chars], "section": section})
    return chunks


def _extract_date_facts(pages: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    facts: List[Dict[str, Any]] = []
    for source in _sentences_by_page(pages):
        page_number = source["page"]
        sentence = source["text"]
        lower = sentence.casefold()
        matches = list(DATE_PATTERN.finditer(sentence))
        for match in matches:
            before = lower[max(0, match.start() - 90):match.start()]
            if re.search(r"made on|dated|document date|executed on|\bdate\s*[:\-]", before):
                kind = "document date"
            elif re.search(r"expir|end|until|through|terminat|ending", before):
                kind = "agreement end"
            elif re.search(r"commenc|begin|start|effective|from", before):
                kind = "agreement start"
            elif re.search(r"renew", before):
                kind = "renewal"
            elif re.search(r"due|pay", before):
                kind = "payment due"
            else:
                kind = "other date"
            facts.append({"date": match.group(0), "kind": kind, "page": page_number, "text": sentence})

        if re.search(r"(?:rent|payment).{0,80}(?:due|pay|payable)|(?:due|pay|payable).{0,80}(?:rent|payment)", lower):
            due_match = re.search(r"\b(?:\d{1,2}(?:st|nd|rd|th)?\s+day of (?:each|every) (?:calendar )?month|\d{1,2}(?:st|nd|rd|th)? of (?:each|every) month)\b", sentence, re.IGNORECASE)
            if due_match:
                facts.append({"date": due_match.group(0), "kind": "payment due", "page": page_number, "text": sentence})
        if re.search(r"notice|terminate|termination", lower):
            notice_match = re.search(r"\b(?:\d+|one|two|three|four|five|six|seven|eight|nine|ten|thirty|forty|forty-five|sixty|ninety)\s+days?\b", sentence, re.IGNORECASE)
            deadline_after_notice = notice_match and re.search(rf"\b(?:within|after)\s+{re.escape(notice_match.group(0))}\b", sentence, re.IGNORECASE)
            if notice_match and not deadline_after_notice:
                facts.append({"date": notice_match.group(0), "kind": "notice period", "page": page_number, "text": sentence})
    unique = []
    seen = set()
    for fact in facts:
        key = (fact["date"].casefold(), fact["kind"])
        if key not in seen:
            seen.add(key)
            unique.append(fact)
    return unique


def _parse_dates(text: str, pages: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
    source_pages = pages or [{"page": 1, "text": text}]
    facts = _extract_date_facts(source_pages)
    start = next((fact["date"] for fact in facts if fact["kind"] == "agreement start"), None)
    end = next((fact["date"] for fact in facts if fact["kind"] == "agreement end"), None)
    notice = next((fact["date"] for fact in facts if fact["kind"] == "notice period"), None)
    due = next((fact["date"] for fact in facts if fact["kind"] == "payment due"), None)
    document_date = next((fact["date"] for fact in facts if fact["kind"] == "document date"), None)
    return {
        "dates": list(dict.fromkeys(fact["date"] for fact in facts if fact["kind"] != "other date")),
        "facts": facts,
        "items": [
            {"type": fact["kind"], "date": fact["date"], "normalized_date": _normalized_date(fact["date"]), "description": fact["text"], "source_page": fact["page"], "source_text": fact["text"]}
            for fact in facts
        ],
        "start_date": start,
        "end_date": end,
        "document_date": document_date,
        "notice_date": notice,
        "rent_due_date": due,
    }


def _answer_question(document: str, question: str) -> str:
    question_analysis = analyze_document("document", document)
    return answer_question_from_document(question, question_analysis)["answer"]


def analyze_legal_document(document: str, question: Optional[str] = None) -> Dict[str, Any]:
    text = _normalize(document)
    lower_text = text.lower()

    key_clauses = _extract_key_clauses(text)
    issues: List[str] = []

    if "confidential" in lower_text:
        issues.append("Confidentiality protections are present and should be reviewed for scope, carve-outs, and permitted disclosures.")
    else:
        issues.append("No confidentiality language detected; sensitive information protections may be insufficient.")

    if "indemnity" not in lower_text:
        issues.append("No indemnity protection is present; risk allocation may be unbalanced.")

    if "limitation of liability" not in lower_text and "liability" not in lower_text:
        issues.append("No clear limitation-of-liability clause is present; the document may expose the parties to excessive risk.")

    if "termination" not in lower_text and "terminate" not in lower_text:
        issues.append("No clear termination or termination-for-cause language is present.")

    if "governing law" in lower_text:
        issues.append("Governing law provisions are identified and should be validated against operational and regulatory requirements.")

    risk_score = 20

    if "confidential" in lower_text:
        risk_score += 15
    else:
        risk_score += 12

    if "indemnity" not in lower_text:
        risk_score += 20
    if "limitation of liability" not in lower_text and "liability" not in lower_text:
        risk_score += 20
    if "termination" not in lower_text and "terminate" not in lower_text:
        risk_score += 15
    if "governing law" in lower_text:
        risk_score += 10
    if "non-compete" in lower_text or "non solicitation" in lower_text or "no non-solicitation" in lower_text:
        risk_score += 10

    risk_score = max(0, min(100, risk_score))

    if risk_score < 35:
        risk_level = "low"
    elif risk_score < 70:
        risk_level = "moderate"
    else:
        risk_level = "high"

    summary = (
        f"This legal document includes {', '.join(key_clauses[:3])}. "
        f"The governing law and risk allocation provisions should be checked for completeness and enforceability. "
        f"Overall risk is {risk_level} with a score of {risk_score}/100."
    )

    return {
        "summary": summary,
        "risk_level": risk_level,
        "risk_score": risk_score,
        "issues": issues[:5],
        "key_clauses": key_clauses,
        "answer": _answer_question(text, question) if question else "No question was asked.",
    }


def extract_pdf_text(file_path: str) -> Dict[str, Any]:
    try:
        reader = PdfReader(file_path)
        pages: List[Dict[str, Any]] = []
        text_chunks: List[str] = []
        for index, page in enumerate(reader.pages, start=1):
            page_text = page.extract_text() or ""
            pages.append({"page": index, "text": page_text, "has_images": bool(page.images)})
            text_chunks.append(page_text)
        extracted_text = "\n\n".join(text_chunks)
        return {"ok": True, "text": extracted_text, "pages": pages}
    except Exception:
        return {"ok": False, "text": "", "pages": []}


def infer_document_type(file_name: str, text: str) -> str:
    name = file_name.lower()
    lower = text.lower()
    if "rental" in name or "lease" in lower or "landlord" in lower or "tenant" in lower:
        return "Rental Agreement"
    if "employment" in name or "offer" in lower or "salary" in lower or "employee" in lower:
        return "Employment Agreement"
    if "nda" in name or "non disclosure" in lower or "confidential information" in lower:
        return "NDA"
    if "notice" in name or "legal notice" in lower or "demand" in lower:
        return "Legal Notice"
    if "service" in name or "vendor" in lower or "consultant" in lower:
        return "Service Agreement"
    return "Legal Agreement"


def _extract_parties(text: str, pages: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    parties: List[Dict[str, Any]] = []
    patterns = (
        ("Landlord", r"\bLandlord\s*[:\-]\s*([^\n,.;]{2,80})"),
        ("Tenant", r"\bTenant\s*[:\-]\s*([^\n,.;]{2,80})"),
        ("Landlord", r"([A-Z][A-Za-z.'-]+(?:\s+[A-Z][A-Za-z.'-]+){1,3})\s*\(\s*Landlord\s*\)"),
        ("Tenant", r"([A-Z][A-Za-z.'-]+(?:\s+[A-Z][A-Za-z.'-]+){1,3})\s*\(\s*Tenant\s*\)"),
        ("Landlord", r"between\s+([A-Z][A-Za-z.'-]+(?:\s+[A-Z][A-Za-z.'-]+){1,3})\s*,?\s+Landlord\b"),
        ("Tenant", r"and\s+([A-Z][A-Za-z.'-]+(?:\s+[A-Z][A-Za-z.'-]+){1,3})\s*,?\s+Tenant\b"),
    )
    for role, pattern in patterns:
        match = re.search(pattern, text)
        if not match:
            continue
        name = re.sub(r"\s+", " ", match.group(1)).strip()
        if name.casefold() in {"landlord", "tenant", "lessor", "lessee", "party", "the landlord", "the tenant"}:
            continue
        if any(party["role"] == role for party in parties):
            continue
        source = next((item for item in _sentences_by_page(pages) if name.casefold() in item["text"].casefold() and role.casefold() in item["text"].casefold()), None)
        address = re.search(rf"\b{role}(?:'s)?\s+address\s*[:\-]\s*([^.;]+)", text, re.IGNORECASE)
        contact = re.search(rf"\b{role}(?:'s)?\s+(?:contact|phone|email)\s*[:\-]\s*([^.;]+)", text, re.IGNORECASE)
        parties.append({
            "name": name,
            "role": role,
            "address": address.group(1).strip() if address else None,
            "contact": contact.group(1).strip() if contact else None,
            "organization": "Party to the agreement",
            "source_page": source["page"] if source else None,
            "source_text": source["text"] if source else None,
        })
    return parties


def _extract_property(pages: List[Dict[str, Any]]) -> Dict[str, Any]:
    patterns = (
        r"(?:premises|property)\b.*?\blocated at\s+(.+?)(?:\.|;|$)",
        r"(?:premises|property)\s+(?:is\s+)?(?:located at|address(?:ed)? at)\s+(.+?)(?:\.|;|$)",
        r"(?:property|premises)\s+address\s*[:\-]\s*(.+?)(?:\.|;|$)",
    )
    for sentence in _sentences_by_page(pages):
        for pattern in patterns:
            match = re.search(pattern, sentence["text"], re.IGNORECASE)
            if match:
                return {"address": match.group(1).strip(), "source_page": sentence["page"], "source_text": sentence["text"]}
    return {"address": None, "source_page": None, "source_text": None}


def _concise_obligation(sentence: str, party: str) -> str:
    prefix = re.compile(rf"^(?:the\s+)?{party}\s+(?:shall|must|will|is required to|agrees to)\s+", re.IGNORECASE)
    concise = prefix.sub("", sentence).strip()
    if concise:
        return concise[0].upper() + concise[1:]
    return sentence.strip()


def _extract_obligations(pages: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    result: List[Dict[str, Any]] = []
    obligation_terms = re.compile(r"\b(shall|must|is responsible|are responsible|agrees? to|is required to|are required to|will)\b", re.IGNORECASE)
    for source in _sentences_by_page(pages):
        sentence = source["text"]
        if not obligation_terms.search(sentence):
            continue
        role = None
        if re.match(r"^(?:the\s+)?tenant\b", sentence, re.IGNORECASE):
            role = "Tenant"
        elif re.match(r"^(?:the\s+)?landlord\b", sentence, re.IGNORECASE) or re.search(r"responsibility of (?:the )?landlord\b", sentence, re.IGNORECASE):
            role = "Landlord"
        elif re.match(r"^monthly rent\b", sentence, re.IGNORECASE) and re.search(r"\bpayable\b", sentence, re.IGNORECASE):
            role = "Tenant"
        if role and not any(item["party"] == role and item["source_text"] == sentence for item in result):
            result.append({
                "party": role,
                "obligation": _concise_obligation(sentence, role),
                "source_page": source["page"],
                "source_text": sentence,
            })
    return result


def _obligation_summary(items: List[Dict[str, Any]], party: str) -> List[str]:
    return [item["obligation"] for item in items if item["party"].casefold() == party.casefold()]


def _expand_query(question: str) -> List[str]:
    terms = _tokenize(question)
    expansions = {
        "obligation": ["duty", "responsibility", "shall", "must", "required"],
        "notice": ["termination", "written", "days"],
        "period": ["term", "commence", "expire", "begin", "end"],
        "date": ["commence", "effective", "expire", "due", "renewal", "notice"],
        "landlord": ["lessor", "owner"],
        "tenant": ["lessee", "renter"],
    }
    expanded = list(terms)
    for key, values in expansions.items():
        if key in terms or (key == "obligation" and any(term.startswith("obligation") or term in {"duty", "responsibility"} for term in terms)):
            expanded.extend(values)
    if "important" in terms and ("date" in terms or "dates" in terms):
        expanded.extend(["start", "end", "payment", "due", "notice", "renewal", "effective"])
    return list(dict.fromkeys(expanded))


def retrieve_relevant_chunks(question: str, chunks: List[Dict[str, Any]], document_id: str) -> List[Dict[str, Any]]:
    scoped = [chunk for chunk in chunks if chunk.get("document_id") == document_id and chunk.get("text", "").strip()]
    if not scoped:
        return []
    tokenized = [_tokenize(f"{chunk.get('section', '')} {chunk['text']}") for chunk in scoped]
    document_frequency = Counter(token for terms in tokenized for token in set(terms))
    average_length = sum(map(len, tokenized)) / len(tokenized) or 1
    query_terms = _expand_query(question)
    scored = []
    for chunk, terms in zip(scoped, tokenized):
        frequencies = Counter(terms)
        score = 0.0
        for token in query_terms:
            frequency = frequencies[token]
            if not frequency:
                continue
            inverse_frequency = math.log(1 + (len(scoped) - document_frequency[token] + 0.5) / (document_frequency[token] + 0.5))
            denominator = frequency + 1.5 * (1 - 0.75 + 0.75 * len(terms) / average_length)
            score += inverse_frequency * frequency * 2.5 / denominator
        if score:
            scored.append((score, chunk))
    scored.sort(key=lambda item: item[0], reverse=True)
    return [chunk for _, chunk in scored[:6]]


def _source_pages(chunks: List[Dict[str, Any]]) -> str:
    page_numbers = sorted({int(chunk.get("page", 1)) for chunk in chunks})
    if not page_numbers:
        return "N/A"
    if len(page_numbers) == 1:
        return f"Page {page_numbers[0]}"
    return "Pages " + ", ".join(map(str, page_numbers))


def _grounded_response(answer: str, chunks: List[Dict[str, Any]]) -> Dict[str, Any]:
    unique = []
    seen = set()
    for chunk in chunks:
        source_text = chunk.get("evidence_text", chunk.get("text", ""))
        key = (chunk.get("page"), source_text)
        if key not in seen:
            seen.add(key)
            unique.append(chunk)
    evidence = "\n".join(chunk.get("evidence_text", chunk["text"]).strip() for chunk in unique if chunk.get("evidence_text", chunk.get("text", "")).strip())
    return {
        "answer": answer,
        "source": _source_pages(unique),
        "evidence": evidence,
        "evidence_chunks": unique,
    }


def _records_as_evidence(records: List[Dict[str, Any]], document_id: str, text_key: str = "source_text", page_key: str = "source_page") -> List[Dict[str, Any]]:
    chunks = []
    for record in records:
        text = record.get(text_key)
        page = record.get(page_key)
        if text and page is not None:
            chunks.append(_source_chunk(document_id, int(page), text, record.get("type", record.get("party", "Document clause"))))
    return chunks


def _safe_money(value: str) -> str:
    if not value:
        return "Not specified"
    return value


def _source_chunk(document_id: str, page: int, text: str, section: str = "Document text") -> Dict[str, Any]:
    return {"document_id": document_id, "page": page, "text": text, "evidence_text": text, "section": section}


def _extract_financial_terms(pages: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    terms = []
    for source in _sentences_by_page(pages):
        sentence = source["text"]
        rent_match = re.search(r"monthly rent\D{0,40}(?:(INR|Rs\.?|₹)\s*)?([0-9][0-9,]*)", sentence, re.IGNORECASE)
        if rent_match:
            currency = rent_match.group(1) or ("INR" if "inr" in sentence.casefold() else "INR")
            amount = rent_match.group(2).rstrip(",")
            due_match = re.search(r"(?:payable|due)\s+((?:on or before|by|on)\s+[^.;,]+(?:\s+(?:of|each|every|calendar)\s+(?:each\s+|every\s+|calendar\s+)?month)?)", sentence, re.IGNORECASE)
            due_date = due_match.group(1).strip() if due_match else None
            terms.append({
                "type": "monthly_rent",
                "amount": amount,
                "currency": "INR" if currency.casefold() in {"inr", "rs", "rs."} else currency,
                "display_amount": f"₹{amount}" if currency.casefold() in {"inr", "rs", "rs."} else f"{currency}{amount}",
                "frequency": "monthly",
                "due_date": due_date,
                "source_page": source["page"],
                "source_text": sentence,
            })
        deposit_match = re.search(r"security deposit\s+(?:of|is|shall be|amounting to)\s+(?:(INR|Rs\.?|₹)\s*)?([0-9][0-9,]*)", sentence, re.IGNORECASE)
        if deposit_match:
            currency = deposit_match.group(1) or ("INR" if "inr" in sentence.casefold() else "INR")
            amount = deposit_match.group(2).rstrip(",")
            terms.append({
                "type": "security_deposit",
                "amount": amount,
                "currency": "INR" if currency.casefold() in {"inr", "rs", "rs."} else currency,
                "display_amount": f"₹{amount}" if currency.casefold() in {"inr", "rs", "rs."} else f"{currency}{amount}",
                "frequency": "one-time",
                "due_date": None,
                "source_page": source["page"],
                "source_text": sentence,
            })
    unique = []
    seen = set()
    for term in terms:
        key = (term["type"], term["amount"], term["source_page"])
        if key not in seen:
            seen.add(key)
            unique.append(term)
    return unique


def _extract_clause_items(chunks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    grouped: Dict[tuple, List[str]] = {}
    for chunk in chunks:
        section = chunk.get("section", "Document text")
        if section == "Document text":
            continue
        key = (section, chunk.get("page", 1))
        grouped.setdefault(key, []).append(chunk["text"])
    return [
        {"title": title, "content": " ".join(dict.fromkeys(contents)), "source_page": page, "source_text": " ".join(dict.fromkeys(contents))}
        for (title, page), contents in grouped.items()
    ]


def _duration_from_source(dates: Dict[str, Any], text: str) -> Optional[str]:
    explicit = re.search(r"\b(?:period of\s+)?(\d+\s+months?)\b", text, re.IGNORECASE)
    if explicit:
        return explicit.group(1)
    start = dates.get("start_date")
    end = dates.get("end_date")
    if start and end:
        start_iso = _normalized_date(start)
        end_iso = _normalized_date(end)
        if start_iso and end_iso:
            from datetime import date

            start_date = date.fromisoformat(start_iso)
            end_date = date.fromisoformat(end_iso)
            months = (end_date.year - start_date.year) * 12 + end_date.month - start_date.month
            if end_date.day >= start_date.day and months > 0:
                return f"{months} months"
    return None


def analyze_document(
    file_name: str,
    text: str,
    pages: Optional[List[Dict[str, Any]]] = None,
    document_id: Optional[str] = None,
) -> Dict[str, Any]:
    document_id = document_id or str(uuid.uuid4())
    source_pages = pages or [{"page": 1, "text": text}]
    source_pages = [
        {
            "page": int(page.get("page", index)),
            "text": clean_extracted_text(str(page.get("text", ""))),
            "has_images": bool(page.get("has_images", False)),
        }
        for index, page in enumerate(source_pages, start=1)
    ]
    text = "\n\n".join(re.sub(r"(?<![.!?])\n(?!\n)", " ", page["text"]) for page in source_pages).strip()
    document_type = infer_document_type(file_name, text)
    chunks = create_document_chunks(source_pages, document_id)
    parties = _extract_parties(text, source_pages)
    dates = _parse_dates(text, source_pages)
    duration = _duration_from_source(dates, text) or "Not specified"
    if duration != "Not specified" and dates["start_date"] and dates["end_date"]:
        duration = f"{duration} ({dates['start_date']} to {dates['end_date']})"
    property_info = _extract_property(source_pages)
    financial_terms = _extract_financial_terms(source_pages)
    monthly_rent = next((term for term in financial_terms if term["type"] == "monthly_rent"), None)
    security_deposit = next((term for term in financial_terms if term["type"] == "security_deposit"), None)
    financial = {
        "monthly_rent": monthly_rent["display_amount"] if monthly_rent else "Not specified",
        "security_deposit": security_deposit["display_amount"] if security_deposit else "Not specified",
        "payment_due_date": monthly_rent["due_date"] if monthly_rent and monthly_rent["due_date"] else (dates["rent_due_date"] or "Not specified"),
        "penalties": "Not specified",
    }

    clause_items = _extract_clause_items(chunks)
    clauses = list(dict.fromkeys(clause["title"] for clause in clause_items))
    if not clauses:
        clauses = _extract_key_clauses(text)
    obligation_items = _extract_obligations(source_pages)
    obligations = {
        "tenant": _obligation_summary(obligation_items, "Tenant"),
        "landlord": _obligation_summary(obligation_items, "Landlord"),
    }
    party_names = {party["role"]: party["name"] for party in parties}
    if document_type == "Rental Agreement" and party_names:
        landlord_name = party_names.get("Landlord")
        tenant_name = party_names.get("Tenant")
        named_parties = []
        if landlord_name:
            named_parties.append(f"{landlord_name} (landlord)")
        if tenant_name:
            named_parties.append(f"{tenant_name} (tenant)")
        summary = f"This rental agreement identifies {' and '.join(named_parties)}."
    else:
        summary = f"This document is a {document_type.lower()}."
    if property_info["address"]:
        summary += f" The property is at {property_info['address']}."
    summary += " It covers " + ", ".join(clauses[:5]) + "." if clauses else ""
    timeline = [
        {
            "date": item["date"],
            "event": item["type"],
            "description": item["description"],
            "page": item["source_page"],
            "source_text": item["source_text"],
        }
        for item in dates["items"]
        if item["type"] in {"agreement start", "agreement end", "payment due", "notice period", "renewal", "document date"}
    ]

    notice_fact = next((fact for fact in dates["facts"] if fact["kind"] == "notice period"), None)
    notice_match = re.search(r"\b(\d+|one|two|three|four|five|six|seven|eight|nine|ten|thirty|forty|forty-five|sixty|ninety)\s+days?\b", notice_fact["date"], re.IGNORECASE) if notice_fact else None
    notice = {
        "duration": notice_match.group(1) if notice_match else None,
        "unit": "days" if notice_match else None,
        "description": notice_fact["text"] if notice_fact else None,
        "source_page": notice_fact["page"] if notice_fact else None,
        "source_text": notice_fact["text"] if notice_fact else None,
    }

    graph_nodes = [{"id": "document", "label": document_type, "type": "document"}]
    graph_edges = []
    for index, party in enumerate(parties, start=1):
        node_id = f"party{index}"
        graph_nodes.append({"id": node_id, "label": party["name"], "type": party["role"].lower(), "source_page": party["source_page"], "details": party["source_text"]})
        graph_edges.append({"from": "document", "to": node_id, "label": party["role"]})
    if monthly_rent:
        graph_nodes.append({"id": "rent", "label": monthly_rent["display_amount"], "type": "financial", "source_page": monthly_rent["source_page"], "details": monthly_rent["source_text"]})
        graph_edges.append({"from": "document", "to": "rent", "label": "Monthly rent"})
    if property_info["address"]:
        graph_nodes.append({"id": "property", "label": property_info["address"], "type": "property", "source_page": property_info["source_page"], "details": property_info["source_text"]})
        graph_edges.append({"from": "document", "to": "property", "label": "Property"})
    if duration != "Not specified":
        graph_nodes.append({"id": "period", "label": duration, "type": "date", "source_page": dates["facts"][0]["page"] if dates["facts"] else None, "details": duration})
        graph_edges.append({"from": "document", "to": "period", "label": "Agreement period"})

    return {
        "document_id": document_id,
        "document_name": file_name,
        "document_type": document_type,
        "summary": summary,
        "what_is_this_document_about": summary,
        "parties": parties,
        "dates": {
            "start_date": dates["start_date"],
            "end_date": dates["end_date"],
            "document_date": dates["document_date"],
            "notice_date": dates["notice_date"],
            "rent_due_date": dates["rent_due_date"],
            "facts": dates["facts"],
            "items": dates["items"],
        },
        "property": property_info,
        "duration": duration,
        "financial": financial,
        "financial_terms": financial_terms,
        "clauses": clauses,
        "clause_items": clause_items,
        "obligations": obligations,
        "obligation_items": obligation_items,
        "notice": notice,
        "key_terms": {
            "term": duration if duration != "Not specified" else None,
            "rent": monthly_rent,
            "deposit": security_deposit,
            "notice": notice,
            "rent_due": monthly_rent["due_date"] if monthly_rent else None,
        },
        "timeline": timeline,
        "graph": {"title": document_type, "nodes": graph_nodes, "edges": graph_edges},
        "pages": source_pages,
        "chunks": chunks,
        "raw_text": text,
        "risk": analyze_legal_document(text),
    }


def answer_question_from_document(
    question: str,
    analysis: Dict[str, Any],
    document_id: Optional[str] = None,
) -> Dict[str, Any]:
    document_id = document_id or analysis.get("document_id")
    chunks = analysis.get("chunks") or create_document_chunks(
        analysis.get("pages") or [{"page": 1, "text": analysis.get("raw_text", "")}],
        document_id or "current",
    )
    scoped_id = document_id or "current"
    chunks = [chunk for chunk in chunks if chunk.get("document_id") == scoped_id]
    retrieved = retrieve_relevant_chunks(question, chunks, scoped_id)
    query = (question or "").casefold()
    if not query.strip():
        return _grounded_response(NO_ANSWER, retrieved[:1] or chunks[:1])

    financial_terms = analysis.get("financial_terms", [])
    if "rent" in query:
        rent = next((term for term in financial_terms if term["type"] == "monthly_rent"), None)
        if rent:
            answer = f"{rent['display_amount']} per month"
            if rent.get("due_date"):
                answer += f", payable {rent['due_date']}"
            evidence = _records_as_evidence([rent], scoped_id)
            return _grounded_response(answer + ".", evidence)
        return _grounded_response(NO_ANSWER, retrieved[:2] or chunks[:2])

    if "deposit" in query:
        deposit = next((term for term in financial_terms if term["type"] == "security_deposit"), None)
        if deposit:
            return _grounded_response(f"The security deposit is {deposit['display_amount']}.", _records_as_evidence([deposit], scoped_id))
        return _grounded_response(NO_ANSWER, retrieved[:2] or chunks[:2])

    obligation_query = any(word in query for word in ("obligation", "obligations", "duty", "duties", "responsibilit"))
    tenant_obligation = obligation_query and any(word in query for word in ("tenant", "lessee", "renter"))
    landlord_obligation = obligation_query and any(word in query for word in ("landlord", "lessor"))
    if tenant_obligation or landlord_obligation:
        role = "Tenant" if tenant_obligation else "Landlord"
        obligations = [item for item in analysis.get("obligation_items", []) if item["party"] == role]
        if obligations:
            descriptions = [re.sub(r"[.;]$", "", item["obligation"]).strip() for item in obligations]
            joined = "; ".join(description[0].lower() + description[1:] if description else description for description in descriptions)
            answer = f"The {role.lower()} must {joined}."
            return _grounded_response(answer, _records_as_evidence(obligations, scoped_id))
        return _grounded_response(NO_ANSWER, retrieved[:2] or chunks[:2])

    important_dates_query = "date" in query and any(word in query for word in ("important", "key", "all", "what"))
    if important_dates_query:
        items = analysis.get("dates", {}).get("items", [])
        important_types = {"agreement start", "agreement end", "payment due", "notice period", "renewal", "document date", "other date"}
        selected = []
        seen = set()
        for item in items:
            key = (item["type"], item["date"].casefold())
            if item["type"] in important_types and key not in seen:
                seen.add(key)
                selected.append(item)
        if selected:
            labels = {
                "agreement start": "Agreement start",
                "agreement end": "Agreement end",
                "payment due": "Rent due",
                "notice period": "Notice period",
                "renewal": "Renewal",
                "document date": "Document date",
                "other date": "Other date",
            }
            answer = "; ".join(f"{labels.get(item['type'], item['type'].title())}: {item['date']}" for item in selected)
            return _grounded_response(answer, _records_as_evidence(selected, scoped_id, text_key="source_text"))
        return _grounded_response(NO_ANSWER, retrieved[:2] or chunks[:2])

    if "notice" in query or "termination" in query:
        notice = analysis.get("notice", {})
        if notice.get("duration") and notice.get("source_text"):
            source = notice["source_text"]
            written = "' written" if re.search(r"\bdays?'?\s+written notice", source, re.IGNORECASE) else ""
            answer = f"{notice['duration']} days{written} notice is required for termination."
            return _grounded_response(answer, [_source_chunk(scoped_id, notice["source_page"], source, "Notice and termination")])
        relevant = [chunk for chunk in retrieved if re.search(r"notice|terminat", chunk["text"], re.IGNORECASE)]
        return _grounded_response(NO_ANSWER, relevant or retrieved[:2] or chunks[:2])

    period_query = any(word in query for word in ("period", "duration", "term")) and not any(word in query for word in ("notice", "payment", "tenant", "landlord"))
    if period_query:
        dates = analysis.get("dates", {})
        start = dates.get("start_date")
        end = dates.get("end_date")
        relevant_items = [item for item in dates.get("items", []) if item["type"] in {"agreement start", "agreement end"}]
        if start and end and relevant_items:
            answer = f"The agreement runs from {start} to {end}."
            duration = analysis.get("duration")
            if duration and duration != "Not specified":
                answer = f"The agreement runs for {duration.split(' (', 1)[0]}, from {start} to {end}."
            return _grounded_response(answer, _records_as_evidence(relevant_items, scoped_id))
        return _grounded_response(NO_ANSWER, retrieved[:2] or chunks[:2])

    if any(word in query for word in ("landlord", "lessor", "owner", "tenant", "lessee", "renter")) and any(word in query for word in ("who", "name", "address", "contact", "information")):
        role = "Landlord" if any(word in query for word in ("landlord", "lessor", "owner")) else "Tenant"
        party = next((item for item in analysis.get("parties", []) if item.get("role", "").casefold() == role.casefold()), None)
        if party:
            wants_address = "address" in query
            value = party.get("address") if wants_address else party["name"]
            if value and party.get("source_text"):
                answer = f"{role} address: {value}." if wants_address else f"The {role.lower()} is {value}."
                evidence_record = {"source_page": party["source_page"], "source_text": party["source_text"]}
                return _grounded_response(answer, _records_as_evidence([evidence_record], scoped_id))
        return _grounded_response(NO_ANSWER, retrieved[:2] or chunks[:2])

    if any(word in query for word in ("property", "premises", "address", "location")):
        property_info = analysis.get("property", {})
        if property_info.get("address") and property_info.get("source_text"):
            answer = f"The property is at {property_info['address']}."
            return _grounded_response(answer, _records_as_evidence([property_info], scoped_id))
        return _grounded_response(NO_ANSWER, retrieved[:2] or chunks[:2])

    if "renewal" in query or "renew" in query:
        renewal_chunks = [chunk for chunk in chunks if re.search(r"renew|extension", chunk["text"], re.IGNORECASE)]
        if renewal_chunks:
            return _grounded_response(" ".join(chunk["text"] for chunk in renewal_chunks), renewal_chunks)
        term_clauses = [clause for clause in analysis.get("clause_items", []) if re.search(r"term|duration", clause["title"], re.IGNORECASE)]
        if term_clauses:
            return _grounded_response(NO_ANSWER, _records_as_evidence(term_clauses, scoped_id, text_key="source_text"))
        relevant = [chunk for chunk in retrieved if re.search(r"term|expiry|expiration", chunk["text"], re.IGNORECASE)]
        if relevant:
            return _grounded_response(NO_ANSWER, relevant[:2])
        return _grounded_response(NO_ANSWER, chunks[:2])

    if not retrieved:
        return _grounded_response(NO_ANSWER, chunks)
    best_score_terms = set(_tokenize(question))
    relevant_sentences = []
    relevant_chunks = []
    for chunk in retrieved:
        for sentence in re.split(r"(?<=[.!?])\s+|\n+", chunk["text"]):
            sentence = sentence.strip()
            if sentence and best_score_terms.intersection(_tokenize(sentence)):
                relevant_sentences.append(sentence)
                relevant_chunks.append(chunk)
    if not relevant_sentences:
        return _grounded_response(NO_ANSWER, retrieved[:2])
    return _grounded_response(" ".join(dict.fromkeys(relevant_sentences)), relevant_chunks)


def build_document_template(form_data: Dict[str, Any]) -> Dict[str, Any]:
    title = str(form_data.get("documentType") or "Rental Agreement")
    fields = {
        "landlord": str(form_data.get("landlordName") or "____________________________"),
        "landlord_address": str(form_data.get("landlordAddress") or "____________________________"),
        "landlord_contact": str(form_data.get("landlordContact") or "____________________________"),
        "tenant": str(form_data.get("tenantName") or "____________________________"),
        "tenant_address": str(form_data.get("tenantAddress") or "____________________________"),
        "tenant_contact": str(form_data.get("tenantContact") or "____________________________"),
        "property_address": str(form_data.get("propertyAddress") or "____________________________"),
        "property_type": str(form_data.get("propertyType") or "____________________________"),
        "start_date": str(form_data.get("startDate") or "____________________________"),
        "end_date": str(form_data.get("endDate") or "____________________________"),
        "amount": str(form_data.get("monthlyRent") or "____________________________"),
        "deposit": str(form_data.get("securityDeposit") or "____________________________"),
        "rent_due_date": str(form_data.get("rentDueDate") or "____________________________"),
        "notice_period": str(form_data.get("noticePeriod") or "____________________________"),
        "additional_terms": str(form_data.get("additionalTerms") or "____________________________"),
    }
    blocks: List[Dict[str, str]] = [{"type": "title", "text": title.upper()}]

    def add(kind: str, content: str) -> None:
        blocks.append({"type": kind, "text": content})

    def add_signatures(first_label: str, second_label: str) -> None:
        add("section_heading", "SIGNATURES")
        add("signature", f"{first_label} SIGNATURE:  ____________________________________")
        add("paragraph", f"Name: {fields['landlord']}")
        add("paragraph", "Date:  ______________________________")
        add("signature", f"{second_label} SIGNATURE:  ____________________________________")
        add("paragraph", f"Name: {fields['tenant']}")
        add("paragraph", "Date:  ______________________________")

    if title == "Rental Agreement":
        add("paragraph", "This Rental Agreement (the \"Agreement\") is made between the following parties:")
        add("party_heading", "LANDLORD")
        add("paragraph", f"Name: {fields['landlord']}")
        add("paragraph", f"Address: {fields['landlord_address']}")
        add("paragraph", f"Contact: {fields['landlord_contact']}")
        add("paragraph", "AND")
        add("party_heading", "TENANT")
        add("paragraph", f"Name: {fields['tenant']}")
        add("paragraph", f"Address: {fields['tenant_address']}")
        add("paragraph", f"Contact: {fields['tenant_contact']}")
        add("paragraph", f"The premises covered by this Agreement are located at {fields['property_address']} and are described as {fields['property_type']}.")
        add("section_heading", "1. TERM")
        add("paragraph", f"The tenancy commences on {fields['start_date']} and ends on {fields['end_date']}, subject to the termination provisions of this Agreement.")
        add("section_heading", "2. RENT")
        add("paragraph", f"The monthly rent is {fields['amount']}, payable on or before {fields['rent_due_date']}.")
        add("section_heading", "3. SECURITY DEPOSIT")
        add("paragraph", f"The Tenant shall pay a security deposit of {fields['deposit']}.")
        add("section_heading", "4. NOTICE AND TERMINATION")
        add("paragraph", f"Either party may terminate this Agreement by giving {fields['notice_period']} written notice to the other party.")
        add("section_heading", "5. ADDITIONAL TERMS")
        add("paragraph", fields["additional_terms"])
        add_signatures("LANDLORD", "TENANT")
    elif title == "Employment Agreement":
        add("paragraph", f"This Employment Agreement is made between {fields['landlord']} (Employer) and {fields['tenant']} (Employee).")
        add("paragraph", f"Employer address and contact: {fields['landlord_address']}; {fields['landlord_contact']}.")
        add("paragraph", f"Employee address and contact: {fields['tenant_address']}; {fields['tenant_contact']}.")
        add("section_heading", "1. POSITION AND LOCATION")
        add("paragraph", f"Position: {fields['property_type']}. Work location: {fields['property_address']}.")
        add("section_heading", "2. EMPLOYMENT TERM")
        add("paragraph", f"Employment begins on {fields['start_date']} and ends on {fields['end_date']} unless ended earlier under this Agreement.")
        add("section_heading", "3. COMPENSATION")
        add("paragraph", f"Compensation: {fields['amount']}. Additional compensation: {fields['deposit']}.")
        add("section_heading", "4. RESPONSIBILITIES AND TERMS")
        add("paragraph", fields["additional_terms"])
        add("section_heading", "5. TERMINATION")
        add("paragraph", f"Either party may terminate this Agreement by providing {fields['notice_period']} written notice.")
        add_signatures("EMPLOYER", "EMPLOYEE")
    elif title == "NDA":
        add("paragraph", f"This Non-Disclosure Agreement is made between {fields['landlord']} (Disclosing Party) and {fields['tenant']} (Receiving Party).")
        add("section_heading", "1. PURPOSE")
        add("paragraph", f"Purpose: {fields['property_type']}.")
        add("section_heading", "2. CONFIDENTIAL INFORMATION")
        add("paragraph", "The Receiving Party shall protect confidential information received from the Disclosing Party and shall not disclose it except as authorized in writing.")
        add("section_heading", "3. TERM")
        add("paragraph", f"This Agreement takes effect on {fields['start_date']} and ends on {fields['end_date']}.")
        add("section_heading", "4. RETURN OR DELETION")
        add("paragraph", fields["additional_terms"])
        add_signatures("DISCLOSING PARTY", "RECEIVING PARTY")
    elif title == "Legal Notice":
        add("paragraph", f"FROM: {fields['landlord']}\nTO: {fields['tenant']}\nADDRESS: {fields['property_address']}\nDATE: {fields['start_date']}")
        add("section_heading", "SUBJECT")
        add("paragraph", fields["property_type"])
        add("section_heading", "NOTICE")
        add("paragraph", fields["additional_terms"])
        add("paragraph", f"Requested response date: {fields['end_date']}. Response period: {fields['notice_period']}.")
        add("signature", f"SENDER SIGNATURE:  ____________________________________\nName: {fields['landlord']}\nDate:  ______________________________")
    elif title == "Service Agreement":
        add("paragraph", f"This Service Agreement is made between {fields['landlord']} (Service Provider) and {fields['tenant']} (Client).")
        add("section_heading", "1. SERVICES")
        add("paragraph", f"Services: {fields['property_type']}. Service location: {fields['property_address']}.")
        add("section_heading", "2. TERM")
        add("paragraph", f"Services begin on {fields['start_date']} and end on {fields['end_date']}.")
        add("section_heading", "3. FEES AND PAYMENT")
        add("paragraph", f"Service fee: {fields['amount']}. Advance payment: {fields['deposit']}. Payment due: {fields['rent_due_date']}.")
        add("section_heading", "4. TERMINATION")
        add("paragraph", f"Either party may terminate this Agreement by providing {fields['notice_period']} written notice.")
        add("section_heading", "5. ADDITIONAL TERMS")
        add("paragraph", fields["additional_terms"])
        add_signatures("SERVICE PROVIDER", "CLIENT")
    else:
        add("paragraph", f"This Agreement is made between {fields['landlord']} and {fields['tenant']}.")
        add("section_heading", "1. PURPOSE AND TERMS")
        add("paragraph", fields["additional_terms"])
        add_signatures("FIRST PARTY", "SECOND PARTY")

    document_text = "\n\n".join(block["text"] for block in blocks)
    return {
        "id": f"draft-{uuid.uuid4().hex}",
        "title": title,
        "document_text": document_text,
        "document_blocks": blocks,
        "disclaimer": "This document is an AI-generated draft for informational and educational purposes only. It is not a substitute for professional legal advice. Please have the document reviewed by a qualified legal professional before signing or using it.",
    }
