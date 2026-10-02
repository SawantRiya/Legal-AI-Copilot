from legal_ai import analyze_legal_document, build_document_template
from app import build_pdf_document
from legal_ai import analyze_document, answer_question_from_document, extract_pdf_text


def test_analyze_legal_document_extracts_keywords_and_risk():
    document = """
    Employment Agreement

    The Company shall terminate the employee for cause if the employee breaches confidentiality,
    fails to comply with company policies, or engages in misconduct. The employee shall not use
    confidential information for personal gain. The parties agree that the governing law is the
    State of New York.
    """

    result = analyze_legal_document(document, question="What is the governing law?")

    assert "governing law" in result["summary"].lower()
    assert result["risk_level"] in {"low", "moderate", "high"}
    assert result["risk_score"] >= 0
    assert result["risk_score"] <= 100
    assert any("confidential" in clause.lower() for clause in result["key_clauses"])
    assert "new york" in result["answer"].lower()


def test_analyze_detects_confidentiality_risk():
    document = """
    This agreement permits unrestricted use of confidential information and contains no non-solicitation,
    no indemnity clause, and no clear limitation of liability. There is no termination for cause language.
    """

    result = analyze_legal_document(document)

    assert result["risk_level"] in {"moderate", "high"}
    assert result["risk_score"] > 40
    assert any("confidential" in issue.lower() for issue in result["issues"])


def test_document_type_selection_builds_an_employment_agreement():
    result = build_document_template({
        "documentType": "Employment Agreement",
        "landlordName": "Example Employer",
        "tenantName": "Example Employee",
        "propertyType": "Counsel",
        "monthlyRent": "INR 90000",
    })

    assert result["title"] == "Employment Agreement"
    assert "employer" in result["document_text"].lower()
    assert "employee" in result["document_text"].lower()
    assert "this rental agreement" not in result["document_text"].lower()


def test_generated_document_ids_are_unique():
    first = build_document_template({"documentType": "NDA"})
    second = build_document_template({"documentType": "NDA"})

    assert first["id"] != second["id"]


def test_analysis_extracts_rent_when_pdf_font_loses_currency_symbol(tmp_path):
    pdf_bytes = build_pdf_document("Rental Agreement", [
        "The Tenant shall pay a monthly rent of ₹30,000.",
        "The Tenant shall deposit ₹60,000 as a security deposit.",
    ])
    pdf_path = tmp_path / "rental.pdf"
    pdf_path.write_bytes(pdf_bytes)
    extracted_text = extract_pdf_text(str(pdf_path))["text"]
    analysis = analyze_document("rental.pdf", extracted_text)
    answer = answer_question_from_document("What is the monthly rent?", analysis)

    assert analysis["financial"]["monthly_rent"].endswith("30,000")
    assert analysis["financial"]["security_deposit"].endswith("60,000")
    assert "30,000" in answer["answer"]


def test_retrieval_uses_page_metadata_and_never_returns_other_document_chunks():
    analysis = analyze_document(
        "lease.pdf",
        "Payment is due on the fifth day of each month.\nLandlord: Nia Patel",
        pages=[
            {"page": 1, "text": "Payment is due on the fifth day of each month."},
            {"page": 2, "text": "Landlord: Nia Patel"},
        ],
        document_id="document-a",
    )

    answer = answer_question_from_document("Who is the landlord?", analysis, document_id="document-a")
    other_document = answer_question_from_document("Who is the landlord?", analysis, document_id="document-b")

    assert answer["answer"] == "The landlord is Nia Patel."
    assert answer["source"] == "Page 2"
    assert answer["evidence"] == "Landlord: Nia Patel"
    assert other_document["answer"] == "The current document does not contain enough information to answer this question."
    assert other_document["source"] == "N/A"
    assert other_document["evidence"] == ""


def test_supplied_rental_analysis_exposes_structured_terms():
    source_lines = [
        "This Rental Agreement is made on 1 September 2026 between Rahul Mehta (Landlord) and Priya Sharma (Tenant).",
        "The premises is located at Flat 402, Green Heights, Andheri East, Mumbai, Maharashtra.",
        "The tenancy shall commence on 1 September 2026 and continue for a period of 11 months, ending on 31 July 2027.",
        "Monthly rent shall be INR 30,000, payable on or before the 5th day of each calendar month.",
        "The Tenant shall pay a refundable security deposit of INR 60,000.",
        "The Tenant shall keep the premises in reasonable condition, pay electricity and other usage-based charges, and use the premises only for residential purposes.",
        "Major structural repairs shall remain the responsibility of the Landlord.",
        "Either party may terminate this Agreement by giving 30 days' written notice to the other party.",
        "The Tenant shall not sublet, assign, or transfer the premises without prior written consent from the Landlord.",
        "Signatures. Landlord: Rahul Mehta. Tenant: Priya Sharma. Date: 1 September 2026.",
    ]
    pdf_path = __import__("pathlib").Path("/tmp/supplied-rental-analysis.pdf")
    pdf_path.write_bytes(build_pdf_document("RESIDENTIAL RENTAL AGREEMENT", source_lines))
    extracted = extract_pdf_text(str(pdf_path))
    analysis = analyze_document("Rental_Agreement_Demo.pdf", extracted["text"], pages=extracted["pages"], document_id="supplied-rental")

    assert analysis["parties"][0]["name"] == "Rahul Mehta"
    assert analysis["parties"][0]["role"] == "Landlord"
    assert analysis["parties"][0]["source_page"] == 1
    assert analysis["parties"][1]["name"] == "Priya Sharma"
    assert analysis["dates"]["start_date"] == "1 September 2026"
    assert analysis["dates"]["end_date"] == "31 July 2027"
    assert analysis["dates"]["document_date"] == "1 September 2026"
    assert analysis["financial_terms"][0]["amount"] == "30,000"
    assert analysis["financial_terms"][0]["due_date"] == "on or before the 5th day of each calendar month"
    assert analysis["notice"]["duration"] == "30"
    assert analysis["obligation_items"]
    assert all(item["source_page"] == 1 and item["source_text"] for item in analysis["obligation_items"])


def test_answers_synthesize_without_repeating_source_evidence():
    analysis = analyze_document(
        "Rental_Agreement_Demo.pdf",
        "Monthly rent shall be INR 30,000, payable on or before the 5th day of each calendar month.",
        pages=[{"page": 1, "text": "Monthly rent shall be INR 30,000, payable on or before the 5th day of each calendar month."}],
        document_id="rent-answer",
    )

    answer = answer_question_from_document("What is the monthly rent?", analysis)

    assert "30,000" in answer["answer"]
    assert "5th day" in answer["answer"]
    assert answer["evidence"] == "Monthly rent shall be INR 30,000, payable on or before the 5th day of each calendar month."
    assert answer["answer"].casefold() != answer["evidence"].casefold()


def test_rental_draft_has_professional_structure_and_signature_blocks():
    result = build_document_template({
        "documentType": "Rental Agreement",
        "landlordName": "Rahul Mehta",
        "landlordAddress": "10 Main Street",
        "landlordContact": "555-0101",
        "tenantName": "Priya Sharma",
        "tenantAddress": "12 Side Street",
        "tenantContact": "555-0102",
        "propertyAddress": "Flat 402, Green Heights",
        "propertyType": "Residential Apartment",
        "startDate": "2026-09-01",
        "endDate": "2027-07-31",
        "monthlyRent": "INR 30,000",
        "rentDueDate": "5th of each month",
        "securityDeposit": "INR 60,000",
        "noticePeriod": "30 days",
    })

    assert result["document_blocks"]
    assert result["document_blocks"][0]["type"] == "title"
    assert "LANDLORD" in result["document_text"]
    assert "TENANT" in result["document_text"]
    assert "LANDLORD SIGNATURE" in result["document_text"]
    assert "TENANT SIGNATURE" in result["document_text"]
    assert "AI-generated draft" not in result["document_text"]
