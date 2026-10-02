import json
from http.client import HTTPConnection
from http.server import HTTPServer
from pathlib import Path
from threading import Thread

import pytest

from app import (
    MAX_UPLOAD_SIZE,
    LegalAICopilotHandler,
    build_pdf_document,
    clear_active_document,
    validate_pdf_upload,
)


def test_pdf_upload_validation_accepts_pdf_signature():
    assert validate_pdf_upload("agreement.pdf", "application/pdf", b"%PDF-1.7\nbody") is None


def test_pdf_upload_validation_rejects_non_pdf_extensions_and_mime_types():
    assert validate_pdf_upload("notes.txt", "text/plain", b"%PDF-1.7\nbody") == "invalid"
    assert validate_pdf_upload("photo.pdf", "image/jpeg", b"%PDF-1.7\nbody") == "invalid"


def test_pdf_upload_validation_rejects_invalid_signature():
    assert validate_pdf_upload("fake.pdf", "application/pdf", b"not a PDF") == "invalid"


def test_pdf_upload_validation_rejects_oversized_files():
    data = b"%PDF-1.7\n" + b"x" * MAX_UPLOAD_SIZE
    assert validate_pdf_upload("large.pdf", "application/pdf", data) == "too_large"


@pytest.fixture
def upload_server():
    clear_active_document()
    server = HTTPServer(("127.0.0.1", 0), LegalAICopilotHandler)
    thread = Thread(target=server.serve_forever, daemon=True)
    thread.start()
    yield server.server_address
    server.shutdown()
    thread.join()
    server.server_close()
    clear_active_document()


def post_pdf(address, filename, pdf_bytes):
    boundary = "upload-test-boundary"
    body = (
        f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"{filename}\"\r\n"
        "Content-Type: application/pdf\r\n\r\n"
    ).encode() + pdf_bytes + f"\r\n--{boundary}--\r\n".encode()
    connection = HTTPConnection(*address)
    connection.request(
        "POST",
        "/api/documents/upload",
        body=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
    )
    response = connection.getresponse()
    payload = json.loads(response.read())
    connection.close()
    return response.status, payload


def post_question(address, question):
    connection = HTTPConnection(*address)
    body = json.dumps({"question": question})
    connection.request("POST", "/api/documents/current/chat", body=body, headers={"Content-Type": "application/json"})
    response = connection.getresponse()
    payload = json.loads(response.read())
    connection.close()
    return response.status, payload


def build_sample_lease():
    return build_pdf_document("Residential Lease Agreement", [
        "This Rental Agreement is made on 1 September 2026 between Rahul Mehta (Landlord) and Priya Sharma (Tenant).",
        "1. Property",
        "The premises is a two-bedroom residential apartment located at Flat 402, Green Heights, Andheri East, Mumbai, Maharashtra.",
        "2. Term",
        "The tenancy shall commence on 1 September 2026 and shall continue for a period of 11 months, ending on 31 July 2027.",
        "3. Rent and Security Deposit",
        "Monthly rent shall be INR 30,000, payable on or before the 5th day of each calendar month.",
        "The Tenant shall pay a refundable security deposit of INR 60,000.",
        "4. Maintenance and Permitted Use",
        "The Tenant shall keep the premises in reasonable condition, pay electricity and other usage-based charges, and use the premises only for residential purposes.",
        "Major structural repairs shall remain the responsibility of the Landlord.",
        "5. Notice and Termination",
        "Either party may terminate this Agreement by giving 30 days' written notice to the other party.",
        "The Tenant shall return possession of the premises and keys at the end of the notice period.",
        "The Tenant shall not sublet, assign, or transfer the premises to another person without prior written consent from the Landlord.",
        "Signatures. Landlord: Rahul Mehta. Tenant: Priya Sharma. Date: 1 September 2026.",
    ])


def test_uploaded_lease_answers_are_clause_specific_and_source_grounded(upload_server):
    status, response = post_pdf(upload_server, "residential-lease.pdf", build_sample_lease())
    assert status == 200
    document = response["document"]
    chunks = document["analysis"]["chunks"]
    assert chunks
    analysis = document["analysis"]
    assert "Rahul Mehta" in analysis["raw_text"]
    assert "Priya Sharma" in analysis["raw_text"]
    assert analysis["dates"]["start_date"] == "1 September 2026"
    assert analysis["dates"]["end_date"] == "31 July 2027"
    assert analysis["dates"]["notice_date"] == "30 days"
    assert analysis["dates"]["rent_due_date"] == "5th day of each calendar month"
    assert not any("deposit" in obligation.lower() for obligation in analysis["obligations"]["landlord"])
    assert all(chunk["document_id"] == document["id"] for chunk in chunks)
    assert all(chunk["page"] == 1 and chunk["text"].strip() for chunk in chunks)
    assert all("section" in chunk for chunk in chunks)

    expected = {
        "What is the monthly rent?": (("₹30,000 per month", "5th day"), ("Monthly rent shall be INR 30,000", "5th day of each calendar month")),
        "What is the security deposit?": (("₹60,000",), ("refundable security deposit", "INR 60,000")),
        "What is the agreement period?": (("1 September 2026", "31 July 2027"), ("tenancy shall commence", "ending on 31 July 2027")),
        "What is the notice period?": (("30 days' written notice",), ("terminate this agreement", "giving 30 days' written notice")),
        "Who is the landlord?": (("Rahul Mehta",), ("Rahul Mehta (Landlord)",)),
        "Who is the tenant?": (("Priya Sharma",), ("Priya Sharma (Tenant)",)),
        "What are the tenant's obligations?": (("reasonable condition", "electricity", "residential purposes"), ("reasonable condition", "electricity", "residential purposes")),
        "What are the important dates?": (("Agreement start: 1 September 2026", "Agreement end: 31 July 2027", "Rent due: 5th day of each calendar month", "Notice period: 30 days"), ("commence on 1 September 2026", "ending on 31 July 2027", "payable on or before", "30 days' written notice")),
        "What is the property address?": (("Flat 402, Green Heights", "Andheri East"), ("located at Flat 402", "Andheri East")),
    }
    for question, (answer_text, evidence_text) in expected.items():
        answer_status, answer = post_question(upload_server, question)
        assert answer_status == 200
        assert all(value.lower() in answer["answer"].lower() for value in answer_text), (question, answer["answer"])
        assert answer["source"] == "Page 1"
        assert all(value.lower() in answer["evidence"].lower() for value in evidence_text), question
        assert answer["answer"].casefold() != answer["evidence"].casefold()

    renewal_status, renewal = post_question(upload_server, "What information about renewal is present?")
    assert renewal_status == 200
    assert renewal["answer"] == "The current document does not contain enough information to answer this question."
    assert "period of 11 months" in renewal["evidence"].lower()

    absent_status, absent = post_question(upload_server, "What is the governing law?")
    assert absent_status == 200
    assert absent["answer"] == "The current document does not contain enough information to answer this question."
    assert absent["source"] == "Page 1"
    assert absent["evidence"]
    assert absent["evidence"] != "No relevant evidence found."
    assert "Tenant shall keep the premises" in absent["evidence"]


def test_upload_replaces_document_and_does_not_reuse_old_analysis(upload_server):
    rental_pdf = build_sample_lease()
    employment_pdf = build_pdf_document("Employment Agreement", ["Employer: Example Corp", "Annual salary INR 900,000"])

    rental_status, rental_response = post_pdf(upload_server, "rental.pdf", rental_pdf)
    assert rental_status == 200
    rental_document = rental_response["document"]
    assert rental_document["page_count"] == 1
    rental_path = rental_document["file_path"]

    _, rental_answer = post_question(upload_server, "Who is the landlord?")
    assert "Rahul Mehta" in rental_answer["answer"]

    employment_status, employment_response = post_pdf(upload_server, "employment.pdf", employment_pdf)
    assert employment_status == 200
    employment_document = employment_response["document"]
    assert employment_document["id"] != rental_document["id"]
    assert employment_document["filename"] == "employment.pdf"
    assert not Path(rental_path).exists()

    _, answer = post_question(upload_server, "Who is the landlord?")

    assert "does not contain enough information" in answer["answer"].lower()
    assert "Rahul Mehta" not in answer["answer"]
    assert "Rahul Mehta" not in answer["evidence"]