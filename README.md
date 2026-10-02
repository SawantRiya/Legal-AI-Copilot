# Legal-AI-Copilot

AI-Powered Agentic Legal Documentation & Intelligence System

## Overview
This project delivers a lightweight legal AI copilot for reviewing legal text, surfacing key clauses, assessing contract risk, and answering pointed legal questions from the supplied document.

## Features
- Clause extraction for common legal provisions such as confidentiality, indemnity, liability, governance, and termination
- Risk scoring and severity classification for legal document review
- Question answering based on the provided contract text
- Simple browser-based interface with JSON API support

## Run locally
```bash
python app.py
```

Then open:
- http://127.0.0.1:8000/

## API example
```bash
curl -X POST http://127.0.0.1:8000/api/analyze \
  -H "Content-Type: application/json" \
  -d '{
    "document": "The parties agree that the governing law is the State of New York.",
    "question": "What is the governing law?"
  }'
```

## Validation
```bash
pytest -q
```
