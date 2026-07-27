"""Extract GST challan amounts from digital GST PDFs."""

from __future__ import annotations

import io
import re
from decimal import Decimal, InvalidOperation

from pypdf import PdfReader


GSTIN_RE = re.compile(r"\bGSTIN\s*:\s*(\d{2}[A-Z0-9]{10,13})\b", re.IGNORECASE)
TOTAL_AMOUNT_RE = re.compile(
    r"\bTotal\s+Amount\s+(?!\(in\s+words\))([\d,]+(?:\.\d{1,2})?)",
    re.IGNORECASE,
)
STATE_RE = re.compile(r"\b([A-Z][A-Za-z ]+?)\s*,\s*\d{6}\b")


def extract_receipt(data: bytes, filename: str) -> dict:
    """Return the GSTIN, total paid amount, and optional state from one PDF.

    Raises ValueError when a file does not contain a GSTIN and total amount.
    """
    try:
        reader = PdfReader(io.BytesIO(data))
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as exc:
        raise ValueError("could not read this PDF") from exc

    if "GSTIN" not in text.upper():
        raise ValueError("does not contain a GSTIN")

    gstins = {match.upper() for match in GSTIN_RE.findall(text)}
    if len(gstins) != 1:
        raise ValueError("does not contain exactly one GSTIN")

    amount_match = TOTAL_AMOUNT_RE.search(text)
    if not amount_match:
        raise ValueError("does not contain a Total Amount")
    try:
        amount = Decimal(amount_match.group(1).replace(",", ""))
    except InvalidOperation as exc:
        raise ValueError("has an invalid Total Amount") from exc

    state_match = STATE_RE.search(text)
    return {
        "filename": filename,
        "gstin": gstins.pop(),
        "amount": int(amount) if amount == amount.to_integral_value() else float(amount),
        "state": state_match.group(1).strip() if state_match else "",
    }
