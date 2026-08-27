"""
demand_letter_import.py

Parses a COMPLETED demand letter (.docx, produced by the existing demand
letter generator / LOU-main project) and extracts the handful of numbers
that overlap with arbitration-packet dashboard fields, so an agent who
already has the demand letter for a claim can upload it instead of
retyping property damage, loss-of-use, repair hours/days, and the daily
rate breakdown.

This is intentionally regex-based against the FIXED sentence structure the
demand letter template always produces (see DemandTemplate_NY.docx) rather
than a token-substitution reversal — a completed letter has already had its
[TOKENS] replaced with plain text, so there's nothing left to match on
except the surrounding fixed wording.

Nothing here is authoritative: every extracted value lands in the normal
dashboard field for the agent to review, edit, or clear before generating
the arbitration packet.
"""

from __future__ import annotations

import re
from typing import Any, Dict, IO

from docx import Document

_MONEY = r"([\d,]+\.\d{2})"


def _clean_money(s: str) -> str:
    return s.replace(",", "").replace("$", "").strip()


def extract_from_demand_letter(file_obj: "IO[bytes] | str") -> Dict[str, Any]:
    """Best-effort extraction. Returns only the keys it was confident about
    — never guesses, never raises on a section it can't find."""
    doc = Document(file_obj)
    text = "\n".join(p.text for p in doc.paragraphs)

    result: Dict[str, Any] = {}

    def grab(pattern: str, key: str, group: int = 1, transform=lambda s: s.strip()):
        m = re.search(pattern, text, re.I | re.S)
        if m:
            val = transform(m.group(group))
            if val:
                result[key] = val

    grab(r"Property Damage;?\s*\$?" + _MONEY, "property_damage", transform=_clean_money)
    grab(r"Loss of Use;?\s*\$?" + _MONEY, "loss_of_use_amount", transform=_clean_money)
    grab(r"Towing:?\s*\$?" + _MONEY, "towing", transform=_clean_money)
    grab(r"Vehicle Type:?\s*([A-Za-z][A-Za-z ]*)", "vehicle_type")
    grab(r"Date of Loss:?\s*([\d\-/]+)", "date_of_loss")

    m = re.search(
        r"and\s+(\d+)\s+repair hours.*?"
        r"compensation for\s+(\d+)\s+days.*?"
        r"total repair period\s+(\d+)\s+days",
        text, re.I | re.S,
    )
    if m:
        repair_hours, raw_days, total_days = m.group(1), m.group(2), m.group(3)
        result["repair_hours"] = repair_hours
        result["raw_days"] = raw_days
        result["total_lou_days"] = total_days
        try:
            weekend_days = int(total_days) - int(raw_days) - 1  # 1 = paint-cure day
            if weekend_days >= 0:
                result["applicable_weekend_days"] = str(weekend_days)
        except ValueError:
            pass

    m = re.search(
        r"daily replacement value is\s*\$?" + _MONEY + r".*?"
        r"driver wage\s*\(\$?" + _MONEY + r"/hour X 8 hours = " + _MONEY + r"\),?\s*"
        r"the adjusted daily rate is\s*\$?\s*" + _MONEY,
        text, re.I | re.S,
    )
    if m:
        result["repl_daily"] = _clean_money(m.group(1))
        result["driver_wage_hourly"] = _clean_money(m.group(2))
        result["driver_wage_daily"] = _clean_money(m.group(3))
        result["final_daily_rate"] = _clean_money(m.group(4))

    return result
