"""
sections_config.py
Defines the ordered section list for contract review.

To use for a different document:
  - Replace SECTIONS with the new document's sections in order.
  - Keys are the section title as it appears in the document (case-insensitive match).
  - Values are the numeric order used for sorting.

Example for a different contract:
  SECTIONS = {
      "DEFINITIONS": 1.0,
      "SCOPE": 2.0,
      ...
  }
"""

# ─── Active contract section order ────────────────────────────────────────────
# Granite / C3M Power Systems subcontract

SECTIONS = {
    "CONTRACT":                  1.0,
    "SCOPE OF WORK":             2.0,
    "INVESTIGATION":             3.0,
    "EXECUTION & PROGRESS":      4.0,
    "PAYMENT":                   5.0,
    "INSURANCE":                 6.0,
    "INDEMNITY":                 7.0,
    "CHANGES":                   8.0,
    "TRUST FUNDS":               9.0,
    "DELAY":                    10.0,
    "SUSPENSION OR TERMINATION": 11.0,
    "DBE":                      12.0,
    "SAFETY & COMPLIANCE":      13.0,
    "DEFAULT":                  14.0,
    "MECHANICS LIENS":          15.0,
    "BONDS":                    16.0,
    "OTHER CONTRACTS":          17.0,
    "WARRANTY":                 18.0,
    "LABOR CONDITIONS":         19.0,
    "RESPONSIBILITY":           20.0,
    "CLAIMS & DISPUTES":        21.0,
    "LIMITATION OF LIABILITY":  22.0,
    "ARBITRATION":              23.0,
    "INDEPENDENT CONTRACTOR":   24.0,
    "SPECIAL PROVISION":        30.0,
    "ATTACHMENT A.1":           31.0,
    "ATTACHMENT A.2":           32.0,
    "ATTACHMENT A.3":           33.0,
}


def get_section_order(section_title: str) -> float:
    """
    Return the numeric order for a section title.
    Matches case-insensitively and strips leading numbers (e.g. '5.0 PAYMENT' → 'PAYMENT').
    Unknown sections are pushed to the end (999).
    """
    if not section_title:
        return 999.0

    # Strip leading number prefix like "5.0 " or "5.0. "
    import re
    cleaned = re.sub(r"^\d+(\.\d+)*\.?\s*", "", section_title).strip().upper()

    # Try exact match first, then partial match
    if cleaned in SECTIONS:
        return SECTIONS[cleaned]

    for key in SECTIONS:
        if key in cleaned or cleaned in key:
            return SECTIONS[key]

    return 999.0