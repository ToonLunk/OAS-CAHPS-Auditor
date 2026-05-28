# CAHPS Auditor - HCAHPS-specific functions
# Copyright (C) 2026 HST Pathways. All rights reserved.
# Originally developed by Tyler Brock. This copyright notice, including
# authorship credit to Tyler Brock, must be preserved in all copies,
# modifications, and derivative works of this software.
#
# This module contains functions that are specific to HCAHPS auditing:
#   - Required header validation
#   - DRG/APR code loading and ineligibility checks
#   - INEL and EXCLU tab validation
#
# Shared validation utilities (phone, email, address, dates, etc.) live in
# audit_lib_funcs.py and are imported by this module where needed.
import json
import os
import sys


# --- DRG/APR ineligibility rules (loaded from JSON) ---

def _get_drg_apr_config_path():
    """Get the path to drg_apr_codes.json from the installation directory."""
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        return os.path.join(exe_dir, 'drg_apr_codes.json')
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, 'drg_apr_codes.json')


_DRG_APR_LOAD_ERROR = None


def _load_drg_apr_config():
    """Load DRG/APR code configuration from JSON file."""
    global _DRG_APR_LOAD_ERROR
    config_path = _get_drg_apr_config_path()
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
        return {
            'drg_valid_range': config['drg']['valid_range'],
            'drg_ineligible_ranges': config['drg']['ineligible_ranges'],
            'apr_valid_range': config['apr']['valid_range'],
            'apr_ineligible_ranges': config['apr']['ineligible_ranges'],
        }
    except Exception as e:
        _DRG_APR_LOAD_ERROR = f"{config_path}: {e}"
        return {
            'drg_valid_range': [1, 999],
            'drg_ineligible_ranges': [],
            'apr_valid_range': [1, 900],
            'apr_ineligible_ranges': [],
        }


_DRG_APR_CONFIG = _load_drg_apr_config()


def _in_ranges(num, ranges):
    return any(lo <= num <= hi for lo, hi in ranges)


def is_ineligible_drg(drg_raw):
    """
    Determine whether a DRG code is ineligible for HCAHPS.
    Returns (is_ineligible, reason).
    """
    if drg_raw is None or str(drg_raw).strip() == "":
        return False, "blank DRG is OK"
    raw = str(drg_raw).strip()
    if not raw.isdigit():
        return True, "non-numeric DRG code"
    num = int(raw)
    lo, hi = _DRG_APR_CONFIG['drg_valid_range']
    if not (lo <= num <= hi):
        return True, f"outside valid DRG range {lo}-{hi}"
    if _in_ranges(num, _DRG_APR_CONFIG['drg_ineligible_ranges']):
        return True, "in ineligible DRG range for HCAHPS"
    return False, "eligible DRG code"


def is_ineligible_apr(apr_raw):
    """
    Determine whether an APR code is ineligible for HCAHPS.
    Returns (is_ineligible, reason).
    """
    if apr_raw is None or str(apr_raw).strip() == "":
        return False, "blank APR is OK"
    raw = str(apr_raw).strip()
    if not raw.isdigit():
        return True, "non-numeric APR code"
    num = int(raw)
    lo, hi = _DRG_APR_CONFIG['apr_valid_range']
    if not (lo <= num <= hi):
        return True, f"outside valid APR range {lo}-{hi}"
    if _in_ranges(num, _DRG_APR_CONFIG['apr_ineligible_ranges']):
        return True, "in ineligible APR range for HCAHPS"
    return False, "eligible APR code"


def check_req_headers(headers):
    """
    Check that all required HCAHPS headers are present in the CMS tab.
    Returns (mapping, missing_req_headers).
    """
    required_names = [
        "SID",
        "PATIENT NAME",
        "TELEPHONE",
        "D.DATE",
        "AGE",
        "DS",
        "GENDER",
        "UNIT",
        "PHYSICIAN NAME",
        "MRN",
        "DRG",
        "ATT",
        "LAG",
        "ID",
        "FD",
        "LG",
        "EMAIL ADDRESS",
        "CMS INDICATOR",
        "LANGUAGE",
    ]

    mapping = {}
    missing_req_headers = []

    for name in required_names:
        mapping[name] = headers.get(name)
        if mapping[name] is None:
            missing_req_headers.append(name)

    return mapping, missing_req_headers


def validate_exclu_rows(exclu_sheet, show_progress=False):
    """
    Validate HCAHPS EXCLU tab rows.

    Every non-blank row in the EXCLU tab should have at least one cell that is
    highlighted (non-white/non-transparent fill) or has red font, indicating
    the reason for exclusion.

    Returns (exclu_count, row_issues) where exclu_count is the total number of
    non-blank data rows found.
    """
    row_issues = []
    exclu_count = 0

    if exclu_sheet is None:
        return 0, row_issues

    # Find MRN column in header row (fall back to column 1)
    mrn_col = None
    first_row = next(exclu_sheet.iter_rows(min_row=1, max_row=1), None)
    if first_row:
        for cell in first_row:
            if cell.value and str(cell.value).strip().upper() == "MRN":
                mrn_col = cell.column
                break

    def _cell_is_marked(cell):
        """Return True if the cell has a non-white/non-transparent fill or red font."""
        try:
            fill = cell.fill
            if fill and fill.fill_type not in (None, "none"):
                fg = fill.fgColor
                if fg:
                    rgb = getattr(fg, "rgb", None)
                    idx = getattr(fg, "index", None)
                    if isinstance(rgb, str) and rgb not in ("00000000", "FFFFFFFF", "00FFFFFF"):
                        return True
                    if isinstance(idx, str) and idx not in ("00000000", "FFFFFFFF"):
                        return True
        except Exception:
            pass
        try:
            font = cell.font
            if font and font.color:
                rgb = getattr(font.color, "rgb", None)
                if isinstance(rgb, str) and rgb.upper()[-6:] == "FF0000":
                    return True
        except Exception:
            pass
        return False

    max_col = exclu_sheet.max_column

    for row_num in range(2, exclu_sheet.max_row + 1):
        row_cells = list(exclu_sheet.iter_rows(
            min_row=row_num, max_row=row_num, min_col=1, max_col=max_col
        ))[0]

        # Skip entirely blank rows
        if not any(c.value is not None and str(c.value).strip() != "" for c in row_cells):
            continue

        exclu_count += 1

        mrn_val = (
            exclu_sheet.cell(row_num, mrn_col).value if mrn_col
            else (row_cells[0].value if row_cells else None)
        )

        if not any(_cell_is_marked(c) for c in row_cells):
            row_issues.append({
                "row": row_num,
                "mrn": mrn_val,
                "cms": None,
                "issue_type": "EXCLU Missing Highlight",
                "description": "No highlighted cell or red font - exclusion reason not indicated",
            })

    return exclu_count, row_issues
