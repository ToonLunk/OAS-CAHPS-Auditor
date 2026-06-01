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


# --- Discharge Status and Admit Source classification ---

# Vendor crosswalk: some vendors submit codes in the 81-95 range that map to
# canonical HCAHPS discharge status values.
_DS_CROSSWALK = {
    81: 1,  82: 2,  83: 3,  84: 4,  85: 5,  86: 6,
    87: 21, 88: 70, 89: 2,  90: 62, 91: 63, 92: 64,
    93: 65, 94: 66, 95: 70,
}

# Patients with these canonical DS values should be on the EXCLU tab
_DS_EXCLU = {3, 21, 50, 51, 61, 64}
# Patients with these canonical DS values should be on the INEL tab
_DS_INEL  = {20, 30, 41, 42}
# Patients with these AS values should be on the EXCLU tab
_AS_EXCLU = {8}


def _normalize_ds(ds_raw):
    """
    Parse a raw DS cell value to an integer, applying the vendor crosswalk
    for 81-95 range values. Returns None if blank or non-numeric.
    """
    if ds_raw is None:
        return None
    raw = str(ds_raw).strip()
    if not raw:
        return None
    try:
        num = int(float(raw))
    except (ValueError, TypeError):
        return None
    return _DS_CROSSWALK.get(num, num)


def classify_ds(ds_raw):
    """
    Classify a discharge status value after applying the vendor crosswalk.
    Returns ('exclu', reason), ('inel', reason), or (None, None).
    """
    num = _normalize_ds(ds_raw)
    if num is None:
        return None, None
    if num in _DS_EXCLU:
        return 'exclu', f"DS {num} — patient should be on the EXCLU tab"
    if num in _DS_INEL:
        return 'inel', f"DS {num} — patient should be on the INEL tab"
    return None, None


def classify_as(as_raw):
    """
    Classify an admit source value.
    Returns ('exclu', reason) or (None, None).
    """
    if as_raw is None:
        return None, None
    raw = str(as_raw).strip()
    if not raw:
        return None, None
    try:
        num = int(float(raw))
    except (ValueError, TypeError):
        return None, None
    if num in _AS_EXCLU:
        return 'exclu', f"AS {num} — patient should be on the EXCLU tab"
    return None, None


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

    # Find MRN (or fallback identifier) column in header row
    mrn_col, _ = _find_mrn_col(exclu_sheet)

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


def validate_inel_rows(inel_sheet, show_progress=False):
    """
    Validate HCAHPS INEL tab rows.

    Every non-blank row in the INEL tab should have at least one cell that is
    highlighted (non-white/non-transparent fill) or has red font, indicating
    the reason for ineligibility.

    Returns (inel_count, row_issues) where inel_count is the total number of
    non-blank data rows found.
    """
    count, issues = validate_exclu_rows(inel_sheet, show_progress=show_progress)
    for issue in issues:
        issue['issue_type'] = 'INEL Missing Highlight'
        issue['description'] = 'No highlighted cell or red font - ineligibility reason not indicated'
    return count, issues


def count_dup_d_rows(dup_sheet):
    """
    Count rows in the DUP tab where the DUP column contains 'D'.
    Returns an int count (0 if DUP column not found).
    """
    dup_col = None
    first_row = next(dup_sheet.iter_rows(min_row=1, max_row=1), None)
    if first_row:
        for cell in first_row:
            if cell.value is not None and str(cell.value).strip().lower() == "dup":
                dup_col = cell.column
                break
    if dup_col is None:
        return 0
    count = 0
    for row in dup_sheet.iter_rows(min_row=2, values_only=True):
        if not any(cell is not None for cell in row):
            continue
        val = row[dup_col - 1]
        if val is not None and str(val).strip().upper() == "D":
            count += 1
    return count


def count_frame_patients(frame_sheet):
    """
    Count total and duplicate patients in the HCAHPS FRAME tab.

    FRAME structure:
      Row 1     : header row
      Dense block: one row per patient — duplicates (no random number) followed
                  by eligible patients (with random number in col A or similar)
      Separator : 2+ consecutive blank rows
      Sparse block: duplicate MRNs only (pasted for conditional formatting)

    Returns:
        (total_patients, dup_patients)
        total_patients = non-empty rows in the dense block (after header)
        dup_patients   = non-empty rows in the sparse block
        eligible       = total_patients - dup_patients

    Returns (None, None) if the sheet is empty or cannot be parsed.
    """
    rows = list(frame_sheet.iter_rows(values_only=True))
    if not rows:
        return None, None

    nonempty_counts = [
        sum(1 for c in row if c is not None and str(c).strip() != "")
        for row in rows
    ]

    # Find the blank separator: first pair of consecutive empty rows after header
    separator_start = None
    for i in range(1, len(rows) - 1):
        if nonempty_counts[i] == 0 and nonempty_counts[i + 1] == 0:
            separator_start = i
            break

    if separator_start is None:
        # No blank separator — count all data rows, assume no duplicates
        total_patients = sum(1 for i in range(1, len(rows)) if nonempty_counts[i] > 0)
        return total_patients, 0

    # Dense block: rows 1..separator_start-1 (row 0 is the header)
    total_patients = sum(1 for i in range(1, separator_start) if nonempty_counts[i] > 0)

    # Sparse block: everything after the blank separator
    sparse_start = separator_start
    while sparse_start < len(rows) and nonempty_counts[sparse_start] == 0:
        sparse_start += 1
    dup_patients = sum(1 for i in range(sparse_start, len(rows)) if nonempty_counts[i] > 0)

    return total_patients, dup_patients


# --- Same-day discharge check ---

from datetime import datetime, date as _date
from audit_lib_funcs import _expand_aliases, MRN_ALIASES, UNIQUEID_ALIASES

# Base admit date names - admission-specific only (not general service dates).
# ExpandAliases generates the space/underscore/dash variants automatically.
_ADMIT_DATE_BASE = [
    "admit date", "admission date", "date of admission",
    "admit dt", "adm dt", "admission dt", "adm date",
    "admdt", "visit or admit date", "encounter date",
    "ARADMDT", "ARADMDT-N", "ARADMDT-T",
    "a.date", "a date", "a.dt",
    "admit", "admission",
]
_ADMIT_DATE_ALIASES = _expand_aliases(_ADMIT_DATE_BASE)

# Base discharge date names. No VBA list existed for these since OAS doesn't
# use discharge dates, so this is built from the AR field names and common terms.
# TODO: expand this list as new column names are encountered in the wild.
_DISCHARGE_DATE_BASE = [
    "discharge date", "dis date", "disch date",
    "discharge dt", "dis dt", "disch dt",
    "date of discharge", "patient discharge date",
    "hospital discharge date", "patient hospital discharge date",
    "ARDISDT", "ARDISDT-N", "ARDISDT-T",
    "d.date", "disdt", "discdt",
    "discharge",
]
_DISCHARGE_DATE_ALIASES = _expand_aliases(_DISCHARGE_DATE_BASE)

# Two separate expanded sets so priority can be enforced: MRN aliases are always
# tried first; UNIQUEID_ALIASES is only used as a fallback when no MRN alias
# matches.  Both columns often coexist on side tabs (INEL, EXCLU, FRAME, POP) -
# we must not pick uniqueid when MRN is present, regardless of column order.
_MRN_ONLY_ALIASES_EXPANDED = _expand_aliases(MRN_ALIASES)
_UNIQUEID_ALIASES_EXPANDED  = _expand_aliases(UNIQUEID_ALIASES)
# Combined set kept for cases where presence-detection (not priority) is needed.
_MRN_ALIASES_EXPANDED = _MRN_ONLY_ALIASES_EXPANDED | _UNIQUEID_ALIASES_EXPANDED


def _find_mrn_col(sheet):
    """Locate the patient identifier column with explicit priority.

    Pass 1 – try every MRN alias (the vendor renamed uniqueid -> MRN on this tab).
    Pass 2 – fall back to uniqueid aliases only when no MRN alias was found.
    This prevents picking uniqueid over MRN when both columns exist on a tab.
    """
    col, header = _find_col_by_aliases(sheet, _MRN_ONLY_ALIASES_EXPANDED)
    if col is not None:
        return col, header
    return _find_col_by_aliases(sheet, _UNIQUEID_ALIASES_EXPANDED)


def _normalize_date(val):
    """Coerce an openpyxl cell value to a datetime.date, or None."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, _date):
        return val
    s = str(val).strip()
    if not s:
        return None
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y", "%m-%d-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except (ValueError, TypeError):
            pass
    return None


def _find_col_by_aliases(sheet, alias_set):
    """
    Scan the first row of sheet and return (col_index, matched_header_name)
    for the first cell whose value (lowercased, stripped) is in alias_set.
    Returns (None, None) if no match found.
    """
    first_row = next(sheet.iter_rows(min_row=1, max_row=1), None)
    if not first_row:
        return None, None
    for cell in first_row:
        if cell.value is None:
            continue
        if str(cell.value).strip().lower() in alias_set:
            return cell.column, str(cell.value).strip()
    return None, None


def _col_looks_like_dates(sheet, col_idx, sample_rows=15):
    """
    Spot-check whether a column actually contains date values rather than
    short codes (e.g. admit SOURCE = '1', discharge STATUS = 'AMA').
    Returns True if at least one sampled non-blank cell parses as a date.
    Gives benefit of the doubt when the column is empty or all-blank.
    """
    checked = 0
    for row_num in range(2, min(sheet.max_row + 1, 2 + sample_rows)):
        val = sheet.cell(row_num, col_idx).value
        if val is None:
            continue
        checked += 1
        if _normalize_date(val) is not None:
            return True
    return checked == 0  # empty column - give benefit of the doubt


def find_admit_date_col(frame_sheet):
    """Locate the admit date column in FRAME. Returns (col_index, header_name) or (None, None)."""
    col, header = _find_col_by_aliases(frame_sheet, _ADMIT_DATE_ALIASES)
    if col is not None and not _col_looks_like_dates(frame_sheet, col):
        return None, None
    return col, header


def find_discharge_date_col(frame_sheet):
    """Locate the discharge date column in FRAME. Returns (col_index, header_name) or (None, None)."""
    col, header = _find_col_by_aliases(frame_sheet, _DISCHARGE_DATE_ALIASES)
    if col is not None and not _col_looks_like_dates(frame_sheet, col):
        return None, None
    return col, header


def check_same_day_discharges(wb):
    """
    Scan the FRAME tab for rows where admit date == discharge date.
    These patients are ineligible and should be on the INEL tab.

    Returns (issues, row_issues).
    - row_issues is None if the check could not run (FRAME missing, or a
      required column - admit date, discharge date, or MRN - could not be
      located). The issues list will contain a specific warning for each
      column that was not found.
    - row_issues is an empty list if the check ran cleanly with no problems.
    - row_issues is a populated list if same-day discharges were found.

    # TODO: HCAHPS discharge status aliases - disqualifying discharge status
    # codes for HCAHPS differ from OAS and will need a separate check here.
    """
    issues = []

    if 'FRAME' not in wb.sheetnames:
        return issues, None

    frame_sheet = wb['FRAME']

    admit_col,   admit_header   = find_admit_date_col(frame_sheet)
    disch_col,   disch_header   = find_discharge_date_col(frame_sheet)
    mrn_col,     mrn_header     = _find_mrn_col(frame_sheet)

    missing = []
    if admit_col is None:
        missing.append("admit date")
    if disch_col is None:
        missing.append("discharge date")
    if mrn_col is None:
        missing.append("MRN")

    if missing:
        issues.append(
            f"<strong>WARNING:</strong> Could not locate the following column(s) in FRAME tab: "
            f"{', '.join(missing)} - same-day discharge check skipped"
        )
        return issues, None

    # Narrow types for static analysis (already validated by the missing check above)
    if admit_col is None or disch_col is None or mrn_col is None:
        return issues, None

    # Compare admit and discharge dates row by row within FRAME
    row_issues = []
    for row_num in range(2, frame_sheet.max_row + 1):
        row = list(frame_sheet.iter_rows(
            min_row=row_num, max_row=row_num, values_only=True
        ))[0]
        if not any(c is not None for c in row):
            continue

        mrn_raw    = row[mrn_col   - 1]
        admit_raw  = row[admit_col - 1]
        disch_raw  = row[disch_col - 1]

        if mrn_raw is None or admit_raw is None or disch_raw is None:
            continue

        admit_d = _normalize_date(admit_raw)
        disch_d = _normalize_date(disch_raw)

        if admit_d is None or disch_d is None:
            continue

        if admit_d == disch_d:
            row_issues.append({
                'row': row_num,
                'mrn': str(mrn_raw).strip(),
                'cms': None,
                'issue_type': 'Same-Day Discharge',
                'description': (
                    f"Admit and discharge date are both "
                    f"{disch_d.strftime('%m/%d/%Y')} - "
                    f"patient should be on INEL tab"
                ),
            })

    if row_issues:
        issues.append(
            f"<strong>WARNING:</strong> {len(row_issues)} row(s) in FRAME have matching admit and "
            f"discharge dates and should be on the INEL tab "
            f"('{admit_header}' vs '{disch_header}')"
        )

    return issues, row_issues


# --- FRAME address column aliases (ported from VBA H_Addr1/2, H_City, H_State, H_Zip) ---

_ADDR1_BASE = [
    "mailing address1", "mailing address 1", "patient full address", "address",
    "patient address line 1", "patient address 1", "address1", "address 1",
    "addr1", "patientmailingaddress1", "addr 1", "patient mailing address 1",
    "patient address", "pm_address1", "patient mailing address", "add",
    "pataddr1", "street address", "address_1", "patient address1",
    "patient_street1", "patient mailing address street 1", "mailing address",
    "pt_address_1", "pat_mailing_address_1", "per_addr:per addr street 1",
    "araddr1", "addr:per addr street 1", "address_line_1", "patient address-1",
    "addr", "arpataddr1-t", "arpataddr1", "araddr1-t", "patient - address", "street",
    "street 1", "pt mailing address", "pat home addr line1",
    "pt mailing address 1", "pt. street", "add1", "add2"
]
_ADDR1_ALIASES = _expand_aliases(_ADDR1_BASE)

_ADDR2_BASE = [
    "mailing address2", "mailling address 2", "patient address line 2",
    "patient address 2", "address2", "address 2", "addr2",
    "patientmailingaddress2", "addr 2", "patient mailing address 2",
    "pm_address2", "pataddr2", "address_2", "patient_street2",
    "patient mailing address street 2", "patient address2", "mailing address 2",
    "pt_address_2", "pat_mailing_address_2", "per_addr:per addr street 2",
    "araddr2", "addr:per addr street 2", "address_line_2", "patient address-2",
    "arpataddr2-t", "arpataddr2", "araddr2-t", "street 2", "pat home addr line2",
]
_ADDR2_ALIASES = _expand_aliases(_ADDR2_BASE)

_CITY_BASE = [
    "patient city", "patcity", "mailing city", "city", "town",
    "patientaddresscity", "patient mailing city", "pm_city",
    "patient address city", "pt_city", "pat_address_city",
    "per_addr:per addr city", "arcity", "addr:per addr city",
    "arpatcity-t", "arpatcity", "arcity-t", "patient - city", "pat home addr city",
    "pt mailing city", "pt. city",
]
_CITY_ALIASES = _expand_aliases(_CITY_BASE)

_STATE_BASE = [
    "patst", "patient state", "state", "mailing state", "patientaddressstate",
    "patient mailing state", "pm_state", "patient address state", "patstate",
    "pt_state", "pat_address_st", "per_addr:per addr state", "arstate",
    "addr:per addr state", "st", "arpatstate-t", "pat_address_state",
    "arpatstate", "arstate-t", "patient - state", "pat home addr st",
    "pt mailing state", "pt. state",
]
_STATE_ALIASES = _expand_aliases(_STATE_BASE)

_ZIP_BASE = [
    "patient zip code", "patient zip", "mailing zip code", "zip", "zip code",
    "postal code", "patientaddresszipcode", "patient mailing zip code",
    "zipcode", "patient zipcode", "pm_zipkey", "patzip",
    "patient address zip code", "pt_zip", "pat_address_zip",
    "per_addr:per addr zip key", "arzip", "addr:per addr zip",
    "patinet mailing zip code", "zip5", "arpatzip-n", "arzip-n",
]
_ZIP_ALIASES = _expand_aliases(_ZIP_BASE)


def get_frame_col_map(wb):
    """
    Scan the FRAME tab header row and match all known column types by alias.
    Returns a dict: field_label -> (col_idx, matched_header) or (None, None).
    Returns None if the FRAME tab is missing.
    """
    if 'FRAME' not in wb.sheetnames:
        return None
    frame = wb['FRAME']
    return {
        'MRN':            _find_mrn_col(frame),
        'Admit Date':     find_admit_date_col(frame),
        'Discharge Date': find_discharge_date_col(frame),
        'Address 1':      _find_col_by_aliases(frame, _ADDR1_ALIASES),
        'Address 2':      _find_col_by_aliases(frame, _ADDR2_ALIASES),
        'City':           _find_col_by_aliases(frame, _CITY_ALIASES),
        'State':          _find_col_by_aliases(frame, _STATE_ALIASES),
        'ZIP':            _find_col_by_aliases(frame, _ZIP_ALIASES),
    }


def check_frame_addresses(wb, frame_col_map):
    """
    Run address validation on the FRAME tab using pre-detected column indices.
    FRAME already has INEL/EXCLU filtered out, so all rows are validated.
    Returns (invalid_addresses, noted_addresses).
    """
    from audit_lib_funcs import check_address
    if 'FRAME' not in wb.sheetnames or frame_col_map is None:
        return [], []
    frame = wb['FRAME']
    addr1_col, _ = frame_col_map.get('Address 1', (None, None))
    addr2_col, _ = frame_col_map.get('Address 2', (None, None))
    city_col,  _ = frame_col_map.get('City',      (None, None))
    state_col, _ = frame_col_map.get('State',     (None, None))
    zip_col,   _ = frame_col_map.get('ZIP',       (None, None))
    mrn_col,   _ = frame_col_map.get('MRN',       (None, None))
    return check_address(
        frame, addr1_col, city_col, state_col, zip_col,
        mrn_col=mrn_col, cms_col=None, em_col=None,
        street_address_2_col=addr2_col,
    )


_MONTH_NUMS = {
    'january': 1, 'february': 2, 'march': 3, 'april': 4,
    'may': 5, 'june': 6, 'july': 7, 'august': 8,
    'september': 9, 'october': 10, 'november': 11, 'december': 12,
}


def parse_filename_date_range(name_part, filename_year=None, service_date_range=None):
    """
    Detect a day-range in an HCAHPS filename segment, e.g.:
      ' HCAHPS APRIL 12 - 30'  or  ' HCAHPS APRIL 16-30'
    Returns (start_date, end_date) as datetime.date objects, or (None, None).
    Year comes from filename_year, then service_date_range, then current year.
    """
    import re as _re
    import datetime as _dt
    m = _re.search(
        r'(january|february|march|april|may|june|july|august'
        r'|september|october|november|december)'
        r'\s+(\d{1,2})\s*-\s*(\d{1,2})',
        name_part, _re.IGNORECASE,
    )
    if not m:
        return None, None
    month_num = _MONTH_NUMS[m.group(1).lower()]
    day1 = int(m.group(2))
    day2 = int(m.group(3))
    year = filename_year
    if year is None and service_date_range:
        try:
            year = _dt.datetime.strptime(
                service_date_range.split(' - ')[0].strip(), "%m/%d/%Y"
            ).year
        except Exception:
            pass
    if year is None:
        year = _dt.date.today().year
    try:
        return _dt.date(year, month_num, day1), _dt.date(year, month_num, day2)
    except ValueError:
        return None, None


def check_cms_discharge_date_range(sheet, ddate_col, mrn_col, cms_col, start_date, end_date):
    """
    Validate that all discharge dates in the CMS tab fall within [start_date, end_date].
    Returns (issues, row_issues).
    """
    import datetime as _dt
    row_issues = []
    range_str = (
        f"{start_date.strftime('%m/%d/%Y')} – {end_date.strftime('%m/%d/%Y')}"
    )
    for r, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if not any(cell is not None for cell in row):
            continue
        mrn_val   = row[mrn_col   - 1] if mrn_col   else None
        cms_val   = row[cms_col   - 1] if cms_col   else None
        ddate_val = row[ddate_col - 1]
        if ddate_val is None or str(ddate_val).strip() == '':
            continue
        try:
            if isinstance(ddate_val, _dt.datetime):
                d = ddate_val.date()
            elif isinstance(ddate_val, _dt.date):
                d = ddate_val
            else:
                d = _dt.datetime.strptime(str(ddate_val).strip(), "%m/%d/%Y").date()
        except Exception:
            continue  # invalid format handled elsewhere
        if d < start_date or d > end_date:
            row_issues.append({
                'row': r,
                'mrn': mrn_val,
                'cms': cms_val,
                'issue_type': 'Discharge Date Outside Filename Range',
                'description': (
                    f"Discharge date {d.strftime('%m/%d/%Y')} is outside "
                    f"the filename range ({range_str})"
                ),
            })
    return [], row_issues
