import re
import datetime
import json
import os
import sys
import csv
from openpyxl.worksheet.worksheet import Worksheet
import phonenumbers


# --- SID Registry lookup ---
def _get_sids_csv_path():
    """Get the path to SIDs.csv from the installation directory."""
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller bundle - use installation directory
        exe_dir = os.path.dirname(sys.executable)
        return os.path.join(exe_dir, 'SIDs.csv')
    else:
        # Running as script
        base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, 'SIDs.csv')


def lookup_sid_client_name(sid_prefix, show_missing_warning=False):
    """Look up client name from SIDs.csv by 3-letter SID code.
    
    Args:
        sid_prefix: 3-letter SID code (e.g., 'ANM', 'AGM')
        show_missing_warning: If True, print warning when SIDs.csv is missing
        
    Returns:
        Client name string if found, None if not found or error
    """
    if not sid_prefix or len(sid_prefix) != 3:
        return None
        
    csv_path = _get_sids_csv_path()
    
    # Check if file exists
    if not os.path.exists(csv_path):
        if show_missing_warning:
            print("\n" + "="*60)
            print("NOTE: SIDs.csv not found")
            print("="*60)
            print("The SID registry file (SIDs.csv) is not present in the")
            print("installation directory. SID validation will be skipped.")
            print("")
            print("To enable SID registry checking, contact your IT department")
            print("or system administrator to obtain the SIDs.csv file.")
            print("="*60 + "\n")
        return None
    
    try:
        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 2 and row[0].strip().upper() == sid_prefix.upper():
                    return row[1].strip()
    except Exception as e:
        # Fail silently if CSV can't be read (but exists)
        pass
    
    return None


# --- CPT ineligibility rules (loaded from JSON)
def _get_cpt_config_path():
    """Get the path to cpt_codes.json from the installation directory."""
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller bundle - use installation directory
        exe_dir = os.path.dirname(sys.executable)
        return os.path.join(exe_dir, 'cpt_codes.json')
    else:
        # Running as script
        base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, 'cpt_codes.json')


def _load_cpt_config():
    """Load CPT code configuration from JSON file."""
    config_path = _get_cpt_config_path()
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
        return {
            'valid_ranges': config.get('valid_ranges', []),
            'invalid_ranges': config.get('invalid_ranges', []),
            'valid_codes': set(str(c).upper() for c in config.get('valid_codes', [])),
            'invalid_codes': set(str(c).upper() for c in config.get('invalid_codes', []))
        }
    except Exception as e:
        print(f"Warning: Could not load CPT config from {config_path}: {e}")
        # Return defaults if file not found
        return {
            'valid_ranges': [[10004, 69990], [93451, 93462], [93566, 93572], [93985, 93986]],
            'invalid_ranges': [],
            'valid_codes': set(),
            'invalid_codes': set()
        }


_CPT_CONFIG = _load_cpt_config()
EXPLICIT_VALID_SET = _CPT_CONFIG['valid_codes']
INVALID_CPT_SET = _CPT_CONFIG['invalid_codes']


def get_hf_text(item):
    if item is None:
        return ""
    txt = getattr(item, "text", None)
    if isinstance(txt, str) and txt:
        return txt
    val = getattr(item, "value", None)
    if isinstance(val, str) and val:
        return val
    return str(item) if item else ""


def clean_hf_text(text: str) -> str:
    if not text:
        return ""
    cleaned = re.sub(r"&[A-Z]", "", text)
    cleaned = cleaned.replace("_x000a_", " ")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def pick_header(sheet):
    return (
        get_hf_text(sheet.oddHeader)
        or get_hf_text(sheet.evenHeader)
        or get_hf_text(sheet.firstHeader)
        or ""
    )


def pick_footer(sheet):
    return (
        get_hf_text(sheet.oddFooter)
        or get_hf_text(sheet.evenFooter)
        or get_hf_text(sheet.firstFooter)
        or ""
    )


def normalize_postal_code(raw):
    if raw is None:
        return None
    s = str(raw).strip()
    if s == "":
        return None
    m = re.search(r"\b(\d{4,5})(?:[-\s]\d{4})?\b", s)
    if not m:
        return s
    digits = m.group(1)
    if len(digits) == 4:
        digits = digits.zfill(5)
    return digits


def check_address(
    sheet,
    street_address_1_col,
    city_col,
    state_col,
    postal_code_col,
    mrn_col=None,
    cms_col=None,
    em_col=None,
):
    from i18naddress import normalize_address, InvalidAddressError

    invalid_addresses = []
    noted_addresses = []

    for row_number, row in enumerate(
        sheet.iter_rows(min_row=2, values_only=True), start=2
    ):
        if not any(cell is not None and str(cell).strip() != "" for cell in row):
            continue

        mrn = row[mrn_col - 1] if mrn_col else ""
        cms = row[cms_col - 1] if cms_col else ""
        em = row[em_col - 1] if em_col else ""
        street_str = str(row[street_address_1_col - 1] or "").strip()
        city_str = str(row[city_col - 1] or "").strip() or None
        state_str = str(row[state_col - 1] or "").strip() or None
        postal_str = normalize_postal_code(row[postal_code_col - 1])

        # Check for missing fields first
        missing = []
        if not street_str:
            missing.append("street")
        if not city_str:
            missing.append("city")
        if not state_str:
            missing.append("state")
        if not postal_str:
            missing.append("zip")

        if missing:
            invalid_addresses.append(
                f"Row: {row_number} - MRN: '{mrn}' - CMS: '{cms}' - E/M: '{em}' - ADDRESS: '{{'street_address': '{street_str}', 'city': '{city_str}', 'country_area': '{state_str}', 'postal_code': '{postal_str}'}}' - REASON: 'Missing: {', '.join(missing)}'"
            )
            continue

        address_data = {
            "country_code": "US",
            "street_address": street_str,
            "city": city_str,
            "country_area": state_str,
            "postal_code": postal_str,
        }

        try:
            normalize_address(address_data)
        except InvalidAddressError as e:
            invalid_addresses.append(
                f"Row: {row_number} - MRN: '{mrn}' - CMS: '{cms}' - E/M: '{em}' - ADDRESS: '{address_data}' - REASON: '{e}'"
            )

        # Check if city, state, or zip are in the street address field
        if city_str and state_str and postal_str:
            try:
                city_pattern = rf"(?i)(?:(?<=^)|(?<=[\s,])){re.escape(city_str)}(?=(?:,\s*{re.escape(state_str)}|\s+{re.escape(state_str)})(?:\b))"
                state_pattern = rf"(?i)(?:(?<=^)|(?<=[\s,])){re.escape(city_str)}(?=(?:,\s*{re.escape(state_str)}|\s+{re.escape(state_str)})(?:\b)).*?(?:,\s*{re.escape(state_str)}|\s+{re.escape(state_str)})"
                zip_pattern = (
                    rf"(?:(?<=^)|(?<=[\s,])){re.escape(postal_str)}(?:(?=$)|(?=[\s,]))"
                )

                issues = []
                if re.search(city_pattern, street_str):
                    issues.append(city_str)
                if re.search(state_pattern, street_str):
                    issues.append(state_str)
                if re.search(zip_pattern, street_str, re.IGNORECASE):
                    issues.append(postal_str)

                if issues:
                    noted_addresses.append(
                        f"Row: {row_number} - MRN: '{mrn}' - CMS: '{cms}' - E/M: '{em}' - ADDRESS: '{street_str}' - REASON(s): '{', '.join(issues)}'"
                    )
            except Exception:
                pass

    return invalid_addresses, noted_addresses


def calc_e_m_total(sheet, cms_col, em_col):
    emails = 0
    mailings = 0
    non_reported = 0
    cms1_count = 0

    for row in sheet.iter_rows(min_row=2, values_only=True):
        cms_val = row[cms_col - 1]  # type: ignore
        em_val = row[em_col - 1]  # type: ignore

        # normalize cms_val to an int when possible
        try:
            cms_num = int(cms_val) if cms_val is not None else None
        except (ValueError, TypeError):
            cms_num = None

        # normalize em_val to uppercase string for reliable comparison
        em_str = str(em_val).strip().upper() if em_val is not None else ""

        if cms_num == 1:
            cms1_count += 1
            if em_str == "E":
                emails += 1
            elif em_str == "M":
                mailings += 1
        elif cms_num == 2:
            non_reported += 1

    total_em = emails + mailings
    return total_em, emails, mailings, non_reported, cms1_count


def classify_cpt(cpt_code: str) -> int:
    """Return expected surgical category for a CPT code (existing logic)."""
    if not cpt_code:
        return 5
    txt = str(cpt_code).strip().lower()
    # Exact text codes
    if txt in ("g0105", "g0121", "g0104"):
        return 1
    if txt == "g0260":
        return 2
    # Numeric ranges
    if txt.isdigit():
        num = int(txt)
        if 40490 <= num <= 49999:
            return 1
        elif 20000 <= num <= 29999:
            return 2
        elif 65091 <= num <= 68999:
            return 3
        elif (
            (10004 <= num <= 19999)
            or (30000 <= num <= 39999)
            or (50000 <= num <= 64999)
            or (68900 <= num <= 69990)
            or (92920 <= num <= 93986)
        ):
            return 4
    return 5


def count_nonempty_rows(sheet):
    """Count rows that actually contain data (ignores blanks/formatting)."""
    count = 0
    for row in sheet.iter_rows(min_row=2, values_only=True):  # skip header
        if any(cell is not None and str(cell).strip() != "" for cell in row):
            count += 1
    return count


def is_blank_row(row) -> bool:
    """Return True if the row contains no meaningful (non-empty) values."""
    for cell in row:
        if cell is not None and str(cell).strip() != "":
            return False
    return True


def parse_dob(raw):
    """Parse DOB from various formats (M/D/YYYY, MM/DD/YYYY, YYYY-MM-DD, etc.).
    Returns (ok, normalized, error_reason). Normalized is always MM/DD/YYYY when ok.
    Flags dates that are invalid, more than 120 years in the past, or in the future.
    """
    if raw is None:
        return False, None, "blank"

    s = str(raw).strip()
    if s.startswith("'"):
        s = s[1:].strip()

    # Try common date formats
    formats = [
        "%m/%d/%Y",  # 3/10/1986 or 03/10/1986
        "%Y-%m-%d",  # 1986-03-10
        "%m-%d-%Y",  # 03-10-1986
    ]

    # Handle datetime objects or strings with time components
    if " " in s:
        s = s.split()[0]  # Take just the date part

    dt = None
    for fmt in formats:
        try:
            dt = datetime.datetime.strptime(s, fmt)
            break
        except ValueError:
            continue

    if dt is None:
        return False, None, "invalid date format"

    # Check if date is valid and within reasonable range
    now = datetime.datetime.now()
    years_ago_120 = now - datetime.timedelta(days=120 * 365.25)

    if dt > now:
        return False, None, "future date"
    if dt < years_ago_120:
        return False, None, "more than 120 years old"

    # Return normalized format MM/DD/YYYY
    return True, dt.strftime("%m/%d/%Y"), None


# CPT eligibility check
def cpt_is_ineligible(cpt_raw) -> tuple[bool, str]:
    """
    Determine whether a CPT code is ineligible.
    Returns (is_ineligible, reason).
    """
    if cpt_raw is None:
        return False, "blank CPT is OK"
    cpt = str(cpt_raw).strip().upper()
    if cpt == "":
        return False, "blank CPT is OK"

    # Exact explicit valid codes are always valid
    if cpt in EXPLICIT_VALID_SET:
        return False, "explicitly valid"

    # Exact invalid list
    if cpt in INVALID_CPT_SET:
        return True, "explicitly invalid list"

    # If purely numeric, check valid ranges
    if cpt.isdigit():
        num = int(cpt)
        
        # Check invalid ranges first
        for range_pair in _CPT_CONFIG['invalid_ranges']:
            if range_pair[0] <= num <= range_pair[1]:
                return True, "numeric in invalid range"
        
        # Check valid ranges
        for range_pair in _CPT_CONFIG['valid_ranges']:
            if range_pair[0] <= num <= range_pair[1]:
                return False, "numeric in valid range"
        
        return True, "outside valid ranges"

    # Non-numeric codes that are not in EXPLICIT_VALID_SET are ineligible
    return True, "not explicitly valid, and not in ranges"


def find_frame_inel_count(
    frame_sheet: Worksheet,
    top_nonempty_threshold: int = 3,
    min_block_rows: int = 3,
    max_blank_within_block: int = 1,
) -> int:
    """
    Locate the lower sparse block and count non-empty values in column B for that block.
    Returns integer count.
    """
    rows = list(frame_sheet.iter_rows(values_only=True))
    if not rows:
        return 0

    # count non-empty cells per row
    nonempty_counts = [
        sum(1 for c in row if c is not None and str(c).strip() != "") for row in rows
    ]

    # find last dense row
    last_dense_index = -1
    for i, cnt in enumerate(nonempty_counts):
        if cnt >= top_nonempty_threshold:
            last_dense_index = i

    # candidate start of sparse region
    start_idx = last_dense_index + 1
    if start_idx >= len(rows):
        return 0

    # accumulate a sparse run starting at start_idx allowing a small number of blanks inside
    sparse_run = 0
    blanks_in_run = 0
    i = start_idx
    while i < len(rows):
        cnt = nonempty_counts[i]
        if cnt <= 2:
            sparse_run += 1
            i += 1  # ✅ Add this line
        else:
            if sparse_run >= min_block_rows:
                break
            sparse_run = 0
            start_idx = i + 1
            i += 1

    if sparse_run < min_block_rows:
        # fallback: scan whole sheet for any run of rows with <=2 non-empty values
        start_idx = None
        for i in range(len(rows)):
            if nonempty_counts[i] <= 2 and nonempty_counts[i] != 0:
                run = 1
                blanks = 0
                for j in range(i + 1, len(rows)):
                    if nonempty_counts[j] == 0:
                        blanks += 1
                        if blanks > max_blank_within_block:
                            break
                    elif nonempty_counts[j] <= 2:
                        run += 1
                    else:
                        break
                if run >= min_block_rows:
                    start_idx = i
                    sparse_run = run
                    break
        if start_idx is None:
            return 0

    end_idx = start_idx + sparse_run  # exclusive

    # Count non-empty values in column B (index 1)
    pt_id_count = 0
    for r in range(start_idx, end_idx):
        row = rows[r]
        # ensure row has at least two columns
        if len(row) >= 2:
            val = row[1]
            if val is not None and str(val).strip() != "":
                pt_id_count += 1

    return pt_id_count


# function to check for required headers
def check_req_headers(headers):
    required_names = [
        "SID",
        "PATIENT NAME",
        "ADDRESS1",
        "CITY",
        "STATE",
        "ZIP",
        "TELEPHONE",
        "SERVICE DATE",
        "GENDER",
        "AGE",
        "PROVIDER NAME",
        "MRN",
        "P.TYPE",
        "SURGICAL CATEGORY",
        "ATT",
        "LAG",
        "ID",
        "FD",
        "LG",
        "E/M",
        "EMAIL ADDRESS",
        "CMS INDICATOR",
        "SURVEY LANGUAGE",
    ]

    mapping = {}
    missing_req_headers = []

    for name in required_names:
        mapping[name] = headers.get(name)
        if mapping[name] is None:
            missing_req_headers.append(name)

    # Return mapping and list of missing headers without raising exception
    return mapping, missing_req_headers


def validate_sid_sequence(sheet, sid_col, cms_col, header_sid=None):
    """
    Validate SID sequence for proper formatting, uniqueness, and numerical order.
    Only validates rows where CMS INDICATOR = 1.
    Returns (issues, row_issues) lists.
    
    Args:
        sheet: The worksheet to validate
        sid_col: Column index for SID (1-based), or None if column missing
        cms_col: Column index for CMS INDICATOR (1-based), or None if column missing
        header_sid: The SID from the header (should be first SID - 1)
    """
    issues = []
    row_issues = []
    
    # Return empty results if required columns are missing
    if sid_col is None or cms_col is None:
        return issues, row_issues
    
    sid_pattern = re.compile(r'^([A-Z]{3})(\d+)$')
    
    sids_found = []
    expected_prefix = None
    expected_start_num = None
    cms1_rows_processed = 0
    first_sid_encountered = False
    
    if header_sid:
        header_match = sid_pattern.match(str(header_sid).strip().upper())
        if header_match:
            expected_prefix = header_match.group(1)
            expected_start_num = int(header_match.group(2)) + 1
        else:
            issues.append(f"Header SID '{header_sid}' does not match expected format (3 letters + numbers)")
    
    row_num = 2
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not any(cell for cell in row):
            break
            
        cms_value = row[cms_col - 1] if cms_col <= len(row) else None
        
        try:
            cms_int = int(float(str(cms_value).strip())) if cms_value is not None and str(cms_value).strip() != "" else None
        except (ValueError, TypeError):
            cms_int = None
        
        if cms_int != 1:
            row_num += 1
            continue
            
        cms1_rows_processed += 1
        sid_value = row[sid_col - 1] if sid_col <= len(row) else None
        mrn_value = row[0] if len(row) > 0 else None
        
        if sid_value is None or str(sid_value).strip() == "":
            row_issues.append({
                'row': row_num,
                'mrn': mrn_value,
                'cms': cms_value,
                'issue_type': 'SID Missing',
                'description': f"Row {row_num}: SID is missing or empty (CMS=1)"
            })
            row_num += 1
            continue
            
        sid_str = str(sid_value).strip().upper()
        
        match = sid_pattern.match(sid_str)
        if not match:
            row_issues.append({
                'row': row_num,
                'mrn': mrn_value,
                'cms': cms_value,
                'issue_type': 'SID Format',
                'description': f"Row {row_num}: SID '{sid_str}' does not match format (3 letters + numbers)"
            })
            row_num += 1
            continue
            
        prefix = match.group(1)
        number = int(match.group(2))
        if not first_sid_encountered:
            first_sid_encountered = True
            if expected_prefix is None:
                expected_prefix = prefix
            if expected_start_num is None:
                expected_start_num = number
                expected_start_num = number
        
        if prefix != expected_prefix:
            row_issues.append({
                'row': row_num,
                'mrn': mrn_value,
                'cms': cms_value,
                'issue_type': 'SID Prefix',
                'description': f"Row {row_num}: SID prefix '{prefix}' does not match expected '{expected_prefix}'"
            })
        
        if sid_str in sids_found:
            row_issues.append({
                'row': row_num,
                'mrn': mrn_value,
                'cms': cms_value,
                'issue_type': 'SID Duplicate',
                'description': f"Row {row_num}: Duplicate SID '{sid_str}'"
            })
            sids_found.append(sid_str)
            
            if expected_start_num is not None:
                expected_num = expected_start_num + (cms1_rows_processed - 1)
                if number != expected_num:
                    row_issues.append({
                        'row': row_num,
                        'mrn': mrn_value,
                        'cms': cms_value,
                        'issue_type': 'SID Sequence',
                        'description': f"Row {row_num}: Expected SID '{expected_prefix}{expected_num:05d}', found '{sid_str}'"
                    })
            
            row_num += 1
        
        if row_issues:
            issues.append(f"Found {len(row_issues)} SID validation issues")
        
        return issues, row_issues


def validate_inel_repeat_rows(inel_sheet):
    """
    Validate INEL tab REPEAT entries.
    
    For rows marked as REPEAT (duplicates):
    - All cells in the row should have red font (RGB 255, 0, 0)
    - "REPEAT" should appear in the rightmost column
    - The "REPEAT" cell should have yellow background fill and bold red font
    - No other cells should have highlighting (yellow background)
    
    For rows with no cell-level highlighting:
    - They MUST have "REPEAT" marker, otherwise there's no indication why they're in INEL
    
    Returns (issues, row_issues) lists.
    """
    issues = []
    row_issues = []
    
    if inel_sheet is None:
        return issues, row_issues
    
    # Get the maximum column used in the sheet
    max_col = inel_sheet.max_column
    
    for row_num in range(2, inel_sheet.max_row + 1):
        row = list(inel_sheet[row_num])
        
        # Skip completely empty rows
        if not any(cell.value is not None and str(cell.value).strip() != "" for cell in row):
            continue
        
        # Check if "REPEAT" exists in the rightmost column
        has_repeat = False
        repeat_cell = None
        rightmost_cell = inel_sheet.cell(row_num, max_col)
        
        if rightmost_cell.value and str(rightmost_cell.value).strip().upper() == "REPEAT":
            has_repeat = True
            repeat_cell = rightmost_cell
        
        # Check for yellow highlighting (background fill) in non-REPEAT cells
        cells_with_yellow_bg = []
        cells_with_red_font = []
        
        for col_num, cell in enumerate(row[:max_col-1], start=1):  # Exclude rightmost column
            if cell.value is not None and str(cell.value).strip() != "":
                # Check for yellow background fill
                if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb:
                    rgb = cell.fill.fgColor.rgb
                    # Yellow is typically FFFFFF00 or variations
                    if isinstance(rgb, str) and len(rgb) >= 6:
                        # Extract RGB values (ignoring alpha channel if present)
                        rgb_str = rgb[-6:] if len(rgb) == 8 else rgb
                        if rgb_str.upper() in ['FFFF00', 'FFFFE0', 'FFFFCC']:  # Common yellow shades
                            cells_with_yellow_bg.append((row_num, col_num))
                
                # Check for red font
                if cell.font and cell.font.color and cell.font.color.rgb:
                    rgb = cell.font.color.rgb
                    if isinstance(rgb, str) and len(rgb) >= 6:
                        rgb_str = rgb[-6:] if len(rgb) == 8 else rgb
                        if rgb_str.upper() == 'FF0000':  # Red font
                            cells_with_red_font.append((row_num, col_num))
        
        # Validate REPEAT rows
        if has_repeat and repeat_cell is not None:
            # Check REPEAT cell formatting
            repeat_font_ok = False
            repeat_bg_ok = False
            repeat_bold_ok = False
            
            if repeat_cell.font is not None:
                if repeat_cell.font.color and repeat_cell.font.color.rgb:
                    rgb = repeat_cell.font.color.rgb
                    if isinstance(rgb, str) and len(rgb) >= 6:
                        rgb_str = rgb[-6:] if len(rgb) == 8 else rgb
                        if rgb_str.upper() == 'FF0000':
                            repeat_font_ok = True
                if repeat_cell.font.bold:
                    repeat_bold_ok = True
            
            if repeat_cell.fill is not None and repeat_cell.fill.fgColor and repeat_cell.fill.fgColor.rgb:
                rgb = repeat_cell.fill.fgColor.rgb
                if isinstance(rgb, str) and len(rgb) >= 6:
                    rgb_str = rgb[-6:] if len(rgb) == 8 else rgb
                    if rgb_str.upper() in ['FFFF00', 'FFFFE0', 'FFFFCC']:
                        repeat_bg_ok = True
            
            # Check if there are other highlighted cells (conflicting indicators)
            if cells_with_yellow_bg:
                row_issues.append({
                    'row': row_num,
                    'issue_type': 'INEL REPEAT Conflict',
                    'description': f"Row {row_num}: Has 'REPEAT' marker but also has {len(cells_with_yellow_bg)} other highlighted cell(s) - conflicting INEL reasons"
                })
                issues.append(f"INEL Row {row_num}: REPEAT marker conflicts with other cell highlighting")
            
            # Check if all cells have red font
            expected_red_cells = len([c for c in row[:max_col-1] if c.value is not None and str(c.value).strip() != ""])
            actual_red_cells = len(cells_with_red_font)
            
            if actual_red_cells < expected_red_cells:
                row_issues.append({
                    'row': row_num,
                    'issue_type': 'INEL REPEAT Formatting',
                    'description': f"Row {row_num}: REPEAT row should have red font on ALL cells ({actual_red_cells}/{expected_red_cells} cells have red font)"
                })
                issues.append(f"INEL Row {row_num}: REPEAT row doesn't have red font on all cells")
            
            # Check REPEAT cell formatting
            formatting_issues = []
            if not repeat_font_ok:
                formatting_issues.append("red font")
            if not repeat_bold_ok:
                formatting_issues.append("bold")
            if not repeat_bg_ok:
                formatting_issues.append("yellow background")
            
            if formatting_issues:
                row_issues.append({
                    'row': row_num,
                    'issue_type': 'INEL REPEAT Cell Format',
                    'description': f"Row {row_num}: REPEAT cell missing {', '.join(formatting_issues)}"
                })
                issues.append(f"INEL Row {row_num}: REPEAT cell missing {', '.join(formatting_issues)}")
        
        # Check rows with no highlighting - they should have REPEAT
        elif not cells_with_yellow_bg:
            # No REPEAT and no highlighted cells = no indication of INEL reason
            row_issues.append({
                'row': row_num,
                'issue_type': 'INEL Missing Reason',
                'description': f"Row {row_num}: No highlighted cells and no REPEAT marker - no indication of why row is in INEL"
            })
            issues.append(f"INEL Row {row_num}: Missing INEL reason indicator (no highlighting or REPEAT marker)")
    
    return issues, row_issues
        
        
# --- Cross-tab consistency checking ---

# MRN and Email alias mappings (from VBA script)
MRN_ALIASES = [
    "chart id",
    "patid",
    "medical account number",
    "patient account number",
    "patient acct no",
    "medical record number",
    "mrn",
    "patient id",
    "patient mrn",
    "medicalrecordnumber",
    "medrec",
    "md rc",
    "acct#",
    "patient account #",
    "account number",
    "patient chart number",
    "acctnum",
    "mrnum",
    "patientid",
    "mrno",
    "per nbr",
    "pt.id",
    "chart number",
    "pt account #",
    "mr#",
    "patient_number",
    "pat_med_rec",
    "person mrn",
    "armrnum",
]

EMAIL_ALIASES = [
    "e-mail address",
    "emailaddress",
    "email",
    "email address",
    "patient email",
    "e-mail",
    "patientemailaddress",
    "patient e-mail",
    "pm_email",
    "patient email address",
    "patemail",
    "email addr",
    "pt. email id",
    "pt email",
    "patmail",
    "pt_email_address",
    "per_email:per addr street 1",
    "arpatmail",
]


def find_column_by_aliases(sheet, aliases):
    """
    Find a column in the sheet by checking against a list of aliases.
    Returns the 1-based column index if found, None otherwise.
    Handles sheets with rows spaced apart.
    """
    # Check first few rows for headers (in case of spacing)
    for header_row_idx in range(1, 6):  # Check first 5 rows
        try:
            row = list(
                sheet.iter_rows(
                    min_row=header_row_idx, max_row=header_row_idx, values_only=True
                )
            )[0]
            for col_idx, cell_value in enumerate(row, start=1):
                if cell_value:
                    cell_str = str(cell_value).strip().lower()
                    # Check against all aliases
                    for alias in aliases:
                        if cell_str == alias.lower():
                            return col_idx, header_row_idx
        except (IndexError, AttributeError):
            continue
    return None, None


def normalize_email(email_val):
    """Normalize email for comparison (lowercase, stripped)."""
    if email_val is None:
        return None
    email_str = str(email_val).strip().lower()
    if email_str == "" or email_str == "none":
        return None
    return email_str


def check_pop_upload_email_consistency(
    wb, upload_sheet, mrn_col_upload, email_col_upload
):
    """
    Check that emails in UPLOAD tab match those in POP tab for the same MRN.
    Returns list of mismatches: [(upload_row, mrn, upload_email, pop_email), ...]
    """
    mismatches = []

    # Check if POP tab exists
    if "POP" not in wb.sheetnames:
        return mismatches  # Can't check without POP tab

    pop_sheet = wb["POP"]

    # Find MRN and Email columns in POP using aliases
    mrn_col_pop, mrn_header_row = find_column_by_aliases(pop_sheet, MRN_ALIASES)
    email_col_pop, email_header_row = find_column_by_aliases(pop_sheet, EMAIL_ALIASES)

    if mrn_col_pop is None:
        return [("N/A", "N/A", "N/A", "Could not locate MRN column in POP tab")]

    if email_col_pop is None:
        return [("N/A", "N/A", "N/A", "Could not locate Email column in POP tab")]

    # Build a dictionary of MRN -> Email from POP tab
    # Use the header row with the most columns as the reference
    pop_data_start_row = max(mrn_header_row or 1, email_header_row or 1) + 1

    pop_mrn_to_email = {}
    for row in pop_sheet.iter_rows(min_row=pop_data_start_row, values_only=True):
        # Skip completely empty rows
        if not any(cell is not None and str(cell).strip() != "" for cell in row):
            continue

        # Get MRN and Email from this row
        try:
            mrn_val = row[mrn_col_pop - 1] if mrn_col_pop <= len(row) else None
            email_val = row[email_col_pop - 1] if email_col_pop <= len(row) else None

            if mrn_val:
                mrn_str = str(mrn_val).strip()
                if mrn_str:
                    # Store normalized email
                    pop_mrn_to_email[mrn_str] = normalize_email(email_val)
        except (IndexError, AttributeError):
            continue

    # Now compare UPLOAD tab against POP data
    for upload_row_idx, row in enumerate(
        upload_sheet.iter_rows(min_row=2, values_only=True), start=2
    ):
        # Skip empty rows
        if not any(cell is not None and str(cell).strip() != "" for cell in row):
            continue

        try:
            upload_mrn = row[mrn_col_upload - 1] if mrn_col_upload <= len(row) else None
            upload_email = (
                row[email_col_upload - 1] if email_col_upload <= len(row) else None
            )

            if not upload_mrn:
                continue

            mrn_str = str(upload_mrn).strip()
            upload_email_norm = normalize_email(upload_email)

            # Check if this MRN exists in POP
            if mrn_str in pop_mrn_to_email:
                pop_email_norm = pop_mrn_to_email[mrn_str]

                # Compare emails (only flag if both exist and differ)
                if upload_email_norm and pop_email_norm:
                    if upload_email_norm != pop_email_norm:
                        mismatches.append(
                            (
                                upload_row_idx,
                                mrn_str,
                                upload_email or "",
                                pop_mrn_to_email[mrn_str] or "",
                            )
                        )
        except (IndexError, AttributeError):
            continue

    return mismatches


def column_validations(sheet, headers, mrn_col, cms_col, em_col, issues, row_issues):
    """
    Perform data quality validation checks on OASCAPHS sheet columns.
    Returns updated issues and row_issues lists.
    """
    from collections import defaultdict

    svc_col = headers.get("SERVICE DATE")
    age_col = headers.get("AGE")
    email_col = headers.get("EMAIL ADDRESS")
    lang_col = headers.get("SURVEY LANGUAGE")
    tel_col = headers.get("TELEPHONE")
    dob_col = headers.get("DATE OF BIRTH")
    name_col = headers.get("PATIENT NAME")
    gender_col = headers.get("GENDER")

    # Track service dates to check they're all in the same month
    service_dates = []
    # Track MRNs to check for duplicates
    mrn_tracker = defaultdict(list)

    for r, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if is_blank_row(row):
            continue

        mrn_val = row[mrn_col - 1] if mrn_col else None
        cms_val = row[cms_col - 1] if cms_col else None
        em_val = row[em_col - 1] if em_col else None

        # Track MRN for duplicate check
        if mrn_val:
            mrn_tracker[mrn_val].append(r)

        # GENDER - must be M, F, 0, 1, or 2
        if gender_col:
            gender_val = row[gender_col - 1]
            valid_genders = ["M", "F", "0", "1", "2"]
            gender_str = str(gender_val).strip().upper() if gender_val else ""
            if not gender_str or gender_str not in valid_genders:
                row_issues.append(
                    {
                        "row": r,
                        "mrn": mrn_val,
                        "cms": cms_val,
                        "issue_type": "Invalid Gender",
                        "description": f"Gender '{gender_val}' not in {valid_genders}",
                    }
                )
                issues.append(f"OASCAPHS Row {r}: Invalid gender '{gender_val}'")

        # SERVICE DATE - validate format and collect all dates for month validation
        if svc_col:
            svc_val = row[svc_col - 1]
            if svc_val:
                # Convert to string for validation
                if isinstance(svc_val, datetime.datetime):
                    svc_str = svc_val.strftime("%m/%d/%Y")
                    service_dates.append((r, mrn_val, svc_val))
                else:
                    svc_str = str(svc_val).strip()

                # Check format MM/DD/YYYY
                if not re.match(
                    r"^(0[1-9]|1[0-2])/(0[1-9]|[12][0-9]|3[01])/\d{4}$", svc_str
                ):
                    row_issues.append(
                        {
                            "row": r,
                            "mrn": mrn_val,
                            "cms": cms_val,
                            "issue_type": "Invalid Service Date Format",
                            "description": f"Service Date '{svc_str}' must be MM/DD/YYYY format",
                        }
                    )
                    issues.append(
                        f"OASCAPHS Row {r}: Service Date '{svc_str}' must be MM/DD/YYYY format"
                    )
                    continue

                # Check if date is in the future
                try:
                    svc_date = datetime.datetime.strptime(svc_str, "%m/%d/%Y")
                    if svc_date > datetime.datetime.now():
                        row_issues.append(
                            {
                                "row": r,
                                "mrn": mrn_val,
                                "cms": cms_val,
                                "issue_type": "Service Date In Future",
                                "description": f"Service Date '{svc_str}' is in the future",
                            }
                        )
                        issues.append(
                            f"OASCAPHS Row {r}: Service Date '{svc_str}' is in the future"
                        )
                    else:
                        # Only add valid dates for month validation
                        service_dates.append((r, mrn_val, svc_date))
                except ValueError:
                    row_issues.append(
                        {
                            "row": r,
                            "mrn": mrn_val,
                            "cms": cms_val,
                            "issue_type": "Invalid Service Date",
                            "description": f"Service Date '{svc_str}' is not a valid date",
                        }
                    )
                    issues.append(
                        f"OASCAPHS Row {r}: Service Date '{svc_str}' is not a valid date"
                    )

        # AGE - must be 18 or older (only matters when CMS=1)
        if age_col:
            age_val = row[age_col - 1]
            try:
                age_int = int(float(str(age_val))) if age_val is not None else None
                cms_int = (
                    int(float(str(cms_val)))
                    if cms_val is not None and str(cms_val).strip()
                    else None
                )

                if age_int is not None and age_int < 18 and cms_int == 1:
                    row_issues.append(
                        {
                            "row": r,
                            "mrn": mrn_val,
                            "cms": cms_val,
                            "issue_type": "Age Too Young",
                            "description": f"Age {age_int} is below 18 (CMS=1)",
                        }
                    )
                    issues.append(
                        f"OASCAPHS Row {r}: Age {age_int} is below 18 (CMS=1)"
                    )
            except (ValueError, TypeError):
                pass

        # make sure date of birth is valid (day, month, and year are present and not in the future). it should look exactly like this: 01/01/2025, for example
        if dob_col:
            dob_val = row[dob_col - 1]
            if dob_val:
                ok, normalized, err = parse_dob(dob_val)
                if not ok:
                    issue_type = "DOB In Future" if err == "future" else "Invalid DOB"
                    row_issues.append(
                        {
                            "row": r,
                            "mrn": mrn_val,
                            "cms": cms_val,
                            "issue_type": issue_type,
                            "description": f"DOB '{dob_val}' error: {err}",
                        }
                    )
                    issues.append(f"OASCAPHS Row {r}: DOB '{dob_val}' error: {err}")

        # EMAIL ADDRESS - validate format when present
        if email_col:
            email_val = row[email_col - 1]
            if email_val and str(email_val).strip():
                email_str = str(email_val).strip()
                # Basic email regex validation
                if not re.match(
                    r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", email_str
                ):
                    row_issues.append(
                        {
                            "row": r,
                            "mrn": mrn_val,
                            "cms": cms_val,
                            "issue_type": "Invalid Email Format",
                            "description": f"Email '{email_str}' has invalid format",
                        }
                    )
                    issues.append(
                        f"OASCAPHS Row {r}: Invalid email format '{email_str}'"
                    )

        # SURVEY LANGUAGE - must be en, es, ko, zh, or m (lowercase)
        if lang_col:
            lang_val = row[lang_col - 1]
            valid_langs = ["en", "es", "ko", "zh", "m"]
            lang_str = str(lang_val).strip() if lang_val else ""
            if not lang_str or lang_str not in valid_langs:
                row_issues.append(
                    {
                        "row": r,
                        "mrn": mrn_val,
                        "cms": cms_val,
                        "issue_type": "Invalid Language Code",
                        "description": f"Language '{lang_str}' not in {valid_langs}",
                    }
                )
                issues.append(f"OASCAPHS Row {r}: Invalid language code '{lang_str}'")

        # E/M and CMS INDICATOR logic
        # - If CMS=1, E/M must be 'E' or 'M'
        # - If CMS=2, E/M should NOT be 'E' or 'M'
        if cms_col and em_col:
            try:
                cms_int = (
                    int(float(str(cms_val)))
                    if cms_val is not None and str(cms_val).strip()
                    else None
                )
                em_str = str(em_val).strip().upper() if em_val else ""

                if cms_int == 1:
                    if em_str not in ["E", "M"]:
                        row_issues.append(
                            {
                                "row": r,
                                "mrn": mrn_val,
                                "cms": cms_val,
                                "issue_type": "Missing E/M for CMS=1",
                                "description": f"CMS=1 but E/M is '{em_val}' (expected 'E' or 'M')",
                            }
                        )
                        issues.append(f"OASCAPHS Row {r}: CMS=1 but E/M is '{em_val}'")
                elif cms_int == 2:
                    if em_str in ["E", "M"]:
                        row_issues.append(
                            {
                                "row": r,
                                "mrn": mrn_val,
                                "cms": cms_val,
                                "issue_type": "Unexpected E/M for CMS=2",
                                "description": f"CMS=2 but E/M is '{em_val}' (should be blank)",
                            }
                        )
                        issues.append(
                            f"OASCAPHS Row {r}: CMS=2 but E/M has value '{em_val}'"
                        )
            except (ValueError, TypeError):
                pass

    # Check all SERVICE DATEs are in the same month
    if service_dates:
        # Get month/year from first date
        first_date = service_dates[0][2]
        expected_month = first_date.month
        expected_year = first_date.year

        for r, mrn_val, svc_date in service_dates:
            if svc_date.month != expected_month or svc_date.year != expected_year:
                row_issues.append(
                    {
                        "row": r,
                        "mrn": mrn_val,
                        "cms": None,
                        "issue_type": "Service Date Wrong Month",
                        "description": f"Date {svc_date.strftime('%Y-%m-%d')} not in {expected_year}-{expected_month:02d}",
                    }
                )
                issues.append(
                    f"OASCAPHS Row {r}: Service date {svc_date.strftime('%Y-%m-%d')} not in expected month {expected_year}-{expected_month:02d}"
                )

    # Check for duplicate MRNs
    for mrn, rows in mrn_tracker.items():
        if len(rows) > 1:
            rows_str = ", ".join(str(r) for r in rows)
            for r in rows:
                row_issues.append(
                    {
                        "row": r,
                        "mrn": mrn,
                        "cms": None,
                        "issue_type": "Duplicate MRN",
                        "description": f"MRN appears in rows: {rows_str}",
                    }
                )
            issues.append(f"OASCAPHS: Duplicate MRN '{mrn}' found in rows {rows_str}")

    # check validity of telephone numbers using phonenumbers package
    if tel_col:
        for r, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if is_blank_row(row):
                continue
            tel_val = row[tel_col - 1]
            mrn_val = row[mrn_col - 1] if mrn_col else None
            cms_val = row[cms_col - 1] if cms_col else None

            if tel_val and str(tel_val).strip():
                tel_str = str(tel_val).strip()
                try:
                    phone_number = phonenumbers.parse(tel_str, "US")
                    if not phonenumbers.is_valid_number(phone_number):
                        row_issues.append(
                            {
                                "row": r,
                                "mrn": mrn_val,
                                "cms": cms_val,
                                "issue_type": "Invalid Telephone Number",
                                "description": f"Telephone '{tel_str}' is not a valid number",
                            }
                        )
                        issues.append(
                            f"OASCAPHS Row {r}: Invalid telephone number '{tel_str}'"
                        )
                except phonenumbers.NumberParseException:
                    row_issues.append(
                        {
                            "row": r,
                            "mrn": mrn_val,
                            "cms": cms_val,
                            "issue_type": "Invalid Telephone Number Format",
                            "description": f"Telephone '{tel_str}' has invalid format",
                        }
                    )
                    issues.append(
                        f"OASCAPHS Row {r}: Telephone '{tel_str}' has invalid format"
                    )

    # find placeholder/test names in patient name col
    if name_col:
        placeholder_names = {
            "test",
            "patient",
            "sample",
            "john doe",
            "jane doe",
            "asdf",
            "qwerty",
            "foo bar",
        }
        for r, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if is_blank_row(row):
                continue
            name_val = row[name_col - 1]
            mrn_val = row[mrn_col - 1] if mrn_col else None
            cms_val = row[cms_col - 1] if cms_col else None

            if name_val and str(name_val).strip():
                name_str = str(name_val).strip().lower()
                for name in placeholder_names:
                    if name in name_str:
                        row_issues.append(
                            {
                                "row": r,
                                "mrn": mrn_val,
                                "cms": cms_val,
                                "issue_type": "Placeholder Name",
                                "description": f"Patient Name '{name_val}' contains placeholder/test name",
                            }
                        )
                        issues.append(
                            f"OASCAPHS Row {r}: Patient Name '{name_val}' contains placeholder/test name"
                        )
                        break

    return issues, row_issues
