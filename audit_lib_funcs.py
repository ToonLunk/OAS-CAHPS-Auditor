import re
import datetime
from openpyxl.worksheet.worksheet import Worksheet


# --- CPT ineligibility rules
INVALID_CPT_SET = {
    "11042",
    "11045",
    "16020",
    "16025",
    "16030",
    "19081",
    "19083",
    "19085",
    "20560",
    "20561",
    "25246",
    "27093",
    "29000",
    "29010",
    "29015",
    "29020",
    "29025",
    "29035",
    "29040",
    "29044",
    "29046",
    "29049",
    "29055",
    "29058",
    "29065",
    "29075",
    "29085",
    "29086",
    "29105",
    "29125",
    "29126",
    "29130",
    "29131",
    "29200",
    "29240",
    "29260",
    "29280",
    "29305",
    "29325",
    "29345",
    "29355",
    "29358",
    "29365",
    "29405",
    "29425",
    "29435",
    "29440",
    "29445",
    "29450",
    "29505",
    "29515",
    "29520",
    "29530",
    "29540",
    "29550",
    "29580",
    "29581",
    "29582",
    "29583",
    "29584",
    "29700",
    "29705",
    "29710",
    "29715",
    "29720",
    "29730",
    "29740",
    "29750",
    "29799",
    "31500",
    "32555",
    "36005",
    "36010",
    "36215",
    "36221",
    "36400",
    "36405",
    "36406",
    "36410",
    "36415",
    "36416",
    "36420",
    "36425",
    "36430",
    "36440",
    "36450",
    "36455",
    "36460",
    "36468",
    "36469",
    "36470",
    "36471",
    "36591",
    "36592",
    "36593",
    "36600",
    "36620",
    "36625",
    "36660",
    "37252",
    "38505",
    "49424",
    "51701",
    "51702",
    "51798",
    "59020",
    "59025",
    "59050",
}

EXPLICIT_VALID_SET = {
    "G0104",
    "G0105",
    "G0121",
    "G0260",
    "92920",
    "92921",
    "92928",
    "92929",
    "92978",
}


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
        # Ranges considered valid by VBA's IsInValidRange
        if (
            (10004 <= num <= 69990)
            or (93451 <= num <= 93462)
            or (93566 <= num <= 93572)
            or (93985 <= num <= 93986)
        ):
            return False, "numeric in valid range"
        else:
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

    if "E/M" in missing_req_headers or "CMS INDICATOR" in missing_req_headers:
        print("  ! AUDIT FAILED: CMS INDICATOR or E/M COLUMN MISSING!")
    if "SURVEY LANGUAGE" in missing_req_headers:
        print("  ! AUDIT FAILED: SURVEY LANGUAGE MISSING!")

    if missing_req_headers:
        # raise a simple exception containing the missing list
        raise ValueError(f"Missing required headers: {missing_req_headers}")

    return mapping
