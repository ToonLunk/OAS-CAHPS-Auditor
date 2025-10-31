import re
import os
import datetime
import sys
import datetime
from openpyxl.worksheet.worksheet import Worksheet
import math


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
        if cnt == 0:
            # allow occasional blank rows inside the sparse block up to max_blank_within_block
            if sparse_run == 0:
                # skip leading blank rows after dense region
                start_idx += 1
                i += 1
                continue
            blanks_in_run += 1
            if blanks_in_run > max_blank_within_block:
                break
        elif cnt <= 2:
            sparse_run += 1
        else:
            # dense row encountered -> end of sparse block
            break
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


def build_report(
    wb,
    sheet,
    file_path,
    version,
    audit_id,
    missing_req_headers,
    patients_submitted,
    eligible_patients,
    sample_size,
    emails,
    mailings,
    total_em,
    non_reported,
    cms1_count,
    headers,
    issues,
    count_nonempty_rows,
    classify_cpt,
    cpt_is_ineligible,
    cms_col=None,
    find_frame_inel_count=None,
):
    """
    Build the textual audit report for saving as .txt
    """

    report_lines = []
    report_lines.append("=================================================")
    report_lines.append(f"      TB's EXCEL AUDITOR v{version}\n")
    tor = datetime.datetime.now()
    time_of_report = tor.strftime("%m/%d/%Y %H:%M:%S")

    # file modified time and client filename-before-#
    modified_ts = "N/A"
    try:
        modified_ts = datetime.datetime.fromtimestamp(
            os.path.getmtime(file_path)
        ).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        pass
    basefname = os.path.basename(file_path)
    base_before_hash = basefname.split("#", 1)[0]

    report_lines.append(f"   Report Date: {time_of_report}")
    report_lines.append(f"   Audit ID: {audit_id}")
    report_lines.append(f"   Client: {base_before_hash}")
    report_lines.append(f"   File last modified/saved: {modified_ts}")
    report_lines.append("=================================================\n")

    # missing required headers -> issues
    if missing_req_headers:
        for header in missing_req_headers:
            issues.append(f"Missing REQUIRED Header: {header}!")

    # Header/Footer extracted values
    report_lines.append(">> HEADER / FOOTER VALUES")
    if patients_submitted is not None:
        report_lines.append(
            f"  • Patients Submitted (from header): {patients_submitted}"
        )
    if eligible_patients is not None:
        report_lines.append(f"  • Eligible Patients (from footer): {eligible_patients}")
    if sample_size is not None:
        report_lines.append(f"  • Sample Size (from footer): {sample_size}")
    report_lines.append("")

    # OASCAPHS tab analysis
    report_lines.append(">> OASCAPHS TAB ANALYSIS")
    report_lines.append(f"  • Emails counted: {emails}")
    report_lines.append(f"  • Mailings counted: {mailings}")
    report_lines.append(f"  • Total of E/M: {total_em}")
    report_lines.append(f"  • Non-Reported entries: {non_reported}")
    report_lines.append(f"  • Rows with CMS INDICATOR = 1: {cms1_count}")
    estimated_percentage = math.ceil((sample_size / eligible_patients) * 100)
    report_lines.append(f"  • Estimated Selection %: ~{estimated_percentage}%")
    report_lines.append("")

    # VALIDATION CHECKS
    report_lines.append(">> VALIDATION CHECKS")
    if sample_size is not None and cms1_count != sample_size:
        issue_msg = f"X *WARNING* Sample Size mismatch: expected {sample_size}, found {cms1_count} rows with CMS=1"
        report_lines.append(issue_msg)
        issues.append(issue_msg)
    else:
        report_lines.append("  ✓ Sample Size matches CMS=1 row count")

    if sample_size is not None and total_em != sample_size:
        issue_msg = (
            f"X *WARNING* E/M total mismatch: {total_em} vs Sample Size {sample_size}"
        )
        report_lines.append(issue_msg)
        issues.append(issue_msg)
    else:
        report_lines.append("  ✓ E/M total matches Sample Size")

    # POP tab check
    if "POP" in wb.sheetnames and patients_submitted is not None:
        pop_sheet = wb["POP"]
        pop_rows = count_nonempty_rows(pop_sheet)
        if patients_submitted != pop_rows:
            issue_msg = f"X *WARNING* Submitted mismatch: header says {patients_submitted}, POP tab has {pop_rows} rows. (if this is within ~4, this is expected due to various client header sizes)"
            report_lines.append(issue_msg)
            issues.append(issue_msg)
        else:
            report_lines.append("  ✓ Submitted count matches POP tab row count")
    else:
        issue_msg = "X *WARNING* POP tab missing or Submitted value not found"
        report_lines.append(issue_msg)
        issues.append(issue_msg)

    # UPLOAD tab row count comparison
    if "UPLOAD" in wb.sheetnames:
        upload_sheet = wb["UPLOAD"]
        upload_rows = count_nonempty_rows(upload_sheet)
        oascaphs_rows = count_nonempty_rows(sheet)
        if upload_rows != oascaphs_rows:
            issue_msg = f"X *WARNING* UPLOAD mismatch: {upload_rows} rows vs {oascaphs_rows} rows in OASCAPHS"
            report_lines.append(issue_msg)
            issues.append(issue_msg)
        else:
            report_lines.append("  ✓ UPLOAD tab row count matches OASCAPHS")
    else:
        issue_msg = "  ! UPLOAD tab missing"
        report_lines.append(issue_msg)
        issues.append(issue_msg)

    # 1. Surgical Category Validation (OASCAPHS)
    report_lines.append("")
    cpt_col = headers.get("CPT")
    cat_col = headers.get("SURGICAL CATEGORY")
    if cpt_col and cat_col:
        for r, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if not any(row):
                continue
            cpt_val = row[cpt_col - 1]
            cat_val = row[cat_col - 1]
            expected = classify_cpt(str(cpt_val) if cpt_val else "")
            if expected != cat_val:
                issues.append(
                    f"OASCAPHS Row {r}: CPT {cpt_val} has category {cat_val}, expected {expected}"
                )
    else:
        issue_msg = "Missing CPT or SURGICAL CATEGORY column in OASCAPHS"
        issues.append(issue_msg)

    # count rows in INEL and FRAME
    inel_count = None
    if "INEL" in wb.sheetnames:
        inel_sheet = wb["INEL"]
        inel_count = count_nonempty_rows(inel_sheet)

    frame_inel_count = None
    if "FRAME" in wb.sheetnames and find_frame_inel_count is not None:
        frame_sheet = wb["FRAME"]
        try:
            frame_inel_count = find_frame_inel_count(frame_sheet)
        except Exception:
            frame_inel_count = None

    # 2. UPLOAD vs OASCAPHS comparison (value-by-value)
    if "UPLOAD" in wb.sheetnames:
        upload_sheet = wb["UPLOAD"]
        up_headers = {
            cell.value: idx
            for idx, cell in enumerate(
                next(upload_sheet.iter_rows(min_row=1, max_row=1)), start=1
            )
        }
        oas_headers = headers
        ignore_cols = {"LG", "FD", "ID", "ATT", "LAG"}
        common_cols = (
            set(up_headers.keys()).intersection(oas_headers.keys()) - ignore_cols
        )

        for col in common_cols:
            up_idx = up_headers[col] - 1
            oas_idx = oas_headers[col] - 1
            max_rows = min(
                count_nonempty_rows(upload_sheet) + 1, count_nonempty_rows(sheet) + 1
            )
            for r in range(2, max_rows + 1):
                up_val = upload_sheet.cell(r, up_idx + 1).value
                oas_val = sheet.cell(r, oas_idx + 1).value
                if up_val != oas_val:
                    issues.append(
                        f"UPLOAD Row {r}, Column {col}: UPLOAD={up_val} vs OASCAPHS={oas_val}"
                    )
    else:
        issues.append("UPLOAD tab missing")

    # Add INEL and FRAME info to report_lines
    if inel_count is not None:
        report_lines.append(f"  • INEL tab non-empty rows: {inel_count}")
    else:
        report_lines.append("  • INEL tab missing")
        issues.append("INEL tab missing")

    if frame_inel_count is not None:
        report_lines.append(
            f"  • FRAME tab 6-month-repeat INEL PT IDs: {frame_inel_count}"
        )
    else:
        report_lines.append("  • FRAME tab missing or FRAME count not computed")
        # do not duplicate issue here unless you want explicit issue

    # Combined check vs submitted (if header has Patients Submitted)
    if patients_submitted is not None:
        total_inel_combined = (inel_count or 0) + (frame_inel_count or 0)
        report_lines.append(
            f"  • Combined Ineligible (INEL tab + 6M-MONTH REPEATS): {total_inel_combined}"
        )
        if eligible_patients is not None:
            if eligible_patients + total_inel_combined != patients_submitted:
                issue_msg = f"X *WARNING* Submitted mismatch: eligible ({eligible_patients}) + combined INEL ({total_inel_combined}) != submitted (submitted: {patients_submitted})"
                report_lines.append(issue_msg)
                issues.append(issue_msg)
            else:
                report_lines.append("  ✓ Eligible + Combined INEL matches Submitted")

    # 3. CPT Ineligibility Check (only report when CMS == 1)
    cpt_ineligible_rows = []
    if cpt_col:
        for r, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if not any(row):
                continue
            cpt_val = row[cpt_col - 1]
            cms_val = row[cms_col - 1] if cms_col else None
            cms_int = None
            try:
                if cms_val is not None and str(cms_val).strip() != "":
                    cms_int = int(float(str(cms_val).strip()))
            except Exception:
                cms_int = None

            ineligible, reason = cpt_is_ineligible(cpt_val)
            if ineligible and cms_int == 1:
                msg = f"OASCAPHS Row {r}: CPT {cpt_val} ineligible ({reason})"
                cpt_ineligible_rows.append((r, cpt_val, reason))
                issues.append(msg)
    else:
        issues.append("CPT column missing in OASCAPHS for ineligibility check")

    # ISSUES section
    report_lines.append("\n>> ISSUES FOUND")
    if issues:
        for idx, issue in enumerate(issues, start=1):
            report_lines.append(f"  • {issue}")
    else:
        report_lines.append("  • no issues found")

    # CPT ineligible summary
    report_lines.append("\n>> CPT INELIGIBLE SUMMARY")
    report_lines.append(
        "Note: Some ineligible CPT codes are expected to be in the non-report (CMS=2) section!"
    )
    if cpt_ineligible_rows:
        report_lines.append(
            f"  • Total ineligible CPT rows found: {len(cpt_ineligible_rows)}"
        )
        for r, cpt, reason in cpt_ineligible_rows:
            report_lines.append(f"    - Row {r}: CPT={cpt} ; Reason={reason}")
    else:
        report_lines.append("  • no ineligible CPT codes found")

        report_lines.append("\n>> ESTIMATED QTR SHEET LINE")
        report_lines.append(
            f"\n{base_before_hash} | {non_reported} | {emails} | {mailings} | (~){estimated_percentage}% | {patients_submitted} | {eligible_patients} | {sample_size}"
        )

        report_lines.append("\n=================================================")
        report_lines.append("        END OF REPORT")
        report_lines.append("=================================================")

    return report_lines, issues


def save_report(file_path, report_lines, failure_reason="", version="0.0-alpha"):
    # --- Write report to .txt file ---
    base_name = os.path.splitext(file_path)[0]
    if not failure_reason:
        report_file = base_name + ".txt"
    else:
        report_file = "UNABLE_TO_AUDIT " + base_name + ".txt"

    # build audits directory next to the original path (or in cwd if no dir)
    base_dir = os.path.dirname(report_file) or "."
    audits_dir = os.path.join(base_dir, "audits")
    os.makedirs(audits_dir, exist_ok=True)

    # timestamp and final filename
    name, ext = os.path.splitext(os.path.basename(report_file))
    timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    final_report_file = os.path.join(audits_dir, f"{name}_{timestamp}{ext}")

    # prevent accidental overwrite (very unlikely because of timestamp, but safe)
    if os.path.isfile(final_report_file):
        print(
            f"--- File already exists! This auditor will not overwrite files. If you wish to run a new audit on this file, please delete the previous audit:  {final_report_file}"
        )
        input("Press enter to exit: ")
        print("\n")
        sys.exit(99)

    with open(final_report_file, "w", encoding="utf-8") as f:
        if not failure_reason:
            f.write("\n".join(report_lines))
        else:
            report_lines_list = []
            report_lines_list.append(
                "================================================="
            )
            report_lines_list.append(f"      TB's EXCEL AUDITOR v{version}\n")
            report_lines_list.append(
                "=================================================\n\n"
            )
            report_lines_list.append(report_lines)
            report_lines_list.append(f"Failure reason: {failure_reason}\n")
            report_lines_list.append(
                "\n================================================="
            )
            report_lines_list.append("        END OF REPORT")
            report_lines_list.append(
                "================================================="
            )
            f.write("\n".join(report_lines_list))

    if not failure_reason:
        print(f"--- Audit complete. Report saved to {final_report_file}\n")
    else:
        print(
            f"--- Audit could not run on this file! Information saved to {final_report_file}\n"
        )
