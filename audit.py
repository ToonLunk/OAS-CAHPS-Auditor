#!/usr/bin/env python3
import openpyxl
import os
import re
import sys
import uuid
from audit_printer import save_report, build_report
from audit_lib_funcs import *

# versioning
version = "0.42-beta2"


def audit_excel(file_path):
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
    except:
        print(
            f"--- Critical Error opening {file_path}! Are you sure it's an Excel file?"
        )
        input("Press enter to continue: ")
        print("\n")
        sys.exit(1)
    sheet = wb["OASCAPHS"]

    # --- Extract and clean header/footer ---
    raw_header = pick_header(sheet)
    raw_footer = pick_footer(sheet)

    header = clean_hf_text(raw_header)
    footer = clean_hf_text(raw_footer)

    # --- Extract values from header/footer text ---
    submitted_match = re.search(r"SUBMITTED\s*=\s*(\d+)", header)
    patients_submitted = int(submitted_match.group(1)) if submitted_match else None

    # finds the final two-letter token at end of header string (uppercase letters only)
    m = re.search(r"([A-Z]{2})(?!.*[A-Z])", header)
    two_letter_code = m.group(1) if m else None

    # convert letters to alphabet positions (A=1, B=2) and append to UUID
    nums = "".join(str(ord(c) - 64) for c in (two_letter_code or ""))
    uuid_code = uuid.uuid4().hex
    audit_id = f"{uuid_code}{nums}"

    el_match = re.search(r"EL\s*=\s*(\d+)", footer)
    ss_match = re.search(r"SS\s*=\s*(\d+)", footer)
    eligible_patients = int(el_match.group(1)) if el_match else None
    sample_size = int(ss_match.group(1)) if ss_match else None

    # --- Find column indexes ---
    first_row = next(sheet.iter_rows(min_row=1, max_row=1))
    headers = {cell.value: idx for idx, cell in enumerate(first_row, start=1)}

    missing_req_headers = []
    try:
        mapping = check_req_headers(headers)
    except Exception as e:
        # e.args[0] is the missing_req_headers list we raised
        missing_req_headers = e.args[0] if e.args else []
        save_report(
            file_path,
            "FAILED - missing required columns!",
            failure_reason=str(missing_req_headers),
            version=version,
        )
        exit(2)

    pat_col = mapping["PATIENT NAME"]
    addr1_col = mapping["ADDRESS1"]
    city_col = mapping["CITY"]
    state_col = mapping["STATE"]
    zip_col = mapping["ZIP"]
    tel_col = mapping["TELEPHONE"]
    svc_col = mapping["SERVICE DATE"]
    gender_col = mapping["GENDER"]
    age_col = mapping["AGE"]
    mrn_col = mapping["MRN"]
    surg_cat_col = mapping["SURGICAL CATEGORY"]
    att_col = mapping["ATT"]
    lag_col = mapping["LAG"]
    id_col = mapping["ID"]
    fd_col = mapping["FD"]
    lg_col = mapping["LG"]
    em_col = mapping["E/M"]
    email_col = mapping["EMAIL ADDRESS"]
    cms_col = mapping["CMS INDICATOR"]
    lang_col = mapping["SURVEY LANGUAGE"]

    issues = []

    try:
        total_em, emails, mailings, non_reported, cms1_count = calc_e_m_total(
            sheet, cms_col, em_col
        )  # type: ignore
    except:
        save_report(
            file_path,
            "FAILED while calculating E/M!",
            failure_reason="no E/M counts",
            version=version,
        )
        exit(3)

    report_lines, issues = build_report(
        wb=wb,
        sheet=sheet,
        file_path=file_path,
        version=version,
        audit_id=audit_id,
        missing_req_headers=missing_req_headers,
        patients_submitted=patients_submitted,
        eligible_patients=eligible_patients,
        sample_size=sample_size,
        emails=emails,
        mailings=mailings,
        total_em=total_em,
        non_reported=non_reported,
        cms1_count=cms1_count,
        headers=headers,
        issues=issues,
        count_nonempty_rows=count_nonempty_rows,
        classify_cpt=classify_cpt,
        cpt_is_ineligible=cpt_is_ineligible,
        addr1_col=addr1_col,
        city_col=city_col,
        state_col=state_col,
        zip_col=zip_col,
        cms_col=cms_col,
        find_frame_inel_count=find_frame_inel_count,  # optional
    )

    return file_path, report_lines


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: audit <excel_file>")
        sys.exit(1)

    file_path = sys.argv[1]
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' not found.")
        sys.exit(1)

    file_path, report_lines = audit_excel(file_path)
    final_file = save_report(file_path, report_lines, version=version)
