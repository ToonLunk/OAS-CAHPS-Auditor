#!/usr/bin/env python3
import openpyxl
import os
import re
import sys
import uuid
from tqdm import tqdm
from dotenv import load_dotenv
from audit_printer import save_report, build_report
from audit_lib_funcs import *

__version__ = "0.60"
version = __version__


def check_for_updates():
    """Check GitHub for latest version and notify if update available."""
    try:
        import urllib.request
        import json
        
        url = "https://api.github.com/repos/ToonLunk/OAS-CAHPS-Auditor/releases"
        req = urllib.request.Request(url)
        req.add_header('User-Agent', 'OAS-CAHPS-Auditor')
        
        with urllib.request.urlopen(req, timeout=3) as response:
            releases = json.loads(response.read().decode())
            if releases:
                latest_version = releases[0].get('tag_name', '').lstrip('v')
                
                if latest_version and latest_version != version:
                    print(f"\nUpdate available: v{latest_version} (current: v{version})")
                    print(f"Download: https://github.com/ToonLunk/OAS-CAHPS-Auditor/releases/latest\n")
    except Exception as e:
        # Silently fail if unable to check
        pass


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

    # finds the two-letter code (like "TB") - can be anywhere in header
    header_clean = re.sub(r"&\[[^\]]+\]", "", header)
    # Match exactly 2 uppercase letters not part of a longer word
    m = re.search(r"(?<![A-Z])([A-Z]{2})(?![A-Z])", header_clean)
    two_letter_code = m.group(1) if m else ""

    # Extract SID from header (should be first SID in sequence)
    sid_match = re.search(r"([A-Z]{3}\d+)", header_clean)
    header_sid = sid_match.group(1) if sid_match else None

    # convert letters to alphabet positions (A=1, B=2) and append to UUID
    nums = "".join(str(ord(c) - 64) for c in two_letter_code)
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
        raise Exception(f"Missing required columns: {missing_req_headers}")

    sid_col = mapping["SID"]
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

    # Validate SID sequence (only for rows with CMS=1)
    sid_issues, sid_row_issues = validate_sid_sequence(sheet, sid_col, cms_col, header_sid)
    issues.extend(sid_issues)

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
        raise Exception("No E/M counts found")

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
        em_col=em_col,
        find_frame_inel_count=find_frame_inel_count,  # optional
        mrn_col=mrn_col,  # optional
        sid_col=sid_col,  # SID column
        sid_row_issues=sid_row_issues,  # SID validation issues
    )

    return file_path, report_lines


if __name__ == "__main__":
    check_for_updates()
    
    if len(sys.argv) != 2:
        print("Usage: audit <excel_file> or audit --all")
        print("Options:")
        print("  --all       Process all Excel files in the current directory")
        print("  --help,-h   Show this help message")
        print("  --version,-v Show version information")
        print("\n")
        print(
            "Need help? Visit https://github.com/ToonLunk/OAS-CAHPS-Auditor, or contact support."
        )
        sys.exit(1)

    arg = sys.argv[1]

    # Handle --all flag to process all files in current directory
    if arg == "--all":
        # Get list of Excel files
        excel_files = [
            f for f in os.listdir(".") if f.endswith((".xlsx", ".xls", ".xlsm"))
        ]

        if not excel_files:
            print("No Excel files found in current directory.")
            sys.exit(0)

        files_processed = 0
        print(f"Found {len(excel_files)} Excel file(s) to process.\n")

        # Process with progress bar
        for filename in tqdm(excel_files, desc="Processing files", unit="file"):
            try:
                file_path, report_lines = audit_excel(filename)
                final_file = save_report(file_path, report_lines, version=version)
                tqdm.write(f"✓ {filename} -> {final_file}")
                files_processed += 1
            except Exception as e:
                tqdm.write(f"✗ {filename}: {e}")

        print(
            f"\nCompleted: {files_processed}/{len(excel_files)} file(s) processed successfully."
        )
        sys.exit(0)

    if arg == "--help" or arg == "-h":
        print("Usage: audit <excel_file> or audit --all")
        print("Options:")
        print("  --all       Process all Excel files in the current directory")
        print("  --help,-h   Show this help message")
        print("  --version,-v Show version information")
        print("\n")
        print(
            "Need help? Visit https://github.com/ToonLunk/OAS-CAHPS-Auditor, or contact support."
        )
        sys.exit(0)
    if arg == "--version" or arg == "-v":
        print(f"OAS auditor version {version}")
        sys.exit(0)

    # Handle single file
    file_path = arg
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' not found.")
        sys.exit(1)

    try:
        file_path, report_lines = audit_excel(file_path)
        final_file = save_report(file_path, report_lines, version=version)
    except Exception as e:
        # For single file mode, exit with error
        sys.exit(1)
