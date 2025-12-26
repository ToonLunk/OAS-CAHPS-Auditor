#!/usr/bin/env python3
import openpyxl
import os
import re
import sys
import uuid
import webbrowser
from tqdm import tqdm
from dotenv import load_dotenv
from audit_printer import save_report, build_report
from audit_lib_funcs import *

__version__ = "0.63.5"
version = __version__


def check_for_updates():
    """Check GitHub for latest version and notify if update available."""
    try:
        import urllib.request
        import json
        from packaging import version as pkg_version
        
        url = "https://api.github.com/repos/ToonLunk/OAS-CAHPS-Auditor/releases"
        req = urllib.request.Request(url)
        req.add_header('User-Agent', 'OAS-CAHPS-Auditor')
        
        with urllib.request.urlopen(req, timeout=3) as response:
            releases = json.loads(response.read().decode())
            if releases:
                latest_version = releases[0].get('tag_name', '').lstrip('v')
                
                # Only notify if the latest version is greater than current version
                if latest_version and pkg_version.parse(latest_version) > pkg_version.parse(version):
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
    
    # Extract SID prefix (3-letter code) for display
    sid_prefix = header_sid[:3] if header_sid and len(header_sid) >= 3 else None
    
    # Look up client name from SID registry
    sid_registry_name = None
    if sid_prefix:
        from audit_lib_funcs import lookup_sid_client_name
        sid_registry_name = lookup_sid_client_name(sid_prefix, show_missing_warning=True)

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

    # Check for required headers (returns mapping and list of any missing)
    mapping, missing_req_headers = check_req_headers(headers)

    sid_col = mapping.get("SID")
    pat_col = mapping.get("PATIENT NAME")
    addr1_col = mapping.get("ADDRESS1")
    city_col = mapping.get("CITY")
    state_col = mapping.get("STATE")
    zip_col = mapping.get("ZIP")
    tel_col = mapping.get("TELEPHONE")
    svc_col = mapping.get("SERVICE DATE")
    gender_col = mapping.get("GENDER")
    age_col = mapping.get("AGE")
    mrn_col = mapping.get("MRN")
    surg_cat_col = mapping.get("SURGICAL CATEGORY")
    att_col = mapping.get("ATT")
    lag_col = mapping.get("LAG")
    id_col = mapping.get("ID")
    fd_col = mapping.get("FD")
    lg_col = mapping.get("LG")
    em_col = mapping.get("E/M")
    email_col = mapping.get("EMAIL ADDRESS")
    cms_col = mapping.get("CMS INDICATOR")
    lang_col = mapping.get("SURVEY LANGUAGE")

    issues = []

    # Validate SID sequence (only for rows with CMS=1) if columns exist
    sid_issues = []
    sid_row_issues = []
    if sid_col and cms_col:
        sid_issues, sid_row_issues = validate_sid_sequence(sheet, sid_col, cms_col, header_sid)  # type: ignore
        issues.extend(sid_issues)

    # Validate INEL tab REPEAT entries
    inel_issues = []
    inel_row_issues = []
    if "INEL" in wb.sheetnames:
        inel_sheet = wb["INEL"]
        inel_issues, inel_row_issues = validate_inel_repeat_rows(inel_sheet)
        issues.extend(inel_issues)

    # Extract service date range and validate blank dates
    service_date_range = None
    blank_date_row_issues = []
    if svc_col:
        service_date_range, blank_date_issues, blank_date_row_issues = extract_service_date_range(
            sheet, svc_col, mrn_col=mrn_col, cms_col=cms_col
        )
        issues.extend(blank_date_issues)

    # Calculate E/M totals if columns exist
    total_em = None
    emails = None
    mailings = None
    non_reported = None
    cms1_count = None
    
    if cms_col and em_col:
        try:
            total_em, emails, mailings, non_reported, cms1_count = calc_e_m_total(
                sheet, cms_col, em_col
            )  # type: ignore
        except Exception as e:
            issues.append(f"Error calculating E/M totals: {str(e)}")

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
        sid_prefix=sid_prefix,
        sid_registry_name=sid_registry_name,
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
        inel_row_issues=inel_row_issues,  # INEL validation issues
        service_date_range=service_date_range,  # Service date range
        blank_date_row_issues=blank_date_row_issues,  # Blank date issues
    )

    return file_path, report_lines, service_date_range


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
                file_path, report_lines, service_date_range = audit_excel(filename)
                final_file = save_report(file_path, report_lines, version=version, service_date_range=service_date_range)
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
        file_path, report_lines, service_date_range = audit_excel(file_path)
        final_file = save_report(file_path, report_lines, version=version, service_date_range=service_date_range)
        
        # Open the report in the default browser
        try:
            webbrowser.open('file:///' + os.path.abspath(final_file).replace('\\', '/'))
            print(f"Opening report in your default browser...")
        except Exception as e:
            print(f"Could not automatically open browser: {e}")
        
        # Print clickable link for easy access
        print(f"\nReport link: file:///{os.path.abspath(final_file).replace(chr(92), '/')}")
        
    except Exception as e:
        # For single file mode, exit with error
        sys.exit(1)
