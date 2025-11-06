import os
import sys
import datetime
import math
from audit_lib_funcs import check_address


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
    addr1_col,
    city_col,
    state_col,
    zip_col,
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
        TOL = 4
        if abs(patients_submitted - pop_rows) > TOL:
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
            issue_msg = f"--X *WARNING* UPLOAD mismatch: {upload_rows} rows vs {oascaphs_rows} rows in OASCAPHS"
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

    # INVALID ADDRESSES section
    report_lines.append("\n>> INVALID ADDRESSES FOUND")
    # audit addresses using google's package
    invalid_addresses, noted_addresses = check_address(
        sheet, addr1_col, city_col, state_col, zip_col
    )
    if invalid_addresses:
        for address in invalid_addresses:
            report_lines.append(
                f"X *WARNING* Potential invalid Address found: {address}"
            )
    else:
        report_lines.append("  • no invalid addresses found")

    report_lines.append("\n>> PROBLEMATIC ADDRESSES FOUND")
    # possibly problematic addresses
    if noted_addresses:
        for address in noted_addresses:
            report_lines.append(f"? *NOTE* Potential invalid Address found: {address}")
    else:
        report_lines.append("  • no problematic addresses found")

    report_lines.append("\n>> ESTIMATED QTR SHEET LINE")
    report_lines.append(
        f"\n{base_before_hash} | {non_reported} | {emails} | {mailings} | (~){estimated_percentage}% | {patients_submitted} | {eligible_patients} | {sample_size}"
    )

    report_lines.append("\n=================================================")
    report_lines.append("        END OF REPORT")
    report_lines.append("=================================================")

    return report_lines, issues


def save_report(file_path, report_lines, failure_reason="", version="0.0-alpha"):
    """
    Write report to .txt file in audits directory
    """
    # --- Write report to .txt file ---
    base_name = os.path.splitext(file_path)[0]
    report_file = base_name + ".txt"

    # build audits directory next to the original path (or in cwd if no dir)
    base_dir = os.path.dirname(report_file) or "."
    audits_dir = os.path.join(base_dir, "audits")
    if failure_reason:
        audits_dir = os.path.join(audits_dir, "unable_to_run_audit")
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
        # if there was an error/failure reason
        else:
            report_lines_list = []
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

            report_lines_list.append(
                "================================================="
            )
            report_lines_list.append(f"      TB's EXCEL AUDITOR v{version}\n")
            report_lines_list.append(f"   Report Date: {time_of_report}")
            report_lines_list.append(
                f"   No Audit ID available (failed audits do not get an ID)"
            )
            report_lines_list.append(f"   Client: {base_before_hash}")
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

    # return the full file name and path in case it needs to be read again
    return final_report_file
