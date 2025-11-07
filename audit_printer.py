import os
import sys
import datetime
import math
import re
from collections import defaultdict

from requests import head
from audit_lib_funcs import check_address, check_pop_upload_email_consistency


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
    em_col=None,
    find_frame_inel_count=None,
    mrn_col=None,
):
    """
    Build the HTML audit report for saving as .html
    """

    basefname = os.path.basename(file_path)
    base_before_hash = basefname.split("#", 1)[0]

    # Start HTML document with helper function
    report_lines = _build_html_header(file_path, version, audit_id)

    # Track row-based issues separately for table display
    row_issues = []  # List of dicts: {row, mrn, cms, issue_type, description}

    # missing required headers -> issues
    if missing_req_headers:
        for header in missing_req_headers:
            issues.append(f"Missing REQUIRED Header: {header}!")

    # Header/Footer extracted values
    report_lines.append("<h2>OASCAPHS TAB ANALYSIS</h2>")
    report_lines.append("<table class='data-table'>")
    if patients_submitted is not None:
        report_lines.append(
            f"<tr><td>Patients Submitted (from header)</td><td>{patients_submitted}</td></tr>"
        )
    if eligible_patients is not None:
        report_lines.append(
            f"<tr><td>Eligible Patients (from footer)</td><td>{eligible_patients}</td></tr>"
        )
    if sample_size is not None:
        report_lines.append(
            f"<tr><td>Sample Size (from footer)</td><td>{sample_size}</td></tr>"
        )
    report_lines.append("</table>")

    # OASCAPHS tab analysis
    report_lines.append("<table class='data-table'>")
    report_lines.append(f"<tr><td>Emails counted</td><td>{emails}</td></tr>")
    report_lines.append(f"<tr><td>Mailings counted</td><td>{mailings}</td></tr>")
    report_lines.append(f"<tr><td>Total of E/M</td><td>{total_em}</td></tr>")
    report_lines.append(
        f"<tr><td>Non-Reported entries</td><td>{non_reported}</td></tr>"
    )
    report_lines.append(
        f"<tr><td>Rows with CMS INDICATOR = 1</td><td>{cms1_count}</td></tr>"
    )
    report_lines.append("</table>")

    # count rows in INEL and FRAME (needed for validation checks)
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

    # VALIDATION CHECKS
    report_lines.append("<h2>VALIDATION CHECKS</h2>")

    # Tab counts in table format
    report_lines.append("<table class='data-table'>")
    if inel_count is not None:
        report_lines.append(f"<tr><td>INEL tab rows</td><td>{inel_count}</td></tr>")
    else:
        issues.append("INEL tab missing")

    if frame_inel_count is not None:
        report_lines.append(
            f"<tr><td>FRAME tab 6-month repeats</td><td>{frame_inel_count}</td></tr>"
        )

    if patients_submitted is not None:
        total_inel_combined = (inel_count or 0) + (frame_inel_count or 0)
        report_lines.append(
            f"<tr><td>Combined Ineligible</td><td>{total_inel_combined}</td></tr>"
        )
    report_lines.append("</table>")

    # Validation checks in table format
    report_lines.append("<table class='data-table'>")

    # Check 1: Sample Size matches Reported
    if sample_size is not None and cms1_count != sample_size:
        issue_msg = f"<strong>WARNING:</strong> Sample Size mismatch: expected {sample_size}, found {cms1_count} rows with CMS=1"
        report_lines.append(
            f"<tr><td style='color: red;'>{issue_msg}</td><td>✗</td></tr>"
        )
        issues.append(issue_msg)
    else:
        report_lines.append(
            "<tr><td>Sample Size matches Reported</td><td style='color: #28a745;'>✓</td></tr>"
        )

    # Check 2: E/M total matches Sample Size
    if sample_size is not None and total_em != sample_size:
        issue_msg = f"<strong>WARNING:</strong> Reported total mismatch: <strong>{total_em}</strong> vs Sample Size <strong>{sample_size}</strong>"
        report_lines.append(
            f"<tr><td '>{issue_msg}</td><td style='color: red;>✗</td></tr>"
        )
        issues.append(issue_msg)
    else:
        report_lines.append(
            "<tr><td>E/M total matches Sample Size</td><td style='color: #28a745;'>✓</td></tr>"
        )

    # Check 3: Submitted matches POP tab
    if "POP" in wb.sheetnames and patients_submitted is not None:
        pop_sheet = wb["POP"]
        pop_rows = count_nonempty_rows(pop_sheet)
        TOL = 4
        if abs(patients_submitted - pop_rows) > TOL:
            issue_msg = f"Submitted mismatch: header says {patients_submitted}, POP tab has {pop_rows} rows. (if this is within ~4, this is expected due to various client header sizes)"
            report_lines.append(
                f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>"
            )
            issues.append(f"<strong>WARNING:</strong> {issue_msg}")
        else:
            report_lines.append(
                "<tr><td>Submitted matches POP tab</td><td style='color: #28a745;'>✓</td></tr>"
            )
    else:
        issue_msg = (
            "<strong>WARNING:</strong> POP tab missing or Submitted value not found"
        )
        report_lines.append(
            f"<tr><td '>{issue_msg}</td><td style='color: red;>✗</td></tr>"
        )
        issues.append(issue_msg)

    # Check 4: UPLOAD and OASCAPHS row counts match
    if "UPLOAD" in wb.sheetnames:
        upload_sheet = wb["UPLOAD"]
        upload_rows = count_nonempty_rows(upload_sheet)
        oascaphs_rows = count_nonempty_rows(sheet)
        if upload_rows != oascaphs_rows:
            issue_msg = f"<strong>WARNING:</strong> UPLOAD mismatch: {upload_rows} rows vs {oascaphs_rows} rows in OASCAPHS"
            report_lines.append(
                f"<tr><td '>{issue_msg}</td><td style='color: red;>✗</td></tr>"
            )
            issues.append(issue_msg)
        else:
            report_lines.append(
                "<tr><td>UPLOAD and OASCAPHS row counts match</td><td style='color: #28a745;'>✓</td></tr>"
            )
    else:
        issue_msg = "UPLOAD tab missing"
        report_lines.append(
            f"<tr><td>{issue_msg}</td><td style='color: red;>✗</td></tr>"
        )
        issues.append(issue_msg)

    estimated_percentage = math.ceil((sample_size / eligible_patients) * 100)

    report_lines.append("</table>")

    report_lines.append("<h2>ESTIMATED QTR SHEET LINE</h2>")
    report_lines.append("<table class='excel-style'>")
    report_lines.append("<tr>")
    report_lines.append(
        "<th>Client</th>" "<th>Non-Reported</th>" "<th>Emails</th>" "<th>Mailings</th>"
    )
    report_lines.append(
        "<th>Selection %</th>"
        "<th>Submitted</th>"
        "<th>Eligible</th>"
        "<th>Sample Size</th>"
    )
    report_lines.append("</tr>")
    report_lines.append("<tr>")
    report_lines.append(f"<td>{base_before_hash}</td>")
    report_lines.append(f"<td>{non_reported}</td>")
    report_lines.append(f"<td>{emails}</td>")
    report_lines.append(f"<td>{mailings}</td>")
    report_lines.append(f"<td>~{estimated_percentage}%</td>")
    report_lines.append(f"<td>{patients_submitted}</td>")
    report_lines.append(f"<td>{eligible_patients}</td>")
    report_lines.append(f"<td>{sample_size}</td>")
    report_lines.append("</tr>")
    report_lines.append("</table>")

    # ==============================================================
    # DATA QUALITY VALIDATION SECTION
    # ==============================================================
    # Check for various data quality issues and add to row_issues

    svc_col = headers.get("SERVICE DATE")
    age_col = headers.get("AGE")
    email_col = headers.get("EMAIL ADDRESS")
    lang_col = headers.get("SURVEY LANGUAGE")
    tel_col = headers.get("TELEPHONE")

    # Track service dates to check they're all in the same month
    service_dates = []
    # Track MRNs to check for duplicates
    mrn_tracker = defaultdict(list)

    for r, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if not any(row):
            continue

        mrn_val = row[mrn_col - 1] if mrn_col else None
        cms_val = row[cms_col - 1] if cms_col else None
        em_val = row[em_col - 1] if em_col else None

        # Track MRN for duplicate check
        if mrn_val:
            mrn_tracker[mrn_val].append(r)

        # 1. SERVICE DATE - collect all dates for month validation
        if svc_col:
            svc_date = row[svc_col - 1]
            if isinstance(svc_date, datetime.datetime):
                service_dates.append((r, mrn_val, svc_date))

        # 2. AGE - must be 18 or older (only matters when CMS=1)
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

        # 3. EMAIL ADDRESS - validate format when present
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

        # 4. SURVEY LANGUAGE - must be en, es, ko, zh, or m (lowercase)
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

        # 5. E/M and CMS INDICATOR logic
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

    # ==============================================================
    # END OF DATA QUALITY VALIDATION SECTION
    # ==============================================================

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
                # Get MRN and CMS for this row
                mrn_val = row[mrn_col - 1] if mrn_col else None
                cms_val = row[cms_col - 1] if cms_col else None

                row_issues.append(
                    {
                        "row": r,
                        "mrn": mrn_val,
                        "cms": cms_val,
                        "issue_type": "Surgical Category Mismatch",
                        "description": f"CPT {cpt_val} has category {cat_val}, expected {expected}",
                    }
                )
                issues.append(
                    f"OASCAPHS Row {r}: CPT {cpt_val} has category {cat_val}, expected {expected}"
                )
    else:
        issue_msg = "Missing CPT or SURGICAL CATEGORY column in OASCAPHS"
        issues.append(issue_msg)

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

    # 2b. Cross-tab consistency: POP vs UPLOAD email matching
    if "UPLOAD" in wb.sheetnames:
        upload_sheet = wb["UPLOAD"]
        up_headers = {
            cell.value: idx
            for idx, cell in enumerate(
                next(upload_sheet.iter_rows(min_row=1, max_row=1)), start=1
            )
        }

        # Get MRN and Email columns from UPLOAD
        upload_mrn_col = up_headers.get("MRN")
        upload_email_col = up_headers.get("EMAIL ADDRESS")

        if upload_mrn_col and upload_email_col:
            email_mismatches = check_pop_upload_email_consistency(
                wb, upload_sheet, upload_mrn_col, upload_email_col
            )

            # Add mismatches to row_issues for table display
            for upload_row, mrn, upload_email, pop_email in email_mismatches:
                # Check if this is an error message (when upload_row is "N/A")
                if upload_row == "N/A":
                    issues.append(
                        f"<strong>WARNING:</strong> POP/UPLOAD Email Check: {pop_email}"
                    )
                else:
                    row_issues.append(
                        {
                            "row": f"UPLOAD {upload_row}",
                            "mrn": mrn,
                            "cms": None,
                            "issue_type": "Email Mismatch (POP vs UPLOAD)",
                            "description": f"UPLOAD: '{upload_email}' vs POP: '{pop_email}'",
                        }
                    )
                    issues.append(
                        f"UPLOAD Row {upload_row}: Email mismatch for MRN {mrn} - UPLOAD: '{upload_email}' vs POP: '{pop_email}'"
                    )

    # Check combined ineligible math (moved validation to earlier section)
    if patients_submitted is not None and eligible_patients is not None:
        total_inel_combined = (inel_count or 0) + (frame_inel_count or 0)
        if eligible_patients + total_inel_combined != patients_submitted:
            issue_msg = f"<strong>WARNING:</strong> Math error: Eligible ({eligible_patients}) + Combined INEL ({total_inel_combined}) = {eligible_patients + total_inel_combined}, but Submitted = {patients_submitted}"
            issues.append(issue_msg)

    # 3. CPT Ineligibility Check (only report when CMS == 1)
    cpt_ineligible_rows = []
    if cpt_col:
        for r, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if not any(row):
                continue
            cpt_val = row[cpt_col - 1]
            cms_val = row[cms_col - 1] if cms_col else None
            mrn_val = row[mrn_col - 1] if mrn_col else None
            cms_int = None
            try:
                if cms_val is not None and str(cms_val).strip() != "":
                    cms_int = int(float(str(cms_val).strip()))
            except Exception:
                cms_int = None

            ineligible, reason = cpt_is_ineligible(cpt_val)
            if ineligible and cms_int == 1:
                msg = f"OASCAPHS Row {r}: CPT {cpt_val} ineligible ({reason})"
                cpt_ineligible_rows.append((r, cpt_val, reason, mrn_val, cms_val))

                row_issues.append(
                    {
                        "row": r,
                        "mrn": mrn_val,
                        "cms": cms_val,
                        "issue_type": "CPT Ineligible",
                        "description": f"CPT {cpt_val} ineligible ({reason})",
                    }
                )
                issues.append(msg)
    else:
        issues.append("CPT column missing in OASCAPHS for ineligibility check")

    # ISSUES section
    report_lines.append("<h2>ISSUES FOUND</h2>")

    # Display row-based issues in table format
    if row_issues:
        report_lines.append("<table class='excel-style' style='font-size: 0.85em;'>")
        report_lines.append(
            "<tr><th style='background-color: #000; color: #fff; padding: 4px 8px;'>ROW</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>MRN</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>CMS</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>ISSUE TYPE</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>DESCRIPTION</th></tr>"
        )
        for issue in row_issues:
            mrn_display = issue.get("mrn") if issue.get("mrn") is not None else ""
            cms_display = issue.get("cms") if issue.get("cms") is not None else ""
            report_lines.append(
                f"<tr><td style='padding: 3px 8px;'>{issue['row']}</td><td style='padding: 3px 8px;'>{mrn_display}</td><td style='padding: 3px 8px;'>{cms_display}</td><td style='padding: 3px 8px;'>{issue['issue_type']}</td><td style='padding: 3px 8px;'>{issue['description']}</td></tr>"
            )
        report_lines.append("</table>")

    # Display general/non-row issues as list
    non_row_issues = [
        iss
        for iss in issues
        if not any(
            iss.startswith(f"OASCAPHS Row") or iss.startswith(f"UPLOAD Row")
            for iss in [iss]
        )
    ]
    if non_row_issues:
        report_lines.append("<h3>General Issues</h3>")
        report_lines.append("<ul>")
        for issue in non_row_issues:
            report_lines.append(f"<li>{issue}</li>")
        report_lines.append("</ul>")

    if not row_issues and not non_row_issues:
        report_lines.append("<p>No issues found</p>")

    # CPT ineligible summary

    if cpt_ineligible_rows:
        report_lines.append("<h2>CPT INELIGIBLE SUMMARY</h2>")
        report_lines.append(
            "<p><em>Note: Some ineligible CPT codes are expected to be in the non-report (CMS=2) section!</em></p>"
        )
        report_lines.append(
            f"<p><strong>Total ineligible CPT rows found: {len(cpt_ineligible_rows)}</strong></p>"
        )
        report_lines.append("<table class='excel-style' style='font-size: 0.85em;'>")
        report_lines.append(
            "<tr><th style='background-color: #000; color: #fff; padding: 4px 8px;'>ROW</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>MRN</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>CMS</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>CPT</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>REASON</th></tr>"
        )
        for r, cpt, reason, mrn, cms in cpt_ineligible_rows:
            mrn_display = mrn if mrn is not None else ""
            cms_display = cms if cms is not None else ""
            report_lines.append(
                f"<tr><td style='padding: 3px 8px;'>{r}</td><td style='padding: 3px 8px;'>{mrn_display}</td><td style='padding: 3px 8px;'>{cms_display}</td><td style='padding: 3px 8px;'>{cpt}</td><td style='padding: 3px 8px;'>{reason}</td></tr>"
            )
        report_lines.append("</table>")

    # INVALID ADDRESSES section
    # audit addresses using google's package
    invalid_addresses, noted_addresses = check_address(
        sheet, addr1_col, city_col, state_col, zip_col, mrn_col, cms_col, em_col
    )
    if invalid_addresses:
        report_lines.append("<h2>INVALID ADDRESSES FOUND</h2>")
        report_lines.append("<table class='excel-style' style='font-size: 0.85em;'>")
        report_lines.append(
            "<tr><th style='background-color: #000; color: #fff; padding: 4px 8px;'>ROW</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>MRN</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>CMS</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>E/M</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>STREET</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>CITY</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>STATE</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>ZIP</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>REASON</th></tr>"
        )
        for address in invalid_addresses:
            # Parse format: "Row: 5 - MRN: '123' - CMS: '1' - E/M: 'E' - ADDRESS: '{'country_code': 'US', ...}' - REASON: 'Invalid state'"
            parts = address.split(" - ")
            row_num = parts[0].replace("Row: ", "").strip()
            mrn_val = parts[1].replace("MRN: ", "").strip("'")
            cms_val = parts[2].replace("CMS: ", "").strip("'")
            em_val = parts[3].replace("E/M: ", "").strip("'")

            # Replace 'None' with empty string for display
            if mrn_val == "None":
                mrn_val = ""
            if cms_val == "None":
                cms_val = ""
            if em_val == "None":
                em_val = ""

            # Extract dictionary from ADDRESS part
            addr_dict_str = parts[4].replace("ADDRESS: ", "").strip("'")
            try:
                addr_dict = eval(addr_dict_str)
                street = addr_dict.get("street_address") or ""
                city = addr_dict.get("city") or ""
                state = addr_dict.get("country_area") or ""
                zip_code = addr_dict.get("postal_code") or ""

                # Clean up "None" strings
                if street == "None":
                    street = ""
                if city == "None":
                    city = ""
                if state == "None":
                    state = ""
                if zip_code == "None":
                    zip_code = ""
            except:
                street = city = state = zip_code = ""

            reason_text = (
                parts[5].replace("REASON: ", "").strip("'") if len(parts) > 5 else ""
            )
            report_lines.append(
                f"<tr><td style='padding: 3px 8px;'>{row_num}</td><td style='padding: 3px 8px;'>{mrn_val}</td><td style='padding: 3px 8px;'>{cms_val}</td><td style='padding: 3px 8px;'>{em_val}</td><td style='padding: 3px 8px;'>{street}</td><td style='padding: 3px 8px;'>{city}</td><td style='padding: 3px 8px;'>{state}</td><td style='padding: 3px 8px;'>{zip_code}</td><td style='padding: 3px 8px;'>{reason_text}</td></tr>"
            )
        report_lines.append("</table>")

    # possibly problematic addresses
    if noted_addresses:
        report_lines.append("<h2>PROBLEMATIC ADDRESSES FOUND</h2>")
        report_lines.append("<table class='excel-style' style='font-size: 0.85em;'>")
        report_lines.append(
            "<tr><th style='background-color: #000; color: #fff; padding: 4px 8px;'>ROW</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>MRN</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>CMS</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>E/M</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>ADDRESS</th><th style='background-color: #000; color: #fff; padding: 4px 8px;'>ISSUE(S)</th></tr>"
        )
        for address in noted_addresses:
            # Parse format: "Row: 5 - MRN: '123' - CMS: '1' - E/M: 'E' - ADDRESS: '123 Main St' - REASON(s): 'city, state'"
            parts = address.split(" - ")
            row_num = parts[0].replace("Row: ", "").strip()
            mrn_val = parts[1].replace("MRN: ", "").strip("'")
            cms_val = parts[2].replace("CMS: ", "").strip("'")
            em_val = parts[3].replace("E/M: ", "").strip("'")

            # Replace 'None' with empty string for display
            if mrn_val == "None":
                mrn_val = ""
            if cms_val == "None":
                cms_val = ""
            if em_val == "None":
                em_val = ""

            addr_text = parts[4].replace("ADDRESS: ", "").strip("'")
            if addr_text == "None":
                addr_text = ""

            reason_text = (
                parts[5].replace("REASON(s): ", "").strip("'") if len(parts) > 5 else ""
            )
            report_lines.append(
                f"<tr><td style='padding: 3px 8px;'>{row_num}</td><td style='padding: 3px 8px;'>{mrn_val}</td><td style='padding: 3px 8px;'>{cms_val}</td><td style='padding: 3px 8px;'>{em_val}</td><td style='padding: 3px 8px;'>{addr_text}</td><td style='padding: 3px 8px;'>{reason_text}</td></tr>"
            )
        report_lines.append("</table>")

    report_lines.append("<hr>")
    report_lines.append(
        "<p style='text-align: center;'><strong>END OF REPORT</strong></p>"
    )
    report_lines.append("</div>")
    report_lines.append("</body>")
    report_lines.append("</html>")

    return report_lines, issues


def _build_html_header(file_path, version, audit_id=None):
    """
    Build the HTML header section (reusable for both success and failure reports)
    """
    tor = datetime.datetime.now()
    time_of_report = tor.strftime("%m/%d/%Y %H:%M:%S")

    modified_ts = "N/A"
    try:
        modified_ts = datetime.datetime.fromtimestamp(
            os.path.getmtime(file_path)
        ).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        pass
    basefname = os.path.basename(file_path)
    base_before_hash = basefname.split("#", 1)[0]

    header_lines = []
    header_lines.append("<!DOCTYPE html>")
    header_lines.append("<html>")
    header_lines.append("<head>")
    header_lines.append("    <meta charset='UTF-8'>")
    title_prefix = "Failed Audit" if audit_id is None else "Audit Report"
    header_lines.append(f"    <title>{title_prefix} - {base_before_hash}</title>")
    header_lines.append("    <style>")

    # Load CSS from external file
    css_path = os.path.join(os.path.dirname(__file__), "audit_report.css")
    try:
        with open(css_path, "r", encoding="utf-8") as css_file:
            for line in css_file:
                header_lines.append(f"        {line.rstrip()}")
    except FileNotFoundError:
        # Fallback to basic styling if CSS file not found
        header_lines.append("        body { font-family: sans-serif; }")

    header_lines.append("    </style>")
    header_lines.append("</head>")
    header_lines.append("<body>")
    header_lines.append("<div class='report-container'>")

    # Updated header presentation
    header_lines.append("<div style='padding-bottom: 15px; margin-bottom: 20px;'>")
    header_lines.append(f"<h1 style='margin: 0 0 5px 0;'>OAS-CAHPS Audit Report</h1>")
    header_lines.append(
        f"<p style='margin: 0; color: #bdc3c7; font-size: 0.85em;'><a href='https://tylercbrock.com' style='color: inherit; text-decoration: none;'>Auditor</a> v{version}</p>"
    )
    header_lines.append("</div>")

    # Info grid layout
    header_lines.append(
        "<div style='display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 15px;'>"
    )
    header_lines.append(f"<div><strong>Client:</strong> {base_before_hash}</div>")
    if audit_id is None:
        header_lines.append("<div><strong>Audit ID:</strong> N/A (audit failed)</div>")
    else:
        header_lines.append(f"<div><strong>Audit ID:</strong> {audit_id}</div>")
    header_lines.append(f"<div><strong>Report Date:</strong> {time_of_report}</div>")
    header_lines.append(f"<div><strong>File Modified:</strong> {modified_ts}</div>")
    header_lines.append("</div>")

    header_lines.append("<hr>")

    return header_lines


def save_report(file_path, report_lines, failure_reason="", version="0.0-alpha"):
    """
    Write report to .html file in audits directory
    """
    # --- Write report to .html file ---
    base_name = os.path.splitext(file_path)[0]
    report_file = base_name + ".html"

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
        else:
            # Build failure report using the helper function
            failure_html = _build_html_header(file_path, version, audit_id=None)
            failure_html.append("<h2>Audit Failed</h2>")
            failure_html.append(f"<p>{report_lines}</p>")
            failure_html.append(
                f"<p><strong>Failure reason:</strong> {failure_reason}</p>"
            )
            failure_html.append("<hr>")
            failure_html.append(
                "<p style='text-align: center;'><strong>END OF REPORT</strong></p>"
            )
            failure_html.append("</div>")
            failure_html.append("</body>")
            failure_html.append("</html>")
            f.write("\n".join(failure_html))

    if not failure_reason:
        print(f"--- Audit complete. Report saved to {final_report_file}\n")
    else:
        print(
            f"--- Audit could not run on this file! Information saved to {final_report_file}\n"
        )

    # return the full file name and path in case it needs to be read again
    return final_report_file
