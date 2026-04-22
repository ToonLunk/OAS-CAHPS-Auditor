import os
import sys
import datetime
import math
import base64
from tqdm import tqdm
from dotenv import load_dotenv

from requests import head
from audit_lib_funcs import check_address, check_pop_upload_email_consistency, count_nonempty_rows_after_header, collect_lookup_candidates, build_person_search_urls, check_email_quality_all_rows


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
    addr2_col,
    city_col,
    state_col,
    zip_col,
    cms_col=None,
    em_col=None,
    find_frame_inel_count=None,
    mrn_col=None,
    sid_col=None,
    sid_row_issues=None,
    inel_row_issues=None,
    sid_prefix=None,
    sid_registry_name=None,
    service_date_range=None,
    blank_date_row_issues=None,
    facility_matches=None,
):
    """
    Build the HTML audit report for saving as .html
    """

    # Track row-based issues separately for table display
    row_issues = []  # List of dicts: {row, mrn, cms, issue_type, description}

    basefname = os.path.basename(file_path)
    base_before_hash = basefname.split("#", 1)[0]

    try:
        after_hash = basefname.split("#", 1)[1]
    except IndexError:
        row_issues.append(
            {
                "row": "FILE",
                "mrn": None,
                "cms": None,
                "issue_type": "Filename Issue",
                "description": "Filename is missing '#' separator",
            }
        )
    else:
        # Remove extension
        name_part = os.path.splitext(after_hash)[0]

        # Extract month name
        month = name_part.split()[0].lower()

        months = {
            "january", "february", "march", "april", "may", "june",
            "july", "august", "september", "october", "november", "december",
        }

        if month not in months:
            row_issues.append(
                {
                    "row": "FILE",
                    "mrn": None,
                    "cms": None,
                    "issue_type": "Filename Issue",
                    "description": f"Invalid or misspelled month in filename: '{month}'",
                }
            )

    # Start HTML document with helper function
    report_lines = _build_html_header(file_path, version, audit_id, sid_prefix, service_date_range)
        
    # Add SID row issues if provided
    if sid_row_issues:
        row_issues.extend(sid_row_issues)
    
    # Add INEL row issues if provided
    if inel_row_issues:
        row_issues.extend(inel_row_issues)
    
    # Add blank date row issues if provided
    if blank_date_row_issues:
        row_issues.extend(blank_date_row_issues)

    # missing required headers -> issues
    if missing_req_headers:
        for header in missing_req_headers:
            issues.append(f"Missing REQUIRED Header: {header}!")

    # Header/Footer extracted values
    report_lines.append("<h2>DATA SUMMARY</h2>")
    report_lines.append("<div class='section-subheader'>OASCAPHS TAB ANALYSIS</div>")
    report_lines.append("<div class='three-column-flex'>")
    
    # Column 1: Patients Submitted
    report_lines.append("<div class='column'>")
    report_lines.append("<div class='label'>Patients Submitted (from header)</div>")
    if patients_submitted is not None:
        report_lines.append(f"<div class='value'>{patients_submitted}</div>")
    else:
        report_lines.append("<div class='value' style='color: orange;'>NOT FOUND</div>")
        issues.append("<strong>WARNING:</strong> SUBMITTED value not found in header")
    report_lines.append("</div>")
    
    # Column 2: Eligible Patients
    report_lines.append("<div class='column'>")
    report_lines.append("<div class='label'>Eligible Patients (from footer)</div>")
    if eligible_patients is not None:
        report_lines.append(f"<div class='value'>{eligible_patients}</div>")
    else:
        report_lines.append("<div class='value' style='color: orange;'>NOT FOUND</div>")
        issues.append("<strong>WARNING:</strong> EL value not found in footer")
    report_lines.append("</div>")
    
    # Column 3: Sample Size
    report_lines.append("<div class='column'>")
    report_lines.append("<div class='label'>Sample Size (from footer)</div>")
    if sample_size is not None:
        report_lines.append(f"<div class='value'>{sample_size}</div>")
    else:
        report_lines.append("<div class='value' style='color: orange;'>NOT FOUND</div>")
        issues.append("<strong>WARNING:</strong> SS value not found in footer")
    report_lines.append("</div>")
    
    report_lines.append("</div>")

    # OASCAPHS tab analysis
    report_lines.append("<div class='section-subheader'>CONTACT INFORMATION</div>")
    report_lines.append("<table class='data-table'>")
    report_lines.append(
        f"<tr><td>Rows with CMS INDICATOR = 1</td><td>{cms1_count}</td></tr>"
    )
    report_lines.append(f"<tr><td>Emails counted</td><td>{emails}</td></tr>")
    report_lines.append(f"<tr><td>Mailings counted</td><td>{mailings}</td></tr>")
    report_lines.append(f"<tr><td>Total of Emails + Mailings</td><td>{total_em}</td></tr>")
    report_lines.append(
        f"<tr><td>Non-Reported entries (CMS INDICATOR = 2)</td><td>{non_reported}</td></tr>"
    )
    report_lines.append("</table>")

    # count rows in INEL and FRAME (needed for validation checks)
    inel_count = None
    inel_highlighted_count = 0
    if "INEL" in wb.sheetnames:
        inel_sheet = wb["INEL"]
        # Import SERVICE_DATE_ALIASES and find_column_by_aliases here
        from audit_lib_funcs import SERVICE_DATE_ALIASES, find_column_by_aliases
        
        # Find service date column
        service_date_col, header_row = find_column_by_aliases(inel_sheet, SERVICE_DATE_ALIASES)
        start_row = header_row + 1 if header_row else 2
        
        # Count rows, but skip ones with highlighted service dates
        # Optimized: Load all rows at once instead of per-row iter_rows calls
        inel_count = 0
        all_rows = list(inel_sheet.iter_rows(min_row=start_row, max_row=inel_sheet.max_row, values_only=False))
        
        for row_offset, row_cells in enumerate(tqdm(all_rows, desc="Processing INEL rows", disable=len(all_rows) < 1000)):
            row_idx = start_row + row_offset
            # Get row values
            row = [cell.value for cell in row_cells]
            
            # Check if row has any data
            if not any(cell is not None and str(cell).strip() != "" for cell in row):
                continue
            
            # If we found a service date column, check if it's highlighted
            skip_row = False
            if service_date_col:
                try:
                    cell = row_cells[service_date_col - 1]  # service_date_col is 1-indexed
                    # Check if cell has a fill color (is highlighted)
                    if cell.fill and cell.fill.start_color:
                        color_index = cell.fill.start_color.index
                        if color_index and color_index != '00000000' and color_index != 'FFFFFFFF':
                            skip_row = True
                except (AttributeError, IndexError):
                    pass
            
            if not skip_row:
                inel_count += 1
            else:
                inel_highlighted_count += 1

    frame_inel_count = None
    if "FRAME" in wb.sheetnames and find_frame_inel_count is not None:
        frame_sheet = wb["FRAME"]
        try:
            frame_inel_count = find_frame_inel_count(frame_sheet)
        except Exception:
            frame_inel_count = None

    # VALIDATION CHECKS
    report_lines.append("<h2>VALIDATION SUMMARY</h2>")

    # Tab counts in table format
    report_lines.append("<div class='section-subheader'>INELIGIBLE PATIENTS</div>")
    report_lines.append("<table class='data-table'>")
    if inel_count is not None:
        report_lines.append(f"<tr><td>Patients in INEL tab</td><td>{inel_count}</td></tr>")
        report_lines.append(
            f"<tr><td>Patients with ineligible service dates</td><td>{inel_highlighted_count}</td></tr>"
        )
    else:
        issues.append("INEL tab missing")

    if frame_inel_count is not None:
        report_lines.append(
            f"<tr><td>6-month repeats</td><td>{frame_inel_count}</td></tr>"
        )

    total_inel_combined = (inel_count or 0) + (frame_inel_count or 0)
    if patients_submitted is not None:
        report_lines.append(
            f"<tr><td>Total Ineligible Patients</td><td>{total_inel_combined}</td></tr>"
        )
    report_lines.append("</table>")

    # Validation checks in table format
    report_lines.append("<div class='section-subheader'>ADDITIONAL VALIDATIONS</div>")
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
            f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>"
        )
        issues.append(issue_msg)
    else:
        report_lines.append(
            "<tr><td>E/M total matches Sample Size</td><td style='color: #28a745;'>✓</td></tr>"
        )

    # Check 3: Submitted matches POP tab
    if "POP" in wb.sheetnames and patients_submitted is not None:
        pop_sheet = wb["POP"]
        pop_rows = count_nonempty_rows_after_header(pop_sheet)
        TOL = 4
        expected_submitted = pop_rows - inel_highlighted_count
        if abs(patients_submitted - expected_submitted) > TOL:
            issue_msg = (
                f"Potential patient # mismatch: header says {patients_submitted} patients were submitted, "
                f"but POP has {pop_rows} rows and INEL has {inel_highlighted_count} highlighted service-date rows."
            )
            tooltip_text = (
                "This might not be a problem. The submitted count is expected to equal POP rows minus INEL rows "
                "with highlighted service dates, and some files have extra text/titles before the data starts. "
                "Please verify manually."
            )
            issue_msg_with_tooltip = f"{issue_msg} <span class='info-icon'>i<span class='tooltip'>{tooltip_text}</span></span>"
            report_lines.append(
                f"<tr><td>{issue_msg_with_tooltip}</td><td style='color: red;'>✗</td></tr>"
            )
            issues.append(f"<strong>WARNING:</strong> {issue_msg}")
        else:
            report_lines.append(
                "<tr><td>Submitted # matches POP tab #</td><td style='color: #28a745;'>✓</td></tr>"
            )
    else:
        issue_msg = (
            "<strong>WARNING:</strong> POP tab missing or Submitted value not found"
        )
        report_lines.append(
            f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>"
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
                f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>"
            )
            issues.append(issue_msg)
        else:
            report_lines.append(
                "<tr><td>UPLOAD and OASCAPHS row counts match</td><td style='color: #28a745;'>✓</td></tr>"
            )
    else:
        issue_msg = "UPLOAD tab missing"
        report_lines.append(
            f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>"
        )
        issues.append(issue_msg)

    # Check 5: UPLOAD tab has the correct columns (OASCAPHS minus ATT, LAG, ID, FD, LG, E/M)
    upload_only_cols = {"ATT", "LAG", "ID", "FD", "LG", "E/M"}
    if "UPLOAD" in wb.sheetnames:
        upload_sheet = wb["UPLOAD"]
        up_header_set = {
            cell.value for cell in next(upload_sheet.iter_rows(min_row=1, max_row=1))
            if cell.value is not None
        }
        expected_upload_cols = set(headers.keys()) - upload_only_cols - {None}
        missing_in_upload = expected_upload_cols - up_header_set
        extra_in_upload = up_header_set - expected_upload_cols
        if not missing_in_upload and not extra_in_upload:
            report_lines.append(
                "<tr><td>UPLOAD tab has correct columns</td><td style='color: #28a745;'>✓</td></tr>"
            )
        else:
            parts = []
            if missing_in_upload:
                parts.append(f"missing: {', '.join(sorted(missing_in_upload))}")
            if extra_in_upload:
                parts.append(f"extra: {', '.join(sorted(extra_in_upload))}")
            issue_msg = f"<strong>WARNING:</strong> UPLOAD tab column mismatch ({'; '.join(parts)})"
            report_lines.append(
                f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>"
            )
            issues.append(issue_msg)

    # Calculate estimated percentage if both values are available
    estimated_percentage = None
    if sample_size is not None and eligible_patients is not None and eligible_patients > 0:
        estimated_percentage = int(round((sample_size / eligible_patients) * 100, 0))

    # Check 5: SID validation
    if sid_row_issues is not None:
        if not sid_row_issues:
            report_lines.append(
                "<tr><td>SIDs present and in order</td><td style='color: #28a745;'>✓</td></tr>"
            )
        else:
            issue_types = set(issue['issue_type'] for issue in sid_row_issues)
            issue_summary = ', '.join(issue_types)
            issue_msg = f"<strong>WARNING:</strong> SID validation failed: {issue_summary} ({len(sid_row_issues)} issues)"
            report_lines.append(
                f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>"
            )
    else:
        issue_msg = "SID validation not performed"
        report_lines.append(
            f"<tr><td>{issue_msg}</td><td style='color: orange;'>⚠</td></tr>"
        )

    # Check 6: INEL REPEAT validation
    if inel_row_issues is not None:
        if not inel_row_issues:
            report_lines.append(
                "<tr><td>INEL tab REPEAT entries properly formatted</td><td style='color: #28a745;'>✓</td></tr>"
            )
        else:
            issue_types = set(issue['issue_type'] for issue in inel_row_issues)
            issue_summary = ', '.join(issue_types)
            issue_msg = f"<strong>WARNING:</strong> INEL REPEAT validation failed: {issue_summary} ({len(inel_row_issues)} issues)"
            report_lines.append(
                f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>"
            )
    else:
        if "INEL" in wb.sheetnames:
            issue_msg = "INEL REPEAT validation not performed"
            report_lines.append(
                f"<tr><td>{issue_msg}</td><td style='color: orange;'>⚠</td></tr>"
            )

    # Check 7: Eligible + INEL = Submitted math check
    if patients_submitted is not None and eligible_patients is not None and inel_count is not None:
        math_total = eligible_patients + total_inel_combined
        if math_total != patients_submitted:
            issue_msg = (
                f"<strong>WARNING:</strong> Math error: "
                f"Eligible <strong style='color: red;'>{eligible_patients}</strong> + "
                f"INEL <strong style='color: red;'>{total_inel_combined}</strong> = "
                f"<strong style='color: red;'>{math_total}</strong>, "
                f"but Submitted = <strong style='color: red;'>{patients_submitted}</strong>"
            )
            tooltip_text = (
                "This may be expected in some cases, like if there are multiple facilities in the file and the submitted count only includes patients from one facility. Please verify manually."
            )
            issue_msg_with_tooltip = f"{issue_msg} <span class='info-icon'>i<span class='tooltip'>{tooltip_text}</span></span>"
            report_lines.append(
                f"<tr><td style='background-color: #fff3cd;'>{issue_msg_with_tooltip}</td><td style='color: red;'>✗</td></tr>"
            )
            issues.append(f"<strong>WARNING:</strong> Math error: Eligible ({eligible_patients}) + Combined INEL ({total_inel_combined}) = {math_total}, but Submitted = {patients_submitted}")
        else:
            report_lines.append(
                f"<tr><td>Eligible + INEL = Submitted ({eligible_patients} + {total_inel_combined} = {patients_submitted})</td><td style='color: #28a745;'>✓</td></tr>"
            )

    report_lines.append("</table>") # based on month (November=orange, December=green, etc.)
    qtr_header_color = "#2dbd69"  # Default green
    if service_date_range:
        try:
            date_parts = service_date_range.split(" - ")
            if len(date_parts) == 2:
                start_date = datetime.datetime.strptime(date_parts[0].strip(), "%m/%d/%Y")
                # November (11) is orange, December (12) is green, January (1) is orange, etc.
                # November is month 11, which is odd, so odd months are orange
                month = start_date.month
                if month % 2 == 1:  # Odd months: Jan(1), Mar(3), May(5), Jul(7), Sep(9), Nov(11)
                    qtr_header_color = "#ec8038"  # Orange
                # Even months stay green
        except (ValueError, AttributeError):
            pass  # Keep default green
    
    report_lines.append("<h2>ESTIMATED LOG SHEET LINE"
        " <span class='info-icon'>i<span class='tooltip'>"
        "<b>Data sources:</b><br>"
        "SID — from header<br>"
        "Client — SID registry (SIDs.csv) or filename<br>"
        "Non-Reported — CMS INDICATOR = 2 count<br>"
        "Emails / Mailings — E/M column (CMS=1 rows)<br>"
        "Selection % — Sample Size ÷ Eligible<br>"
        "Submitted — from header<br>"
        "Eligible — from footer (EL)<br>"
        "Sample Size — from footer (SS)"
        "</span></span></h2>")
    c = 'text-align: center;'
    report_lines.append(f"<table class='excel-style' style='--header-color: {qtr_header_color};'>")
    report_lines.append("<tr>")
    report_lines.append(
        f"<th style='background-color: {qtr_header_color}; {c}'>SID</th>" 
        f"<th style='background-color: {qtr_header_color}; width: 30%;'>CLIENT NAME</th>" 
        f"<th style='background-color: {qtr_header_color}; {c}'>NON REPORTED</th>" 
        f"<th style='background-color: {qtr_header_color}; {c}'>REPORTED EMAILS</th>" 
        f"<th style='background-color: {qtr_header_color}; {c}'>MAILINGS TOTAL</th>"
    )
    report_lines.append(
        f"<th style='background-color: {qtr_header_color}; {c}'>EST. %</th>"
        f"<th style='background-color: {qtr_header_color}; {c}'># PATIENTS SUBMITTED</th>"
        f"<th style='background-color: {qtr_header_color}; {c}'>ELIGIBLE PATIENTS</th>"
        f"<th style='background-color: {qtr_header_color}; {c}'>SAMPLE SIZE</th>"
    )
    report_lines.append("</tr>")
    report_lines.append("<tr>")
    report_lines.append(f"<td style='{c}'>{sid_prefix if sid_prefix else 'N/A'}</td>")
    # Use registry name if available, otherwise fall back to file name
    client_name_display = sid_registry_name if sid_registry_name else base_before_hash
    report_lines.append(f"<td>{client_name_display}</td>")
    report_lines.append(f"<td style='{c}'>{non_reported if non_reported is not None else 'N/A'}</td>")
    report_lines.append(f"<td style='{c}'>{emails if emails is not None else 'N/A'}</td>")
    report_lines.append(f"<td style='{c}'>{mailings if mailings is not None else 'N/A'}</td>")
    report_lines.append(f"<td style='{c}'>~{estimated_percentage}%</td>" if estimated_percentage is not None else f"<td style='{c}'>N/A</td>")
    report_lines.append(f"<td style='{c}'>{patients_submitted if patients_submitted is not None else 'N/A'}</td>")
    report_lines.append(f"<td style='{c}'>{eligible_patients if eligible_patients is not None else 'N/A'}</td>")
    report_lines.append(f"<td style='{c}'>{sample_size if sample_size is not None else 'N/A'}</td>")
    report_lines.append("</tr>")
    report_lines.append("</table>")
    
    # Add SID client name comparison if available
    if sid_prefix and sid_registry_name:
        report_lines.append("<h3 style='margin-top: 15px; margin-bottom: 5px;'>SID Registry Check"
            " <span class='info-icon'>i<span class='tooltip'>"
            "SIDs.csv contains the list of client names matched to SID codes. "
            "If facility/site name columns are found in the POP tab, they are shown below. "
            "Download the latest version from the "
            "<a href='https://jlm353-my.sharepoint.com/:f:/g/personal/dcdata_jlm-solutions_com/IgBhYR7tt6YTRbgNTDEh9M7xAc5HSCC3KSaJt6ImfJV65kg?e=hKp0ZU' "
            "style='color: #5dade2;' target='_blank'>shared OneDrive folder</a> "
            "and place it in your installation directory (default: C:\\OAS-CAHPS-Auditor)."
            "</span></span></h3>")
        report_lines.append("<table class='excel-style' style='font-size: 0.9em;'>")
        report_lines.append("<tr>")
        report_lines.append("<th style='background-color: #000; color: #fff;'>SID</th>")
        report_lines.append("<th style='background-color: #000; color: #fff;'>Client Name (from file)</th>")
        report_lines.append("<th style='background-color: #000; color: #fff;'>Client Name (from registry)</th>")
        report_lines.append("</tr>")
        report_lines.append("<tr>")
        report_lines.append(f"<td>{sid_prefix}</td>")
        
        # Normalize both names for comparison
        import re
        # Remove date patterns like "11/1" or "- 11/1" from the end
        # Keeps location names: "Name - Location - 11/1" becomes "Name - Location"
        normalized_registry = re.sub(r'\s*-?\s*\d{1,2}/\d{1,2}(?:/\d{2,4})?\s*$', '', sid_registry_name).strip().lower()
        normalized_filename = base_before_hash.strip().lower()
        
        # Compare normalized names (case-insensitive) and set color
        match_color = "#27ae60" if normalized_registry == normalized_filename else "#e74c3c"  # Green if match, red if not
        
        report_lines.append(f"<td style='color: {match_color}; font-weight: 600;'>{base_before_hash}</td>")
        report_lines.append(f"<td style='color: {match_color}; font-weight: 600;'>{sid_registry_name}</td>")
        report_lines.append("</tr>")
        report_lines.append("</table>")

        # Show facility/location columns found in FRAME/POP tabs (collapsible)
        fac_matches = facility_matches or []
        if fac_matches:
            count_label = f"{len(fac_matches)} column{'s' if len(fac_matches) != 1 else ''} found"
            report_lines.append(
                f"<details style='margin-top: 8px; font-size: 0.9em;'>"
                f"<summary style='cursor: pointer; font-weight: 600;'>"
                f"Facility / Location columns ({count_label})</summary>"
            )
            for match in fac_matches:
                col_name = match.get('header_name', 'N/A')
                tab_name = match.get('tab', 'POP')
                is_delimited = match.get('is_delimited', False)
                delim_note = f" <span style='color:#888; font-weight:400;'>(pipe-delimited)</span>" if is_delimited else (
                    f" <span style='color:#888; font-weight:400;'>(comma-delimited)</span>" if match.get('delimiter') == ',' else '')
                values = match.get('values', [])
                val_count = len(values)
                report_lines.append(
                    f"<div style='margin-top: 8px; margin-bottom: 2px; font-weight: 600;'>"
                    f"{tab_name} &rarr; <em>{col_name}</em>{delim_note} &mdash; "
                    f"{val_count} unique value{'s' if val_count != 1 else ''}"
                    f"</div>"
                )
                if values:
                    report_lines.append("<table class='data-table' style='margin-top: 2px;'>")
                    report_lines.append("<tr><th>#</th><th>Value</th></tr>")
                    for i, val in enumerate(values, start=1):
                        report_lines.append(f"<tr><td style='width: 40px; text-align: center;'>{i}</td><td>{val}</td></tr>")
                    report_lines.append("</table>")
                else:
                    report_lines.append("<p style='margin: 2px 0; color: #888;'><em>No values found</em></p>")
            report_lines.append("</details>")

    else:
        # Show that SID registry check couldn't be performed
        report_lines.append("<h3 style='margin-top: 15px; margin-bottom: 5px;'>SID Registry Check"
            " <span class='info-icon'>i<span class='tooltip'>"
            "SIDs.csv contains the list of client names matched to SID codes. "
            "Download the latest version from the "
            "<a href='https://jlm353-my.sharepoint.com/:f:/g/personal/dcdata_jlm-solutions_com/IgBhYR7tt6YTRbgNTDEh9M7xAc5HSCC3KSaJt6ImfJV65kg?e=hKp0ZU' "
            "style='color: #5dade2;' target='_blank'>shared OneDrive folder</a> "
            "and place it in your installation directory (default: C:\\OAS-CAHPS-Auditor)."
            "</span></span></h3>")
        report_lines.append("<p style='color: #000; margin: 5px 0;'>")
        if not sid_prefix:
            report_lines.append("⚠ Unable to perform SID registry check: SID prefix not found in file")
        else:
            report_lines.append("⚠ Unable to perform SID registry check: Matching SID not found in registry")
        report_lines.append("</p>")

    # DATA QUALITY VALIDATION SECTION
    from audit_lib_funcs import column_validations

    issues, row_issues = column_validations(
        sheet, headers, mrn_col, cms_col, em_col, issues, row_issues
    )

    # Email quality / suspicious-email scan
    email_col = headers.get("EMAIL ADDRESS")
    cms1_email_quality, cms2_email_quality = check_email_quality_all_rows(
        sheet, email_col, mrn_col, cms_col
    )
    # CMS=1 potentially invalid emails go into the main issues table
    for eq in cms1_email_quality:
        desc = "; ".join(eq["warnings"])
        row_issues.append(
            {
                "row": eq["row"],
                "mrn": eq["mrn"],
                "cms": eq["cms"],
                "issue_type": "Potentially Invalid Email",
                "description": f"'{eq['email']}' — {desc}",
            }
        )
        issues.append(f"OASCAPHS Row {eq['row']}: Potentially invalid email '{eq['email']}' — {desc}")

    # 1. Surgical Category Validation (OASCAPHS)
    report_lines.append("")
    cpt_col = headers.get("CPT")
    cat_col = headers.get("SURGICAL CATEGORY")
    if cpt_col and cat_col:
        from audit_lib_funcs import is_blank_row
        
        all_validation_rows = list(sheet.iter_rows(min_row=2, values_only=True))
        for r, row in enumerate(tqdm(all_validation_rows, desc="Validating surgical categories", disable=len(all_validation_rows) < 1000), start=2):
            if is_blank_row(row):
                continue
            cpt_val = row[cpt_col - 1]
            cat_val = row[cat_col - 1]
            expected = classify_cpt(str(cpt_val) if cpt_val else "")
            
            # Skip validation if both CPT and surgical category are blank
            cpt_is_blank = not cpt_val or str(cpt_val).strip() == ""
            cat_is_blank = not cat_val or str(cat_val).strip() == ""
            if cpt_is_blank and cat_is_blank:
                continue
                
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
    # Only run if UPLOAD tab exists AND row counts match — if counts differ,
    # Check 4 above already reported it; positional comparison would be meaningless.
    if "UPLOAD" in wb.sheetnames:
        upload_sheet = wb["UPLOAD"]
        _up_count = count_nonempty_rows(upload_sheet)
        _oas_count = count_nonempty_rows(sheet)
        if _up_count > 0 and _up_count == _oas_count:
            up_headers = {
                cell.value: idx
                for idx, cell in enumerate(
                    next(upload_sheet.iter_rows(min_row=1, max_row=1)), start=1
                )
            }
            oas_headers = headers
            ignore_cols = {"LG", "FD", "ID", "ATT", "LAG", "E/M"}
            common_cols = sorted(
                set(up_headers.keys()).intersection(oas_headers.keys())
                - ignore_cols
                - {None}
            )

            upload_rows = list(upload_sheet.iter_rows(min_row=2, values_only=True))
            oas_rows = list(sheet.iter_rows(min_row=2, values_only=True))

            def _norm(v):
                return "" if v is None else str(v).strip()

            for r_offset, (up_row, oas_row) in enumerate(zip(upload_rows, oas_rows)):
                r = r_offset + 2
                row_mismatches = []
                for col in common_cols:
                    up_idx = up_headers[col] - 1
                    oas_idx = oas_headers[col] - 1
                    up_val = up_row[up_idx] if up_idx < len(up_row) else None
                    oas_val = oas_row[oas_idx] if oas_idx < len(oas_row) else None
                    if _norm(up_val) != _norm(oas_val):
                        row_mismatches.append(
                            f"{col}: OASCAPHS='{oas_val}' UPLOAD='{up_val}'"
                        )
                if row_mismatches:
                    mrn_val = oas_row[mrn_col - 1] if mrn_col else None
                    cms_val = oas_row[cms_col - 1] if cms_col else None
                    row_issues.append(
                        {
                            "row": r,
                            "mrn": mrn_val,
                            "cms": cms_val,
                            "issue_type": "UPLOAD/OASCAPHS Mismatch",
                            "description": "; ".join(row_mismatches),
                        }
                    )
                    issues.append(
                        f"Row {r}: " + "; ".join(row_mismatches)
                    )

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

    # Check combined ineligible math — handled in ADDITIONAL VALIDATIONS table above

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
        report_lines.append("<details open>")
        report_lines.append(f"<summary>Issues ({len(row_issues)} found)</summary>")
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
        report_lines.append("</details>")

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
        report_lines.append("<h2>INELIGIBLE CPT CODES</h2>")
        report_lines.append(
            "<p><em>Note: Some ineligible CPT codes are expected to be in the non-report (CMS=2) section!</em></p>"
        )
        report_lines.append(
            f"<p><strong>Total ineligible CPT rows found: {len(cpt_ineligible_rows)}</strong></p>"
        )
        report_lines.append("<details open>")
        report_lines.append(
            f"<summary>Ineligible CPT Details ({len(cpt_ineligible_rows)} rows)</summary>"
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
        report_lines.append("</details>")

    # INVALID ADDRESSES section
    invalid_addresses, noted_addresses = check_address(
        sheet, addr1_col, city_col, state_col, zip_col, mrn_col, cms_col, em_col, addr2_col
    )
    if invalid_addresses:
        report_lines.append("<h2>INVALID ADDRESSES FOUND</h2>")
        report_lines.append("<details open>")
        report_lines.append(
            f"<summary>Invalid Address Details ({len(invalid_addresses)} found)</summary>"
        )
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
        report_lines.append("</details>")

    # possibly problematic addresses
    if noted_addresses:
        report_lines.append("<h2>PROBLEMATIC ADDRESSES FOUND</h2>")
        report_lines.append("<details open>")
        report_lines.append(
            f"<summary>Problematic Address Details ({len(noted_addresses)} found)</summary>"
        )
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
        report_lines.append("</details>")

    # PEOPLE-SEARCH LOOKUP SECTION
    candidates = collect_lookup_candidates(sheet, headers, mrn_col, cms_col)
    report_lines.append("<h2>CONTACT LOOKUP</h2>")
    if candidates:
        th = "<th style='background-color: #000; color: #fff; padding: 4px 8px;'>"
        report_lines.append("<details open>")
        report_lines.append(f"<summary>CMS=1 patients with contact issues ({len(candidates)} found)</summary>")
        report_lines.append("<table class='excel-style' style='font-size: 0.85em;'>")
        report_lines.append(
            f"<tr>{th}ROW</th>{th}MRN</th>{th}PATIENT NAME</th>{th}AGE</th>"
            f"{th}CITY, STATE</th>{th}REASON(S)</th>{th}SEARCH LINKS</th></tr>"
        )
        for c in candidates:
            mrn_disp  = c["mrn"]  if c["mrn"]  is not None else ""
            name_disp = c["name"] or "&mdash;"
            age_disp  = c["age"]  if c["age"]  is not None else ""
            location  = ", ".join(x for x in [c["city"], c["state"]] if x) or "&mdash;"
            reasons   = "; ".join(c["issues"])
            if c["mode"] == "lookup":
                name_for_lookup = c["name"] or ""
                urls = build_person_search_urls(name_for_lookup, c["city"], c["state"])
                primary_links = " &nbsp; ".join(
                    f"<a href='{url}' target='_blank' "
                    f"style='color:#2980b9;text-decoration:none;white-space:nowrap;'>{label}</a>"
                    for label, url in urls.items()
                )
                tokens = name_for_lookup.strip().split()
                if len(tokens) > 1:
                    # Assume LAST [MIDDLE...] FIRST → rearrange to FIRST [MIDDLE...] LAST
                    rearranged_name = " ".join(tokens[1:] + [tokens[0]])
                    rev_urls = build_person_search_urls(rearranged_name, c["city"], c["state"])
                    rev_links = " &nbsp; ".join(
                        f"<a href='{url}' target='_blank' "
                        f"style='color:#2980b9;text-decoration:none;white-space:nowrap;'>{label}</a>"
                        for label, url in rev_urls.items()
                    )
                    links_html = (
                        f"{primary_links}"
                        f"<details style='margin:0;padding:0;'>"
                        f"<summary style='cursor:pointer;font-size:0.9em;color:#888;list-style:none;padding:2px 4px 2px 0;margin:0;display:inline-block;'>"
                        f"&#9654; Try: {rearranged_name}</summary>"
                        f"<div style='margin:0;padding-left:6px;border-left:2px solid #ccc;'>{rev_links}</div>"
                        f"</details>"
                    )
                else:
                    links_html = primary_links
            else:
                links_html = "&mdash;"
            report_lines.append(
                f"<tr>"
                f"<td style='padding: 3px 8px;'>{c['row']}</td>"
                f"<td style='padding: 3px 8px;'>{mrn_disp}</td>"
                f"<td style='padding: 3px 8px;'>{name_disp}</td>"
                f"<td style='padding: 3px 8px;'>{age_disp}</td>"
                f"<td style='padding: 3px 8px;'>{location}</td>"
                f"<td style='padding: 3px 8px;'>{reasons}</td>"
                f"<td style='padding: 3px 8px;'>{links_html}</td>"
                f"</tr>"
            )
        report_lines.append("</table>")
        report_lines.append("</details>")

    # CMS=2 potentially invalid emails section (closed by default)
    if cms2_email_quality:
        report_lines.append("<h2>CMS=2 POTENTIALLY INVALID EMAILS</h2>")
        report_lines.append(
            "<p><em>These CMS=2 (non-report) rows have emails that may be opt-outs, "
            "placeholders, or disposable addresses. Use your best judgement.</em></p>"
        )
        report_lines.append("<details>")
        report_lines.append(
            f"<summary>CMS=2 Potentially Invalid Emails ({len(cms2_email_quality)} rows)</summary>"
        )
        th = "<th style='background-color: #000; color: #fff; padding: 4px 8px;'>"
        report_lines.append("<table class='excel-style' style='font-size: 0.85em;'>")
        report_lines.append(
            f"<tr>{th}ROW</th>{th}MRN</th>{th}CMS</th>{th}EMAIL</th>{th}REASON(S)</th></tr>"
        )
        for eq in cms2_email_quality:
            mrn_disp = eq["mrn"] if eq["mrn"] is not None else ""
            cms_disp = eq["cms"] if eq["cms"] is not None else ""
            reasons = "; ".join(eq["warnings"])
            report_lines.append(
                f"<tr>"
                f"<td style='padding: 3px 8px;'>{eq['row']}</td>"
                f"<td style='padding: 3px 8px;'>{mrn_disp}</td>"
                f"<td style='padding: 3px 8px;'>{cms_disp}</td>"
                f"<td style='padding: 3px 8px;'>{eq['email']}</td>"
                f"<td style='padding: 3px 8px;'>{reasons}</td>"
                f"</tr>"
            )
        report_lines.append("</table>")
        report_lines.append("</details>")

    report_lines.append("<hr>")
    report_lines.append(
        "<p style='text-align: center;'><strong>END OF REPORT</strong></p>"
    )
    report_lines.append("</div>")
    report_lines.append("</body>")
    report_lines.append("</html>")

    return report_lines, issues


def _build_html_header(file_path, version, audit_id=None, sid_prefix=None, service_date_range=None):
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

    icon_href = None
    icon_candidates = [
        os.path.join(os.path.dirname(__file__), "python-xxl.png"),
        os.path.join(os.path.dirname(__file__), "distribution", "python-xxl.png"),
    ]
    for icon_path in icon_candidates:
        if os.path.isfile(icon_path):
            try:
                with open(icon_path, "rb") as icon_file:
                    icon_b64 = base64.b64encode(icon_file.read()).decode("ascii")
                icon_href = f"data:image/png;base64,{icon_b64}"
                break
            except Exception:
                icon_href = None

    header_lines = []
    header_lines.append("<!DOCTYPE html>")
    header_lines.append("<html>")
    header_lines.append("<head>")
    header_lines.append("    <meta charset='UTF-8'>")
    title_prefix = "Failed Audit" if audit_id is None else "Audit Report"
    header_lines.append(f"    <title>{base_before_hash} - {title_prefix}</title>")
    if icon_href:
        header_lines.append(f"    <link rel='icon' type='image/png' href='{icon_href}'>")
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
    header_lines.append("<div style='padding-bottom: 15px; '>")
    header_lines.append("<div style='display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 5px; border-bottom: 2px solid #27ae60; padding-bottom: 5px;'>")
    header_lines.append(f"<h1 style='margin: 0; border: none; padding: 0;'>OAS-CAHPS Audit Report</h1>")
    if service_date_range:
        # Convert date range to long format (e.g., "November 16th, 2025 - November 30th, 2025")
        try:
            date_parts = service_date_range.split(" - ")
            if len(date_parts) == 2:
                start_date = datetime.datetime.strptime(date_parts[0].strip(), "%m/%d/%Y")
                end_date = datetime.datetime.strptime(date_parts[1].strip(), "%m/%d/%Y")
                
                # Helper function to add ordinal suffix
                def ordinal(day):
                    if 10 <= day % 100 <= 20:
                        suffix = 'th'
                    else:
                        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
                    return f"{day}{suffix}"
                
                start_long = f"{start_date.strftime('%B')} {ordinal(start_date.day)}, {start_date.year}"
                end_long = f"{end_date.strftime('%B')} {ordinal(end_date.day)}, {end_date.year}"
                long_date_range = f"{start_long} - {end_long}"
                
                header_lines.append(f"<div style='font-size: 1.2em; color: #34495e; font-weight: 500;'>{long_date_range}</div>")
            else:
                # Fallback to original if parsing fails
                header_lines.append(f"<div style='font-size: 1.2em; color: #34495e; font-weight: 500;'>{service_date_range}</div>")
        except (ValueError, AttributeError):
            # Fallback to original if parsing fails
            header_lines.append(f"<div style='font-size: 1.2em; color: #34495e; font-weight: 500;'>{service_date_range}</div>")
    header_lines.append("</div>")
    header_lines.append(
        f"<div style='display: flex; justify-content: space-between; align-items: center; margin: 0 0 5px 0; color: #bdc3c7; font-size: 0.85em;'>"
        f"<span><a href='https://tylercbrock.com' style='color: inherit; text-decoration: none;'>Auditor</a> v{version}</span>"
        f"<span><a href='https://github.com/ToonLunk/OAS-CAHPS-Auditor' style='color: inherit; text-decoration: none;'>Need Help?</a></span>"
        f"</div>"
    )
    header_lines.append("</div>")

    # Info grid layout
    header_lines.append(
        "<div style='display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 15px;'>"
    )
    client_display = f"{base_before_hash} ({sid_prefix})" if sid_prefix else base_before_hash
    header_lines.append(f"<div><strong>Client:</strong> {client_display}</div>")
    if audit_id is None:
        header_lines.append("<div><strong>Audit ID:</strong> N/A (audit failed)</div>")
    else:
        header_lines.append(f"<div><strong>Audit ID:</strong> {audit_id}</div>")
    header_lines.append(f"<div><strong>Report Date:</strong> {time_of_report}</div>")
    header_lines.append(f"<div><strong>File Modified:</strong> {modified_ts}</div>")
    header_lines.append("</div>")

    header_lines.append("<hr>")

    return header_lines


def save_report(file_path, report_lines, failure_reason="", version="0.0-alpha", service_date_range=None, update_info=None):
    """
    Write report to .html file in AUDITS directory
    Location depends on ORGANIZE_AUDITS_BY_DATE setting:
      - True: %LOCALAPPDATA%\OAS-CAHPS-Auditor\AUDITS\YEAR\MONTH\
      - False: Next to the audited file in AUDITS folder (default)
    """
    # --- Write report to .html file ---
    base_name = os.path.splitext(file_path)[0]
    report_file = base_name + ".html"
    
    # Load configuration
    load_dotenv()
    organize_by_date = os.getenv("ORGANIZE_AUDITS_BY_DATE", "false").lower() == "true"
    
    # Determine AUDITS directory location
    if organize_by_date:
        # NEW BEHAVIOR: Organize by year/month in LOCALAPPDATA
        filename = os.path.basename(file_path)
        month_folder = "UNKNOWN"
        year_folder = datetime.datetime.now().strftime("%Y")
        
        # Parse filename for month/year after #
        # Format: "ClientName# JANUARY OAS 2026.xlsx"
        if "#" in filename:
            # Get the part after # but before file extension
            parts_after_hash = filename.split("#")[1]
            name_part = os.path.splitext(parts_after_hash)[0].strip()  # Remove extension and trim
            
            # Look for " OAS " to split month and year
            if " OAS " in name_part:
                month_part, year_part = name_part.split(" OAS ", 1)
                month_part = month_part.strip().upper()
                year_part = year_part.strip()
                
                # Map month names to 3-letter abbreviations
                month_map = {
                    'JANUARY': 'JAN', 'FEBRUARY': 'FEB', 'MARCH': 'MAR', 'APRIL': 'APR',
                    'MAY': 'MAY', 'JUNE': 'JUN', 'JULY': 'JUL', 'AUGUST': 'AUG',
                    'SEPTEMBER': 'SEP', 'OCTOBER': 'OCT', 'NOVEMBER': 'NOV', 'DECEMBER': 'DEC'
                }
                if month_part in month_map:
                    month_folder = month_map[month_part]
                elif len(month_part) == 3:
                    month_folder = month_part
                
                # Extract year
                try:
                    year_folder = str(int(year_part))
                except ValueError:
                    pass
        
        # Build AUDITS directory in %LOCALAPPDATA%\OAS-CAHPS-Auditor\AUDITS\YEAR\MONTH\
        appdata = os.getenv("LOCALAPPDATA") or os.path.expanduser("~\\AppData\\Local")
        AUDITS_base = os.path.join(appdata, "OAS-CAHPS-Auditor", "AUDITS")
        AUDITS_dir = os.path.join(AUDITS_base, year_folder, month_folder)
    else:
        # OLD BEHAVIOR (DEFAULT): AUDITS folder next to the audited file
        base_dir = os.path.dirname(report_file) or "."
        AUDITS_dir = os.path.join(base_dir, "AUDITS")
    
    if failure_reason:
        AUDITS_dir = os.path.join(AUDITS_dir, "unable_to_run_audit")
    
    os.makedirs(AUDITS_dir, exist_ok=True)

    # Extract month name(s) from service date range for filename
    month_str = ""
    if service_date_range:
        try:
            # Parse dates from format "MM/DD/YYYY - MM/DD/YYYY"
            date_parts = service_date_range.split(" - ")
            if len(date_parts) == 2:
                start_date = datetime.datetime.strptime(date_parts[0].strip(), "%m/%d/%Y")
                end_date = datetime.datetime.strptime(date_parts[1].strip(), "%m/%d/%Y")
                
                # Get month names
                start_month = start_date.strftime("%b")  # Short month name (e.g., "Jan")
                end_month = end_date.strftime("%b")
                
                # Format: if same month, show once; if different, show range
                if start_month == end_month:
                    month_str = f"_{start_month}"
                else:
                    month_str = f"_{start_month}-{end_month}"
        except (ValueError, AttributeError):
            # If parsing fails, just don't add month to filename
            pass

    # timestamp and final filename
    name, ext = os.path.splitext(os.path.basename(report_file))
    timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    final_report_file = os.path.join(AUDITS_dir, f"{name}{month_str}_{timestamp}{ext}")

    # prevent accidental overwrite (very unlikely because of timestamp, but safe)
    if os.path.isfile(final_report_file):
        print(
            f"--- File already exists! This auditor will not overwrite files. If you wish to run a new audit on this file, please delete the previous audit:  {final_report_file}"
        )
        input("Press enter to exit: ")
        print("\n")
        sys.exit(99)

    # Inject update badge next to version text if an update is available
    if update_info and isinstance(report_lines, list):
        _badge = (
            f"<a href=\"{update_info['download_url']}\" "
            "style='margin-left:8px;background:#fffbe6;border:1px solid #ffe58f;"
            "padding:2px 8px;border-radius:3px;color:#8a6d3b;font-size:0.9em;"
            "text-decoration:none;font-weight:500;'"
            f" title='A newer version was available when this audit was generated'>"
            f"&#8595; Click here to download v{update_info['latest_version']}"
            "</a>"
        )
        for i, line in enumerate(report_lines):
            if "Auditor</a> v" in line:
                report_lines[i] = line.replace("</span>", f"{_badge}</span>", 1)
                break

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
