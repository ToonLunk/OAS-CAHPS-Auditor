# CAHPS Auditor
# Copyright (C) 2026 HST Pathways. All rights reserved.
# Originally developed by Tyler Brock. This copyright notice, including
# authorship credit to Tyler Brock, must be preserved in all copies,
# modifications, and derivative works of this software.
import ast
import os
import sys
import datetime
import base64
from tqdm import tqdm
from dotenv import load_dotenv

from audit_lib_funcs import check_address, check_pop_upload_email_consistency, count_nonempty_rows_after_header, collect_lookup_candidates, build_person_search_urls, check_email_quality_all_rows, OAS_SIDS_ONEDRIVE_LINK, HCAHPS_SIDS_ONEDRIVE_LINK

HCAPHS_HEADER_COLOR = "#be9bff"  # Pastel lavender used for all HCAHPS header accents


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
    exclu_count=None,
    exclu_row_issues=None,
    inel_count=None,
    inel_tab_row_issues=None,
    dup_count=None,
    frame_col_map=None,
    frame_invalid_addresses=None,
    frame_noted_addresses=None,
    audit_type="OAS",
    header_text="",
):
    """
    Build the HTML audit report for saving as .html
    """

    # Track row-based issues separately for table display
    row_issues = []  # List of dicts: {row, mrn, cms, issue_type, description}

    basefname = os.path.basename(file_path)
    base_before_hash = basefname.split("#", 1)[0]
    main_tab_name = "CMS" if audit_type == "HCAHPS" else "OASCAPHS"

    filename_year = None

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

        parts = name_part.split()

        months = {
            "january", "february", "march", "april", "may", "june",
            "july", "august", "september", "october", "november", "december",
        }

        # Find month anywhere in the parts (handles both "MONTH TYPE YEAR" and "TYPE MONTH YEAR")
        month = next((p.lower() for p in parts if p.lower() in months), None)

        # Convert month name to number (used for per-row date validation)
        _month_name_to_num = {
            "january": 1, "february": 2, "march": 3, "april": 4,
            "may": 5, "june": 6, "july": 7, "august": 8,
            "september": 9, "october": 10, "november": 11, "december": 12,
        }
        filename_month = _month_name_to_num.get(month) if month else None

        if month is None:
            row_issues.append(
                {
                    "row": "FILE",
                    "mrn": None,
                    "cms": None,
                    "issue_type": "Filename Issue",
                    "description": "Month not found in filename after '#'",
                }
            )
            issues.append("<strong>WARNING:</strong> Month not found in filename after '#' — date validation may be inaccurate")

        # Extract year (scan for any 4-digit number in 2000-2100 range)
        for _p in parts:
            try:
                _yr = int(_p)
                if 2000 <= _yr <= 2100:
                    filename_year = _yr
                    break
            except ValueError:
                pass

        import re as _re
        _is_biweekly = bool(_re.search(r'\b\d{1,2}\s*-\s*\d{1,2}\b', name_part))

        if filename_year is None and (audit_type == "OAS" or not _is_biweekly):
            row_issues.append(
                {
                    "row": "FILE",
                    "mrn": None,
                    "cms": None,
                    "issue_type": "Filename Issue",
                    "description": "Year not found in filename after '#'",
                }
            )
            issues.append("<strong>WARNING:</strong> Year not found in filename after '#' — date validation may be inaccurate")

    # Start HTML document with helper function
    report_lines = _build_html_header(file_path, version, audit_id, sid_prefix, service_date_range, audit_type)
        
    # Add SID row issues if provided
    if sid_row_issues:
        row_issues.extend(sid_row_issues)
    
    # Add INEL row issues if provided
    if inel_row_issues:
        row_issues.extend(inel_row_issues)

    # Add EXCLU row issues if provided
    if exclu_row_issues:
        row_issues.extend(exclu_row_issues)

    # Add INEL tab row issues if provided
    if inel_tab_row_issues:
        row_issues.extend(inel_tab_row_issues)
    
    # Add blank date row issues if provided
    if blank_date_row_issues:
        row_issues.extend(blank_date_row_issues)

    # missing required headers -> issues
    if missing_req_headers:
        for header in missing_req_headers:
            issues.append(f"Missing REQUIRED Header: {header}!")

    # Header/Footer extracted values
    report_lines.append("<h2>DATA SUMMARY</h2>")
    report_lines.append(f"<div class='section-subheader'>{'CMS TAB ANALYSIS' if audit_type == 'HCAHPS' else 'OASCAPHS TAB ANALYSIS'}</div>")
    report_lines.append("<div class='three-column-flex'>")
    
    # Column 1: Patients Submitted (OAS) / POP tab count (HCAHPS)
    if audit_type == "OAS":
        report_lines.append("<div class='column'>")
        report_lines.append("<div class='label'>Patients Submitted (from header)</div>")
        if patients_submitted is not None:
            report_lines.append(f"<div class='value'>{patients_submitted}</div>")
        else:
            report_lines.append("<div class='value' style='color: orange;'>NOT FOUND</div>")
            issues.append("<strong>WARNING:</strong> SUBMITTED value not found in header")
        report_lines.append("</div>")
    elif audit_type == "HCAHPS":
        pop_count = count_nonempty_rows_after_header(wb["POP"]) if "POP" in wb.sheetnames else None
        report_lines.append("<div class='column'>")
        report_lines.append("<div class='label'>Patients in POP tab</div>")
        if pop_count is not None:
            report_lines.append(f"<div class='value'>{pop_count}</div>")
        else:
            report_lines.append("<div class='value' style='color: orange;'>NOT FOUND</div>")
        report_lines.append("</div>")
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

    # VALIDATION CHECKS - buffered for collapsible wrapper (starts with CONTACT INFORMATION)
    _val_start_idx = len(report_lines)
    issues_before_val = len(issues)

    # OASCAPHS tab analysis
    report_lines.append("<div class='section-subheader'>CONTACT INFORMATION</div>")
    report_lines.append("<table class='data-table'>")
    report_lines.append(
        f"<tr><td>Rows with CMS INDICATOR = 1</td><td>{cms1_count}</td></tr>"
    )
    if audit_type == "OAS":
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

        # Find service date column (used for OAS only)
        service_date_col, header_row = find_column_by_aliases(inel_sheet, SERVICE_DATE_ALIASES)
        start_row = header_row + 1 if header_row else 2

        def _cell_is_highlighted(cell):
            try:
                if cell.fill and cell.fill.start_color:
                    color_index = cell.fill.start_color.index
                    if color_index and color_index != '00000000' and color_index != 'FFFFFFFF':
                        return True
            except (AttributeError, IndexError):
                pass
            return False

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
            
            skip_row = False
            if service_date_col and audit_type == "OAS":
                # OAS: skip rows with a highlighted service date (out-of-range)
                if _cell_is_highlighted(row_cells[service_date_col - 1]):
                    skip_row = True
            
            if not skip_row:
                inel_count += 1
            else:
                inel_highlighted_count += 1

    frame_inel_count = None
    if audit_type == "OAS" and "FRAME" in wb.sheetnames and find_frame_inel_count is not None:
        frame_sheet = wb["FRAME"]
        try:
            frame_inel_count = find_frame_inel_count(frame_sheet)
        except Exception:
            frame_inel_count = None

    # VALIDATION CHECKS - continued in same collapsible buffer

    # Tab counts in table format
    report_lines.append("<div class='section-subheader'>INELIGIBLE PATIENTS</div>")
    report_lines.append("<table class='data-table'>")
    if inel_count is not None:
        report_lines.append(f"<tr><td>Patients in INEL tab</td><td>{inel_count}</td></tr>")
        if audit_type == "OAS":
            # OAS: highlighted service dates indicate ineligible date range rows
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

    # EXCLU tab row count (HCAHPS only)
    if audit_type == "HCAHPS":
        if exclu_count is not None:
            report_lines.append(f"<tr><td>Patients in EXCLU tab</td><td>{exclu_count}</td></tr>")
        else:
            report_lines.append("<tr><td>EXCLU tab</td><td style='color: #888;'>Not found</td></tr>")
        if dup_count is not None:
            report_lines.append(f"<tr><td>Patients in DUP tab (DUP=D)</td><td>{dup_count}</td></tr>")
        if inel_row_issues is not None and len(inel_row_issues) > 0:
            report_lines.append(f"<tr><td>Same-day discharges (should be on INEL tab)</td><td>{len(inel_row_issues)}</td></tr>")

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

    # Check 1b: Verify EL (eligible patients footer value) against FRAME tab.
    # For HCAHPS, eligible rows are the numbered patient rows in FRAME.
    # Some files also include unnumbered rows and/or a lower sparse duplicate
    # block; those rows should not count toward EL.
    # This works whether or not partial sampling occurred (EL == SS or EL != SS).
    # OAS uses find_frame_inel_count() via its own path; this check is HCAHPS-only.
    if audit_type == "HCAHPS" and eligible_patients is not None and "FRAME" in wb.sheetnames:
        from audit_hcahps_funcs import count_frame_patients
        _frame_total, _frame_excluded = count_frame_patients(wb["FRAME"])
        if _frame_total is not None and _frame_excluded is not None:
            _frame_eligible = _frame_total - _frame_excluded
            if _frame_eligible == eligible_patients:
                _excluded_note = f", {_frame_excluded} non-eligible row(s) excluded" if _frame_excluded else ""
                report_lines.append(
                    f"<tr><td>Eligible count matches FRAME tab "
                    f"({_frame_total} total{_excluded_note} = {eligible_patients})</td>"
                    f"<td style='color: #28a745;'>✓</td></tr>"
                )
            else:
                _excluded_note = f" minus {_frame_excluded} non-eligible row(s)" if _frame_excluded else ""
                issue_msg = (
                    f"<strong>WARNING:</strong> Eligible patient count mismatch: "
                    f"footer says {eligible_patients}, but FRAME tab has "
                    f"{_frame_total} patients{_excluded_note} = {_frame_eligible}"
                )
                report_lines.append(f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>")
                issues.append(issue_msg)
    elif audit_type == "HCAHPS" and eligible_patients is not None:
        report_lines.append(
            f"<tr><td>FRAME tab not found; cannot verify eligible patient count "
            f"(footer: {eligible_patients})</td>"
            f"<td style='color: orange;'>⚠</td></tr>"
        )

    # Check 2: E/M total matches Sample Size (OAS only)
    if audit_type == "OAS":
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

    # Check 3: Submitted matches POP tab (OAS only)
    if audit_type == "OAS" and "POP" in wb.sheetnames and patients_submitted is not None:
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
        if audit_type == "OAS":
            issue_msg = (
                "<strong>WARNING:</strong> POP tab missing or Submitted value not found"
            )
            report_lines.append(
                f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>"
            )
            issues.append(issue_msg)

    # Check 4: UPLOAD and main tab row counts match
    if "UPLOAD" in wb.sheetnames:
        upload_sheet = wb["UPLOAD"]
        upload_rows = count_nonempty_rows(upload_sheet)
        oascaphs_rows = count_nonempty_rows(sheet)
        if upload_rows != oascaphs_rows:
            issue_msg = f"<strong>WARNING:</strong> UPLOAD mismatch: {upload_rows} rows vs {oascaphs_rows} rows in {main_tab_name}"
            report_lines.append(
                f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>"
            )
            issues.append(issue_msg)
        else:
            report_lines.append(
                f"<tr><td>UPLOAD and {main_tab_name} row counts match</td><td style='color: #28a745;'>✓</td></tr>"
            )
    else:
        issue_msg = "UPLOAD tab missing"
        report_lines.append(
            f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>"
        )
        issues.append(issue_msg)

    # Check 5: UPLOAD tab has the correct columns (main tab minus ATT, LAG, ID, FD, LG [and E/M for OAS])
    upload_only_cols = {"ATT", "LAG", "ID", "FD", "LG"} if audit_type == "HCAHPS" else {"ATT", "LAG", "ID", "FD", "LG", "E/M"}
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

    # Check 5b/5c: CMS=1 rows must have a SID value; CMS=2 rows must NOT have a SID value
    if sid_col and cms_col:
        _cms1_missing_sid = []
        _cms2_has_sid = []
        for _r, _row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if not any(_row):
                continue
            _cms_val = _row[cms_col - 1]
            _sid_val = _row[sid_col - 1]
            try:
                _cms_int = int(_cms_val)
            except (ValueError, TypeError):
                continue
            _sid_present = _sid_val is not None and str(_sid_val).strip() != ""
            _mrn_val = _row[mrn_col - 1] if mrn_col else None
            if _cms_int == 1 and not _sid_present:
                _cms1_missing_sid.append({
                    "row": _r, "mrn": _mrn_val, "cms": _cms_val,
                    "issue_type": "CMS=1 Missing SID",
                    "description": "CMS=1 row has no SID value",
                })
            elif _cms_int == 2 and _sid_present:
                _cms2_has_sid.append({
                    "row": _r, "mrn": _mrn_val, "cms": _cms_val,
                    "issue_type": "CMS=2 Has SID",
                    "description": f"CMS=2 row has unexpected SID value: '{_sid_val}'",
                })
        if not _cms1_missing_sid:
            report_lines.append(
                "<tr><td>All CMS=1 rows have a SID value</td><td style='color: #28a745;'>✓</td></tr>"
            )
        else:
            _msg = f"<strong>WARNING:</strong> {len(_cms1_missing_sid)} CMS=1 row(s) missing a SID value"
            report_lines.append(f"<tr><td>{_msg}</td><td style='color: red;'>✗</td></tr>")
            row_issues.extend(_cms1_missing_sid)
            issues.append(_msg)
        if not _cms2_has_sid:
            report_lines.append(
                "<tr><td>No CMS=2 rows have a SID value</td><td style='color: #28a745;'>✓</td></tr>"
            )
        else:
            _msg = f"<strong>WARNING:</strong> {len(_cms2_has_sid)} CMS=2 row(s) have an unexpected SID value"
            report_lines.append(f"<tr><td>{_msg}</td><td style='color: red;'>✗</td></tr>")
            row_issues.extend(_cms2_has_sid)
            issues.append(_msg)

    # Check 6: INEL REPEAT validation (OAS only - HCAHPS does not validate INEL REPEAT)
    if audit_type == "OAS":
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

    # Check 6b: Same-day discharge check (HCAHPS only)
    if audit_type == "HCAHPS" and "FRAME" in wb.sheetnames:
        if inel_row_issues is None:
            report_lines.append(
            "<tr><td>Same-day discharge check skipped &mdash; required FRAME column(s) not found (admit date, discharge date, or MRN)</td><td style='color: orange;'>⚠</td></tr>"
            )
        elif not inel_row_issues:
            report_lines.append(
                "<tr><td>No same-day discharges found in CMS tab</td><td style='color: #28a745;'>✓</td></tr>"
            )
        else:
            issue_msg = f"<strong>WARNING:</strong> {len(inel_row_issues)} same-day discharge(s) found in CMS tab - these patients should be on the INEL tab"
            report_lines.append(
                f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>"
            )

    # Check 7: Eligible + INEL = Submitted math check (OAS only - HCAHPS has no Submitted value)
    if audit_type == "OAS" and patients_submitted is not None and eligible_patients is not None and inel_count is not None:
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

    # EXCLU + INEL validation (HCAHPS only)
    if audit_type == "HCAHPS":
        if exclu_count is not None or inel_count is not None:
            _combined_exclu_inel = list(exclu_row_issues or []) + list(inel_tab_row_issues or [])
            if not _combined_exclu_inel:
                report_lines.append(
                    "<tr><td>EXCLU and INEL rows all marked with exclusion reason</td><td style='color: #28a745;'>✓</td></tr>"
                )
            else:
                _msg_parts = []
                if exclu_row_issues:
                    _msg_parts.append(f"{len(exclu_row_issues)} EXCLU row(s)")
                if inel_tab_row_issues:
                    _msg_parts.append(f"{len(inel_tab_row_issues)} INEL row(s)")
                issue_msg = f"<strong>WARNING:</strong> {' and '.join(_msg_parts)} missing a highlighted cell or red font"
                report_lines.append(
                    f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>"
                )
                issues.append(issue_msg)

    # Check: D.DATE is in ascending order for CMS=1 rows (HCAHPS only)
    if audit_type == "HCAHPS":
        _ddate_col = headers.get("D.DATE")
        if _ddate_col:
            _prev_date = None
            _out_of_order = []
            for _r, _row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
                if not any(_row):
                    continue
                _cms_val = _row[cms_col - 1] if cms_col else None
                try:
                    _cms_int = int(_cms_val) if _cms_val is not None else None
                except (TypeError, ValueError):
                    continue
                if _cms_int != 1:
                    continue
                _raw_date = _row[_ddate_col - 1]
                if _raw_date is None or str(_raw_date).strip() == "":
                    continue
                # Normalize to a comparable date value
                import datetime as _dt
                if isinstance(_raw_date, (_dt.date, _dt.datetime)):
                    _d = _raw_date if isinstance(_raw_date, _dt.date) else _raw_date.date()
                else:
                    try:
                        _d = _dt.datetime.strptime(str(_raw_date).strip(), "%m/%d/%Y").date()
                    except ValueError:
                        try:
                            _d = _dt.datetime.strptime(str(_raw_date).strip(), "%Y-%m-%d").date()
                        except ValueError:
                            continue
                if _prev_date is not None and _d < _prev_date:
                    _mrn_val = _row[mrn_col - 1] if mrn_col else None
                    _out_of_order.append((_r, _mrn_val, _d, _prev_date))
                _prev_date = _d
            if not _out_of_order:
                report_lines.append(
                    "<tr><td>D.DATE is in ascending order</td><td style='color: #28a745;'>✓</td></tr>"
                )
            else:
                issue_msg = f"<strong>WARNING:</strong> D.DATE is not in ascending order ({len(_out_of_order)} row(s) out of order)"
                report_lines.append(f"<tr><td>{issue_msg}</td><td style='color: red;'>✗</td></tr>")
                issues.append(issue_msg)
                for _r, _mrn, _d, _prev in _out_of_order:
                    row_issues.append({
                        "row": _r,
                        "mrn": _mrn,
                        "cms": 1,
                        "issue_type": "D.DATE Out of Order",
                        "description": f"{_d} is earlier than previous date {_prev}",
                    })

    report_lines.append("</table>")

    # Wrap the validation section in a collapsible <details>
    _val_section = report_lines[_val_start_idx:]
    _pass_count = sum(1 for line in _val_section if "\u2713" in line)
    _fail_count = sum(1 for line in _val_section if "\u2717" in line)
    _warn_count = sum(1 for line in _val_section if "\u26a0" in line)
    _has_issues = _fail_count > 0 or _warn_count > 0 or len(issues) > issues_before_val
    _label_parts = []
    if _pass_count:
        _label_parts.append(f"{_pass_count} passed")
    if _fail_count:
        _label_parts.append(f"{_fail_count} failed")
    if _warn_count:
        _label_parts.append(f"{_warn_count} warning{'s' if _warn_count != 1 else ''}")
    _summary_label = ", ".join(_label_parts) if _label_parts else "no checks run"
    _open_attr = " open" if _has_issues else ""
    del report_lines[_val_start_idx:]
    report_lines.append(f"<details{_open_attr} class='validation-summary-details'>")
    report_lines.append(f"<summary>VALIDATION SUMMARY: {_summary_label}</summary>")
    report_lines.extend(_val_section)
    report_lines.append("</details>")

    # Determine quarter header color (odd months = orange, even = green) — OAS only.
    # HCAHPS uses a fixed pastel lavender instead of alternating.
    qtr_header_color = "#2dbd69"  # Default green
    if service_date_range:
        try:
            date_parts = service_date_range.split(" - ")
            if len(date_parts) == 2:
                start_date = datetime.datetime.strptime(date_parts[0].strip(), "%m/%d/%Y")
                if start_date.month % 2 == 1:  # Odd months: Jan, Mar, May, Jul, Sep, Nov
                    qtr_header_color = "#ec8038"  # Orange
        except (ValueError, AttributeError):
            pass  # Keep default green
    if audit_type == "HCAHPS":
        qtr_header_color = HCAPHS_HEADER_COLOR  # Pastel lavender

    # FRAME column detection (HCAHPS only, shown right below validation summary)
    if audit_type == "HCAHPS" and frame_col_map is not None:
        _optional = {'Address 2'}
        _required = ['MRN', 'Admit Date', 'Discharge Date', 'Address 1', 'City', 'State', 'ZIP']
        _all_fields = _required + ['Address 2']
        _found_count = sum(1 for f in _all_fields if frame_col_map.get(f, (None, None))[0] is not None)
        report_lines.append(
            f"<details class='validation-summary-details'>"
            f"<summary style='font-weight: normal;'>"
            f"FRAME Column Detection &mdash; {_found_count}/{len(_all_fields)} fields found"
            f"</summary>"
        )
        report_lines.append("<table class='excel-style' style='font-size: 0.9em; margin-top: 6px;'>")
        report_lines.append(
            f"<tr>"
            f"<th style='background-color: {qtr_header_color};'>Expected Field</th>"
            f"<th style='background-color: {qtr_header_color};'>Matched Header</th>"
            f"<th style='background-color: {qtr_header_color};'>Found?</th>"
            f"</tr>"
        )
        for _field in _all_fields:
            _col_idx, _hdr = frame_col_map.get(_field, (None, None))
            _is_optional = _field in _optional
            _label = f"{_field} <em style='color:#888; font-weight:400;'>(optional)</em>" if _is_optional else _field
            if _col_idx is not None:
                report_lines.append(
                    f"<tr><td>{_label}</td>"
                    f"<td><code>{_hdr}</code></td>"
                    f"<td style='color: #28a745; text-align: center;'>&#10003;</td></tr>"
                )
            elif _is_optional:
                report_lines.append(
                    f"<tr><td>{_label}</td>"
                    f"<td style='color: #888;'>not found</td>"
                    f"<td style='color: #888; text-align: center;'>&mdash;</td></tr>"
                )
            else:
                report_lines.append(
                    f"<tr><td>{_label}</td>"
                    f"<td style='color: #c0392b;'>not found</td>"
                    f"<td style='color: #c0392b; text-align: center;'>&#10007;</td></tr>"
                )
        report_lines.append("</table>")
        report_lines.append("</details>")

    # CMS column detection (HCAHPS only, shown right below FRAME column detection)
    if audit_type == "HCAHPS":
        # Derive the same required list that check_req_headers used, including
        # the DRG-optional rule and AS being required.
        _all_possible_req = [
            "SID", "PATIENT NAME", "TELEPHONE", "D.DATE", "AGE", "AS", "DS",
            "GENDER", "UNIT", "PHYSICIAN NAME", "MRN", "DRG", "ATT",
            "LAG", "ID", "FD", "LG", "EMAIL ADDRESS", "CMS INDICATOR", "LANGUAGE",
        ]
        # DRG is optional when the header says "ALL MEDICAL DRGs"
        _drg_optional = "ALL MEDICAL DRGs".lower() in (header_text or "").lower()
        _cms_required_names = [n for n in _all_possible_req if not (n == "DRG" and _drg_optional)]
        _cms_required_set = set(_cms_required_names)
        # headers keys are the raw cell values from row 1; normalise for lookup
        _cms_found_req = sum(1 for n in _cms_required_names if headers.get(n) is not None)
        _cms_missing_req = len(_cms_required_names) - _cms_found_req
        _cms_summary_label = (
            f"CMS Column Detection &mdash; {_cms_found_req}/{len(_cms_required_names)} required found"
            + (f", {_cms_missing_req} missing" if _cms_missing_req else "")
        )
        _cms_open_attr = " open" if _cms_missing_req else ""
        report_lines.append(f"<details{_cms_open_attr} class='validation-summary-details'>")
        report_lines.append(
            f"<summary style='font-weight: normal;'>{_cms_summary_label}</summary>"
        )
        report_lines.append(
            f"<table class='excel-style' style='font-size: 0.9em; margin-top: 6px;'>"
        )
        report_lines.append(
            f"<tr>"
            f"<th style='background-color: {qtr_header_color};'>Column Name</th>"
            f"<th style='background-color: {qtr_header_color};'>Type</th>"
            f"<th style='background-color: {qtr_header_color};'>Found?</th>"
            f"</tr>"
        )
        # Required headers first
        for _rname in _cms_required_names:
            _col_idx = headers.get(_rname)
            if _col_idx is not None:
                report_lines.append(
                    f"<tr>"
                    f"<td><code>{_rname}</code></td>"
                    f"<td style='color: #555;'>Required</td>"
                    f"<td style='color: #28a745; text-align: center;'>&#10003;</td>"
                    f"</tr>"
                )
            else:
                report_lines.append(
                    f"<tr>"
                    f"<td><code style='color: #c0392b;'>{_rname}</code></td>"
                    f"<td style='color: #555;'>Required</td>"
                    f"<td style='color: #c0392b; text-align: center;'>&#10007;</td>"
                    f"</tr>"
                )
        # Extra (non-required) headers discovered in the sheet
        _extra_headers = [k for k in headers if k is not None and str(k).strip() and k not in _cms_required_set]
        if _extra_headers:
            report_lines.append(
                f"<tr><td colspan='3' style='background-color: #f0f0f0; font-size: 0.85em; "
                f"color: #555; padding: 3px 6px;'>Additional columns found in file</td></tr>"
            )
            for _ename in sorted(_extra_headers, key=lambda x: str(x).lower()):
                report_lines.append(
                    f"<tr>"
                    f"<td><code>{_ename}</code></td>"
                    f"<td style='color: #888;'>Extra</td>"
                    f"<td style='color: #888; text-align: center;'>&mdash;</td>"
                    f"</tr>"
                )
        report_lines.append("</table>")
        report_lines.append("</details>")

    if audit_type == "HCAHPS":
        # --- HCAHPS: DATA AT A GLANCE ---
        report_lines.append("<h2>DATA AT A GLANCE</h2>")
        c = 'text-align: center;'
        report_lines.append(f"<table class='excel-style' style='--header-color: {qtr_header_color};'>")
        report_lines.append("<tr>")
        report_lines.append(
            f"<th style='background-color: {qtr_header_color}; width: 35%;'>FACILITY</th>"
            f"<th style='background-color: {qtr_header_color}; {c}'>SID</th>"
            f"<th style='background-color: {qtr_header_color}; {c}'>DATE RANGE</th>"
            f"<th style='background-color: {qtr_header_color}; {c}'>POP</th>"
            f"<th style='background-color: {qtr_header_color}; {c}'>EL / SS</th>"
            f"<th style='background-color: {qtr_header_color}; {c}'>INEL</th>"
            f"<th style='background-color: {qtr_header_color}; {c}'>EXCLU</th>"
            f"<th style='background-color: {qtr_header_color}; {c}'>DUP</th>"
        )
        report_lines.append("</tr>")
        report_lines.append("<tr>")
        client_name_display = sid_registry_name if sid_registry_name else base_before_hash
        report_lines.append(f"<td>{client_name_display}</td>")
        report_lines.append(f"<td style='{c}'>{sid_prefix if sid_prefix else 'N/A'}</td>")
        # Format date range as "MAR 1-13" (or "MAR 28 - APR 5" if crossing months)
        _hcahps_date_display = 'N/A'
        if service_date_range:
            try:
                _dp = service_date_range.split(' - ')
                if len(_dp) == 2:
                    _d1 = datetime.datetime.strptime(_dp[0].strip(), '%m/%d/%Y')
                    _d2 = datetime.datetime.strptime(_dp[1].strip(), '%m/%d/%Y')
                    _MONTHS = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
                    if _d1.month == _d2.month and _d1.year == _d2.year:
                        _hcahps_date_display = f"{_MONTHS[_d1.month-1]} {_d1.day}-{_d2.day}"
                    else:
                        _hcahps_date_display = f"{_MONTHS[_d1.month-1]} {_d1.day} - {_MONTHS[_d2.month-1]} {_d2.day}"
            except (ValueError, AttributeError):
                _hcahps_date_display = service_date_range
        report_lines.append(f"<td style='{c}'>{_hcahps_date_display}</td>")
        _pop = pop_count if pop_count is not None else 'N/A'
        report_lines.append(f"<td style='{c}'>{_pop}</td>")
        _el = eligible_patients if eligible_patients is not None else 'N/A'
        _ss = sample_size if sample_size is not None else 'N/A'
        report_lines.append(f"<td style='{c}'>{_el} / {_ss}</td>")
        _inel = inel_count if inel_count is not None else 'N/A'
        report_lines.append(f"<td style='{c}'>{_inel}</td>")
        _exclu = exclu_count if exclu_count is not None else 'N/A'
        report_lines.append(f"<td style='{c}'>{_exclu}</td>")
        _dup = dup_count if dup_count is not None else 'N/A'
        report_lines.append(f"<td style='{c}'>{_dup}</td>")
        report_lines.append("</tr>")
        report_lines.append("</table>")
    else:
        # --- OAS: ESTIMATED LOG SHEET LINE ---
        report_lines.append("<h2>ESTIMATED LOG SHEET LINE"
            " <span class='info-icon'>i<span class='tooltip'>"
            "<b>Data sources:</b><br>"
            "SID - from header<br>"
            "Client - SID registry (SIDs.csv) or filename<br>"
            "Non-Reported - CMS INDICATOR = 2 count<br>"
            "Emails / Mailings - E/M column (CMS=1 rows)<br>"
            "Selection % - Sample Size / Eligible<br>"
            "Submitted - from header<br>"
            "Eligible - from footer (EL)<br>"
            "Sample Size - from footer (SS)"
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
    
    _sids_link = HCAHPS_SIDS_ONEDRIVE_LINK if audit_type == "HCAHPS" else OAS_SIDS_ONEDRIVE_LINK
    _install_dir = os.path.join(os.getenv("LOCALAPPDATA", r"%LOCALAPPDATA%"), "OAS-CAHPS-Auditor")
    _install_dir_url = "file:///" + _install_dir.replace("\\", "/")

    # Add SID client name comparison if available
    if sid_prefix and sid_registry_name:
        report_lines.append("<h2>SID REGISTRY CHECK"
            " <span class='info-icon'>i<span class='tooltip'>"
            "SIDs.csv contains the list of client names matched to SID codes. "
            "If facility/site name columns are found in the POP tab, they are shown below. "
            "Download the latest version from the "
            f"<a href='{_sids_link}' "
            "style='color: #5dade2;' target='_blank'>shared OneDrive folder</a> "
            f"and place it in your installation directory (<a href='{_install_dir_url}' style='color: #5dade2;' target='_blank'>{_install_dir}</a>)."
            "</span></span></h2>")
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

    else:
        # Show that SID registry check couldn't be performed
        report_lines.append("<h2>SID Registry Check"
            " <span class='info-icon'>i<span class='tooltip'>"
            "SIDs.csv contains the list of client names matched to SID codes. "
            "Download the latest version from the "
            f"<a href='{_sids_link}' "
            "style='color: #5dade2;' target='_blank'>shared OneDrive folder</a> "
            f"and place it in your installation directory (<a href='{_install_dir_url}' style='color: #5dade2;' target='_blank'>{_install_dir}</a>)."
            "</span></span></h2>")
        report_lines.append("<p style='color: #000; margin: 5px 0;'>")
        if not sid_prefix:
            report_lines.append("⚠ Unable to perform SID registry check: SID prefix not found in file")
        else:
            report_lines.append("⚠ Unable to perform SID registry check: Matching SID not found in registry")
        report_lines.append("</p>")

    # Show facility/location columns found in POP tab
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

    # DATA QUALITY VALIDATION SECTION
    from audit_lib_funcs import column_validations

    issues, row_issues = column_validations(
        sheet, headers, mrn_col, cms_col, em_col, issues, row_issues,
        filename_year=filename_year,
        filename_month=filename_month,
    )

    # Duplicate phone check: cross-column (TELEPHONE vs CELL PHONE) for both audit types.
    # OAS has no CELL PHONE column so cell_col will be None - that's handled gracefully.
    from audit_lib_funcs import check_duplicate_phones_cross_column
    _tel_col  = headers.get("TELEPHONE")
    _cell_col = headers.get("CELL PHONE")
    _dup_phone_issues, _dup_phone_row_issues = check_duplicate_phones_cross_column(
        sheet, _tel_col, _cell_col, mrn_col, cms_col
    )
    issues.extend(_dup_phone_issues)
    row_issues.extend(_dup_phone_row_issues)

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
                "description": f"'{eq['email']}' - {desc}",
            }
        )
        issues.append(f"{main_tab_name} Row {eq['row']}: Potentially invalid email '{eq['email']}' - {desc}")

    # 1. Surgical Category Validation (OAS only)
    cpt_col = headers.get("CPT") if audit_type == "OAS" else None
    cat_col = headers.get("SURGICAL CATEGORY") if audit_type == "OAS" else None
    if audit_type == "OAS":
        report_lines.append("")
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
    elif audit_type == "OAS":
        issue_msg = "Missing CPT or SURGICAL CATEGORY column in OASCAPHS"
        issues.append(issue_msg)

    # 2. UPLOAD vs OASCAPHS comparison (value-by-value)
    # Only run if UPLOAD tab exists AND row counts match - if counts differ,
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
                            f"{col}: {main_tab_name}='{oas_val}' | UPLOAD='{up_val}'"
                        )
                if row_mismatches:
                    mrn_val = oas_row[mrn_col - 1] if mrn_col else None
                    cms_val = oas_row[cms_col - 1] if cms_col else None
                    row_issues.append(
                        {
                            "row": r,
                            "mrn": mrn_val,
                            "cms": cms_val,
                            "issue_type": f"UPLOAD/{main_tab_name} Mismatch",
                            "description": "; ".join(row_mismatches),
                        }
                    )
                    issues.append(
                        f"Row {r}: " + "; ".join(row_mismatches)
                    )

    # 2b. Cross-tab consistency: POP vs UPLOAD email matching (OAS only)
    if audit_type == "OAS" and "UPLOAD" in wb.sheetnames:
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

    # Check combined ineligible math - handled in ADDITIONAL VALIDATIONS table above

    # 3. CPT Ineligibility Check (OAS only, when CMS == 1)
    if audit_type == "OAS" and cpt_col:
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
    elif audit_type == "OAS":
        issues.append("CPT column missing in OASCAPHS for ineligibility check")

    # DRG/APR Ineligibility Check (HCAHPS only, CMS = 1 rows)
    if audit_type == "HCAHPS":
        from audit_hcahps_funcs import is_ineligible_drg, is_ineligible_apr
        _drg_col = headers.get("DRG")
        _apr_col = headers.get("APR")
        for r, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if not any(row):
                continue
            cms_val = row[cms_col - 1] if cms_col else None
            mrn_val = row[mrn_col - 1] if mrn_col else None
            try:
                cms_int = int(float(str(cms_val).strip())) if cms_val is not None and str(cms_val).strip() != "" else None
            except Exception:
                cms_int = None
            if cms_int != 1:
                continue
            if _drg_col:
                drg_val = row[_drg_col - 1]
                ineligible, reason = is_ineligible_drg(drg_val)
                if ineligible:
                    row_issues.append({
                        "row": r, "mrn": mrn_val, "cms": cms_val,
                        "issue_type": "DRG Ineligible",
                        "description": f"DRG {drg_val} not eligible for HCAHPS ({reason})",
                    })
            if _apr_col:
                apr_val = row[_apr_col - 1]
                ineligible, reason = is_ineligible_apr(apr_val)
                if ineligible:
                    row_issues.append({
                        "row": r, "mrn": mrn_val, "cms": cms_val,
                        "issue_type": "APR Ineligible",
                        "description": f"APR {apr_val} not eligible for HCAHPS ({reason})",
                    })

    # DS/AS classification (HCAHPS only, CMS=1 rows)
    if audit_type == "HCAHPS":
        from audit_hcahps_funcs import classify_ds, classify_as
        _ds_col = headers.get("DS")
        _as_col = headers.get("AS")  # optional — not all vendors include this column
        for r, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if not any(row):
                continue
            cms_val = row[cms_col - 1] if cms_col else None
            mrn_val = row[mrn_col - 1] if mrn_col else None
            try:
                cms_int = int(float(str(cms_val).strip())) if cms_val is not None and str(cms_val).strip() != "" else None
            except Exception:
                cms_int = None
            if cms_int != 1:
                continue
            if _ds_col:
                ds_val = row[_ds_col - 1]
                disposition, reason = classify_ds(ds_val)
                if disposition == 'exclu':
                    row_issues.append({
                        "row": r, "mrn": mrn_val, "cms": cms_val,
                        "issue_type": "DS Should Be EXCLU",
                        "description": reason,
                    })
                elif disposition == 'inel':
                    row_issues.append({
                        "row": r, "mrn": mrn_val, "cms": cms_val,
                        "issue_type": "DS Should Be INEL",
                        "description": reason,
                    })
            if _as_col:
                as_val = row[_as_col - 1]
                disposition, reason = classify_as(as_val)
                if disposition == 'exclu':
                    row_issues.append({
                        "row": r, "mrn": mrn_val, "cms": cms_val,
                        "issue_type": "AS Should Be EXCLU",
                        "description": reason,
                    })

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
            issue_type = issue['issue_type']
            is_possible = issue_type.startswith("Possible") or issue_type.startswith("Potentially")
            row_style = "background-color: #fefce8;" if is_possible else ""
            report_lines.append(
                f"<tr style='{row_style}'><td style='padding: 3px 8px;'>{issue['row']}</td><td style='padding: 3px 8px;'>{mrn_display}</td><td style='padding: 3px 8px;'>{cms_display}</td><td style='padding: 3px 8px;'>{issue_type}</td><td style='padding: 3px 8px;'>{issue['description']}</td></tr>"
            )
        report_lines.append("</table>")
        report_lines.append("</details>")

    # Display general/non-row issues as list
    non_row_issues = [
        iss
        for iss in issues
        if not any(
            iss.startswith(f"{main_tab_name} Row") or iss.startswith(f"UPLOAD Row")
            or iss.startswith("Duplicate phone")
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



    # INVALID ADDRESSES section (OAS only)
    invalid_addresses = []
    noted_addresses = []
    if audit_type == "OAS":
        _name_col_addr = headers.get("PATIENT NAME")
        _age_col_addr  = headers.get("AGE")
        invalid_addresses, noted_addresses = check_address(
            sheet, addr1_col, city_col, state_col, zip_col, mrn_col, cms_col, em_col, addr2_col,
            name_col=_name_col_addr, age_col=_age_col_addr
        )
    if invalid_addresses:
        report_lines.append("<h2>(Possibly) INVALID ADDRESSES FOUND</h2>")
        report_lines.append("<details open>")
        report_lines.append(
            f"<summary>Invalid Address Details ({len(invalid_addresses)} found)</summary>"
        )

        # Parse all entries first so we can build the name-flip picker
        _addr_rows = []
        for address in invalid_addresses:
            # Format: "Row: 5 - MRN: '123' - CMS: '1' - E/M: 'E' - NAME: 'Smith, John' - AGE: '45' - ADDRESS: '{...}' - REASON: '...'"
            parts = address.split(" - ")
            row_num  = parts[0].replace("Row: ", "").strip()
            mrn_val  = parts[1].replace("MRN: ", "").strip("'")
            cms_val  = parts[2].replace("CMS: ", "").strip("'")
            em_val   = parts[3].replace("E/M: ", "").strip("'")
            name_val = parts[4].replace("NAME: ", "").strip("'") if len(parts) > 4 else ""
            age_val  = parts[5].replace("AGE: ", "").strip("'")  if len(parts) > 5 else ""

            mrn_val  = "" if mrn_val  == "None" else mrn_val
            cms_val  = "" if cms_val  == "None" else cms_val
            em_val   = "" if em_val   == "None" else em_val
            name_val = "" if name_val == "None" else name_val
            age_val  = "" if age_val  == "None" else age_val

            addr_dict_str = parts[6].replace("ADDRESS: ", "").strip("'") if len(parts) > 6 else ""
            try:
                addr_dict = ast.literal_eval(addr_dict_str)
                street   = addr_dict.get("street_address") or ""
                city     = addr_dict.get("city") or ""
                state    = addr_dict.get("country_area") or ""
                zip_code = addr_dict.get("postal_code") or ""
                street   = "" if street   == "None" else street
                city     = "" if city     == "None" else city
                state    = "" if state    == "None" else state
                zip_code = "" if zip_code == "None" else zip_code
            except Exception:
                street = city = state = zip_code = ""

            reason_text = parts[7].replace("REASON: ", "").strip("'") if len(parts) > 7 else ""
            _addr_rows.append({
                "row_num": row_num, "mrn": mrn_val, "cms": cms_val, "em": em_val,
                "name": name_val, "age": age_val,
                "street": street, "city": city, "state": state, "zip": zip_code,
                "reason": reason_text,
            })

        # Collect up to 3 sample names for the name-order picker
        _addr_sample_names = []
        for _r in _addr_rows:
            _n = _r["name"]
            if len(_n.split()) >= 2:
                _addr_sample_names.append(_n)
            if len(_addr_sample_names) >= 3:
                break
        _addr_show_picker = len(_addr_sample_names) > 0

        th = "<th style='background-color: #000; color: #fff; padding: 4px 8px;'>"

        def _build_addr_table(use_flipped):
            _rows = []
            _rows.append("<table class='excel-style' style='font-size: 0.85em;'>")
            _rows.append(
                f"<tr>{th}ROW</th>{th}MRN</th>{th}CMS</th>{th}E/M</th>"
                f"{th}PATIENT NAME</th>{th}AGE</th>"
                f"{th}STREET</th>{th}CITY</th>{th}STATE</th>{th}ZIP</th>"
                f"{th}REASON</th>"
                f"<th style='background-color:#000;color:#fff;padding:4px 8px;'>SEARCH LINKS</th></tr>"
            )
            for _r in _addr_rows:
                raw_name = _r["name"]
                tokens = raw_name.split()
                if use_flipped and len(tokens) >= 2:
                    lookup_name = " ".join(tokens[1:] + [tokens[0]])
                else:
                    lookup_name = raw_name
                name_disp = lookup_name or "&mdash;"
                if lookup_name:
                    _loc_city  = _r["city"]
                    _loc_state = _r["state"]
                    _urls = build_person_search_urls(lookup_name, _loc_city, _loc_state)
                    links_html = " &nbsp; ".join(
                        f"<a href='{url}' target='_blank' "
                        f"style='color:#2980b9;text-decoration:none;white-space:nowrap;'>{label}</a>"
                        for label, url in _urls.items()
                    )
                else:
                    links_html = "&mdash;"
                # Highlight specific blank cells for missing-field rows
                missing_fields = set()
                if _r["reason"].startswith("Missing:"):
                    for _mf in _r["reason"].replace("Missing:", "").split(","):
                        missing_fields.add(_mf.strip().lower())
                def _td(val, field="", _mf=missing_fields):
                    bg = "background-color: #fff3cd; " if field in _mf else ""
                    return f"<td style='{bg}padding: 3px 8px;'>{val}</td>"
                _rows.append(
                    f"<tr>"
                    f"{_td(_r['row_num'])}"
                    f"{_td(_r['mrn'])}"
                    f"{_td(_r['cms'])}"
                    f"{_td(_r['em'])}"
                    f"{_td(name_disp)}"
                    f"{_td(_r['age'])}"
                    f"{_td(_r['street'] or '&mdash;', 'street')}"
                    f"{_td(_r['city']   or '&mdash;', 'city')}"
                    f"{_td(_r['state']  or '&mdash;', 'state')}"
                    f"{_td(_r['zip']    or '&mdash;', 'zip')}"
                    f"{_td(_r['reason'])}"
                    f"{_td(links_html)}"
                    f"</tr>"
                )
            _rows.append("</table>")
            return _rows

        if _addr_show_picker:
            _flip_names = [
                " ".join(n.split()[1:] + [n.split()[0]]) for n in _addr_sample_names
            ]
            sep = " &nbsp;&middot;&nbsp; "
            raw_html  = sep.join(_addr_sample_names)
            flip_html = sep.join(_flip_names)
            report_lines.append("<input type='radio' name='addr-order' id='ao-raw' class='addr-radio' checked>")
            report_lines.append("<input type='radio' name='addr-order' id='ao-flip' class='addr-radio'>")
            report_lines.append("<p class='lookup-picker-label'>Choose whichever option shows names in the correct order (from first name to last name):</p>")
            report_lines.append(
                "<div class='lookup-order-picker addr-order-picker'>"
                f"<label for='ao-raw'>"
                f"<span class='pick-hint'>Option 1</span>"
                f"<span class='pick-sample'>{raw_html}</span>"
                f"</label>"
                f"<label for='ao-flip'>"
                f"<span class='pick-hint'>Option 2</span>"
                f"<span class='pick-sample'>{flip_html}</span>"
                f"</label>"
                f"</div>"
            )
            report_lines.append("<div class='addr-table-raw'>")
            report_lines.extend(_build_addr_table(use_flipped=False))
            report_lines.append("</div>")
            report_lines.append("<div class='addr-table-flip'>")
            report_lines.extend(_build_addr_table(use_flipped=True))
            report_lines.append("</div>")
        else:
            report_lines.extend(_build_addr_table(use_flipped=False))

        report_lines.append("</details>")

    # possibly problematic addresses
    if noted_addresses:
        report_lines.append("<h2>(Possibly) PROBLEMATIC ADDRESSES FOUND</h2>")
        report_lines.append("<details open>")
        report_lines.append(
            f"<summary>Problematic Address Details ({len(noted_addresses)} found)</summary>"
        )

        _noted_rows = []
        for address in noted_addresses:
            # Format: "Row: 5 - MRN: '123' - CMS: '1' - E/M: 'M' - NAME: 'x' - AGE: '45' - CITY: 'x' - STATE: 'x' - ADDRESS: '123 Main St' - REASON(s): '...'"
            parts = address.split(" - ")
            row_num  = parts[0].replace("Row: ",   "").strip()
            mrn_val  = parts[1].replace("MRN: ",   "").strip("'")
            cms_val  = parts[2].replace("CMS: ",   "").strip("'")
            em_val   = parts[3].replace("E/M: ",   "").strip("'")
            name_val = parts[4].replace("NAME: ",  "").strip("'") if len(parts) > 4 else ""
            age_val  = parts[5].replace("AGE: ",   "").strip("'") if len(parts) > 5 else ""
            city_val = parts[6].replace("CITY: ",  "").strip("'") if len(parts) > 6 else ""
            state_val= parts[7].replace("STATE: ", "").strip("'") if len(parts) > 7 else ""
            addr_val = parts[8].replace("ADDRESS: ","").strip("'") if len(parts) > 8 else ""
            reason_text = " - ".join(parts[9:]).replace("REASON(s): ", "", 1).strip("'") if len(parts) > 9 else ""
            mrn_val   = "" if mrn_val   == "None" else mrn_val
            cms_val   = "" if cms_val   == "None" else cms_val
            em_val    = "" if em_val    == "None" else em_val
            name_val  = "" if name_val  == "None" else name_val
            age_val   = "" if age_val   == "None" else age_val
            city_val  = "" if city_val  == "None" else city_val
            state_val = "" if state_val == "None" else state_val
            addr_val  = "" if addr_val  == "None" else addr_val
            _noted_rows.append({
                "row_num": row_num, "mrn": mrn_val, "cms": cms_val, "em": em_val,
                "name": name_val, "age": age_val,
                "city": city_val, "state": state_val,
                "address": addr_val, "reason": reason_text,
            })

        _noted_sample_names = []
        for _r in _noted_rows:
            _n = _r["name"]
            if len(_n.split()) >= 2:
                _noted_sample_names.append(_n)
            if len(_noted_sample_names) >= 3:
                break
        _noted_show_picker = len(_noted_sample_names) > 0

        _nth = "<th style='background-color: #000; color: #fff; padding: 4px 8px;'>"

        def _build_noted_table(use_flipped):
            _rows = []
            _rows.append("<table class='excel-style' style='font-size: 0.85em;'>")
            _rows.append(
                f"<tr>{_nth}ROW</th>{_nth}MRN</th>{_nth}CMS</th>{_nth}E/M</th>"
                f"{_nth}PATIENT NAME</th>{_nth}AGE</th>"
                f"{_nth}ADDRESS</th>{_nth}ISSUE(S)</th>"
                f"<th style='background-color:#000;color:#fff;padding:4px 8px;'>SEARCH LINKS</th></tr>"
            )
            for _r in _noted_rows:
                raw_name = _r["name"]
                tokens = raw_name.split()
                if use_flipped and len(tokens) >= 2:
                    lookup_name = " ".join(tokens[1:] + [tokens[0]])
                else:
                    lookup_name = raw_name
                name_disp = lookup_name or "&mdash;"
                if lookup_name:
                    _urls = build_person_search_urls(lookup_name, _r["city"], _r["state"])
                    links_html = " &nbsp; ".join(
                        f"<a href='{url}' target='_blank' "
                        f"style='color:#2980b9;text-decoration:none;white-space:nowrap;'>{label}</a>"
                        for label, url in _urls.items()
                    )
                else:
                    links_html = "&mdash;"
                _rows.append(
                    f"<tr>"
                    f"<td style='padding: 3px 8px;'>{_r['row_num']}</td>"
                    f"<td style='padding: 3px 8px;'>{_r['mrn']}</td>"
                    f"<td style='padding: 3px 8px;'>{_r['cms']}</td>"
                    f"<td style='padding: 3px 8px;'>{_r['em']}</td>"
                    f"<td style='padding: 3px 8px;'>{name_disp}</td>"
                    f"<td style='padding: 3px 8px;'>{_r['age']}</td>"
                    f"<td style='padding: 3px 8px;'>{_r['address'] or '&mdash;'}</td>"
                    f"<td style='padding: 3px 8px;'>{_r['reason']}</td>"
                    f"<td style='padding: 3px 8px;'>{links_html}</td>"
                    f"</tr>"
                )
            _rows.append("</table>")
            return _rows

        if _noted_show_picker:
            _flip_names = [
                " ".join(n.split()[1:] + [n.split()[0]]) for n in _noted_sample_names
            ]
            sep = " &nbsp;&middot;&nbsp; "
            raw_html  = sep.join(_noted_sample_names)
            flip_html = sep.join(_flip_names)
            report_lines.append("<input type='radio' name='noted-order' id='no-raw' class='noted-radio' checked>")
            report_lines.append("<input type='radio' name='noted-order' id='no-flip' class='noted-radio'>")
            report_lines.append("<p class='lookup-picker-label'>Choose whichever option shows names in the correct order (from first name to last name):</p>")
            report_lines.append(
                "<div class='lookup-order-picker noted-order-picker'>"
                f"<label for='no-raw'>"
                f"<span class='pick-hint'>Option 1</span>"
                f"<span class='pick-sample'>{raw_html}</span>"
                f"</label>"
                f"<label for='no-flip'>"
                f"<span class='pick-hint'>Option 2</span>"
                f"<span class='pick-sample'>{flip_html}</span>"
                f"</label>"
                f"</div>"
            )
            report_lines.append("<div class='noted-table-raw'>")
            report_lines.extend(_build_noted_table(use_flipped=False))
            report_lines.append("</div>")
            report_lines.append("<div class='noted-table-flip'>")
            report_lines.extend(_build_noted_table(use_flipped=True))
            report_lines.append("</div>")
        else:
            report_lines.extend(_build_noted_table(use_flipped=False))

        report_lines.append("</details>")

    # PEOPLE-SEARCH LOOKUP SECTION
    candidates = collect_lookup_candidates(sheet, headers, mrn_col, cms_col)
    if candidates:
        report_lines.append("<h2>CONTACT LOOKUP</h2>")
        th = "<th style='background-color: #000; color: #fff; padding: 4px 8px;'>"
        report_lines.append("<details open>")
        report_lines.append(f"<summary>CMS=1 patients with contact issues ({len(candidates)} found)</summary>")

        # Collect up to 3 sample names from lookup-mode candidates for the name-order picker
        sample_names = []
        for c in candidates:
            if c["mode"] != "lookup":
                continue
            name = (c["name"] or "").strip()
            if len(name.split()) >= 2:
                sample_names.append(name)
            if len(sample_names) >= 3:
                break
        show_picker = len(sample_names) > 0

        def _build_lookup_table(use_flipped):
            rows = []
            rows.append("<table class='excel-style' style='font-size: 0.85em;'>")
            rows.append(
                f"<tr>{th}ROW</th>{th}MRN</th>{th}PATIENT NAME</th>{th}AGE</th>"
                f"{th}CITY, STATE</th>{th}REASON(S)</th>"
                f"<th style='background-color:#000;color:#fff;padding:4px 8px;'>SEARCH LINKS</th></tr>"
            )
            for c in candidates:
                mrn_disp = c["mrn"] if c["mrn"] is not None else ""
                age_disp = c["age"] if c["age"] is not None else ""
                location = ", ".join(x for x in [c["city"], c["state"]] if x) or "&mdash;"
                reasons  = "; ".join(c["issues"])
                if c["mode"] == "lookup":
                    raw_name = (c["name"] or "").strip()
                    tokens = raw_name.split()
                    if use_flipped and len(tokens) >= 2:
                        lookup_name = " ".join(tokens[1:] + [tokens[0]])
                    else:
                        lookup_name = raw_name
                    name_disp  = lookup_name or "&mdash;"
                    urls       = build_person_search_urls(lookup_name, c["city"], c["state"])
                    links_html = " &nbsp; ".join(
                        f"<a href='{url}' target='_blank' "
                        f"style='color:#2980b9;text-decoration:none;white-space:nowrap;'>{label}</a>"
                        for label, url in urls.items()
                    )
                else:
                    name_disp  = c["name"] or "&mdash;"
                    links_html = "&mdash;"
                rows.append(
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
            rows.append("</table>")
            return rows

        if show_picker:
            flip_names = [
                " ".join(n.split()[1:] + [n.split()[0]]) for n in sample_names
            ]
            sep = " &nbsp;&middot;&nbsp; "
            raw_html  = sep.join(sample_names)
            flip_html = sep.join(flip_names)
            report_lines.append("<input type='radio' name='lookup-order' id='lo-raw' class='lookup-radio' checked>")
            report_lines.append("<input type='radio' name='lookup-order' id='lo-flip' class='lookup-radio'>")
            report_lines.append("<p class='lookup-picker-label'>Choose whichever option shows names in the correct order (from first name to last name):</p>")
            report_lines.append(
                "<div class='lookup-order-picker'>"
                f"<label for='lo-raw'>"
                f"<span class='pick-hint'>Option 1</span>"
                f"<span class='pick-sample'>{raw_html}</span>"
                f"</label>"
                f"<label for='lo-flip'>"
                f"<span class='pick-hint'>Option 2</span>"
                f"<span class='pick-sample'>{flip_html}</span>"
                f"</label>"
                f"</div>"
            )
            report_lines.append("<div class='lookup-table-raw'>")
            report_lines.extend(_build_lookup_table(use_flipped=False))
            report_lines.append("</div>")
            report_lines.append("<div class='lookup-table-flip'>")
            report_lines.extend(_build_lookup_table(use_flipped=True))
            report_lines.append("</div>")
        else:
            report_lines.extend(_build_lookup_table(use_flipped=False))

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

    # FRAME invalid addresses (HCAHPS only)
    if audit_type == "HCAHPS" and frame_col_map is not None:
        # FRAME invalid addresses
        _frame_invalid = frame_invalid_addresses or []
        _frame_noted = frame_noted_addresses or []
        if _frame_invalid or _frame_noted:
            _total_addr = len(_frame_invalid) + len(_frame_noted)
            report_lines.append("<h2>FRAME INVALID ADDRESSES</h2>")
            report_lines.append("<details open>")
            report_lines.append(f"<summary>FRAME address issues ({_total_addr} found)</summary>")
            report_lines.append("<table class='excel-style' style='font-size: 0.85em;'>")
            _ath = "<th style='background-color: #000; color: #fff; padding: 4px 8px;'>"
            report_lines.append(
                f"<tr>{_ath}ROW</th>{_ath}MRN</th>"
                f"{_ath}STREET</th>{_ath}CITY</th>{_ath}STATE</th>{_ath}ZIP</th>"
                f"{_ath}REASON</th></tr>"
            )
            def _parse_frame_addr(addr_str):
                # Format (invalid): Row - MRN - CMS - E/M - NAME - AGE - ADDRESS(dict) - REASON
                # Format (noted):   Row - MRN - CMS - E/M - NAME - AGE - CITY - STATE - ADDRESS(street) - REASON(s)
                parts = addr_str.split(" - ")
                row_num = parts[0].replace("Row: ", "").strip()
                mrn_val = parts[1].replace("MRN: ", "").strip("'") if len(parts) > 1 else ""
                mrn_val = "" if mrn_val == "None" else mrn_val
                # parts[2]=CMS, [3]=E/M, [4]=NAME, [5]=AGE — skip for FRAME display
                is_noted = len(parts) > 6 and parts[6].startswith("CITY: ")
                if is_noted:
                    city_val  = parts[6].replace("CITY: ",    "").strip("'") if len(parts) > 6 else ""
                    state_val = parts[7].replace("STATE: ",   "").strip("'") if len(parts) > 7 else ""
                    street    = parts[8].replace("ADDRESS: ", "").strip("'") if len(parts) > 8 else ""
                    reason    = " - ".join(parts[9:]).replace("REASON(s): ", "", 1).strip("'") if len(parts) > 9 else ""
                    zip_val   = ""
                else:
                    addr_dict_str = parts[6].replace("ADDRESS: ", "").strip("'") if len(parts) > 6 else ""
                    reason = parts[7].replace("REASON: ", "").strip("'") if len(parts) > 7 else ""
                    try:
                        d = ast.literal_eval(addr_dict_str)
                        street    = d.get("street_address") or ""
                        city_val  = d.get("city") or ""
                        state_val = d.get("country_area") or ""
                        zip_val   = d.get("postal_code") or ""
                    except Exception:
                        street = city_val = state_val = zip_val = ""
                street    = "" if street    == "None" else street
                city_val  = "" if city_val  == "None" else city_val
                state_val = "" if state_val == "None" else state_val
                zip_val   = "" if zip_val   == "None" else zip_val
                return row_num, mrn_val, street, city_val, state_val, zip_val, reason
            for _addr in _frame_invalid + _frame_noted:
                _rn, _mrn, _st, _ci, _sta, _zp, _re = _parse_frame_addr(_addr)
                report_lines.append(
                    f"<tr>"
                    f"<td style='padding:3px 8px;'>{_rn}</td>"
                    f"<td style='padding:3px 8px;'>{_mrn}</td>"
                    f"<td style='padding:3px 8px;'>{_st}</td>"
                    f"<td style='padding:3px 8px;'>{_ci}</td>"
                    f"<td style='padding:3px 8px;'>{_sta}</td>"
                    f"<td style='padding:3px 8px;'>{_zp}</td>"
                    f"<td style='padding:3px 8px;'>{_re}</td>"
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


def _build_html_header(file_path, version, audit_id=None, sid_prefix=None, service_date_range=None, audit_type="OAS"):
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
    # Compute header accent color — mirrors logic in build_report
    qtr_header_color = HCAPHS_HEADER_COLOR if audit_type == "HCAHPS" else "#2dbd69"  # lavender or green
    if audit_type != "HCAHPS" and service_date_range:
        try:
            _dp = service_date_range.split(" - ")
            if len(_dp) == 2:
                _sd = datetime.datetime.strptime(_dp[0].strip(), "%m/%d/%Y")
                if _sd.month % 2 == 1:
                    qtr_header_color = "#ec8038"  # Orange for odd months
        except (ValueError, AttributeError):
            pass
    header_lines.append("<!DOCTYPE html>")
    header_lines.append('<html lang="en">')
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
        with open(css_path, "r", encoding="utf-8-sig") as css_file:
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
    header_lines.append(f"<div style='display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 5px; border-bottom: 2px solid {qtr_header_color}; padding-bottom: 5px;'>")
    title_text = "HCAHPS Audit Report" if audit_type == "HCAHPS" else "OAS-CAHPS Audit Report"
    header_lines.append(f"<h1 style='margin: 0; border: none; padding: 0;'>{title_text}</h1>")
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
                
                header_lines.append(f"<div class='header-service-dates'>{long_date_range}</div>")
            else:
                # Fallback to original if parsing fails
                header_lines.append(f"<div class='header-service-dates'>{service_date_range}</div>")
        except (ValueError, AttributeError):
            # Fallback to original if parsing fails
            header_lines.append(f"<div class='header-service-dates'>{service_date_range}</div>")
    header_lines.append("</div>")
    header_lines.append(
        f"<div class='header-meta-row'>"
        f"<span><a href='https://tylercbrock.com' class='meta-link' title='tylercbrock.com' target='_blank'>Auditor</a> v{version}</span>"
        f"<span><a href='https://github.com/ToonLunk/OAS-CAHPS-Auditor' class='meta-link' title='OAS-CAHPS-Auditor on GitHub' target='_blank'>Need Help?</a></span>"
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
      - True: %LOCALAPPDATA%\\OAS-CAHPS-Auditor\\AUDITS\\YEAR\\MONTH\\
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
            
            # Look for " OAS " or " HCAHPS " to split month and year
            _type_tok = next(
                (t for t in (" OAS ", " HCAHPS ") if t in name_part), None
            )
            if _type_tok:
                month_part, year_part = name_part.split(_type_tok, 1)
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
            "class='update-badge' "
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
