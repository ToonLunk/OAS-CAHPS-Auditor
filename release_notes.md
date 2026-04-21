# Version 1.3.1 — Facility Recognition Fix

**Full Changelog**: https://github.com/ToonLunk/OAS-CAHPS-Auditor/compare/v1.3.0...v1.3.1

## What's New

Improved facility recognition logic to ensure facility names are properly identified and matched in reports.

---

## Install / Update

1. Download **`OAS-CAHPS-Auditor-v1.3.1-Setup.exe`** below.
2. Run the installer — it will upgrade in place if you already have a previous version.

Default install location: `C:\OAS-CAHPS-Auditor`

---

## How to Use

**Audit a single file** — hold **Shift**, right-click the Excel file, and select **"Audit this OAS file"**.

**Audit an entire folder** — hold **Shift**, right-click empty space inside the folder, and select **"Audit All OAS Files"**.

Reports are saved in an **AUDITS** folder next to your files and open automatically in your browser.

---

# Version 1.3.0 — Email Quality Checks

**Full Changelog**: https://github.com/ToonLunk/OAS-CAHPS-Auditor/compare/v1.2.0...v1.3.0

## What's New

Email validation has been overhauled. The old regex is gone — the auditor now uses the `email-validator` library and flags a wider range of suspicious addresses:

- **RFC violations** — leading/trailing dots, consecutive dots, and other malformed patterns that used to slip through
- **Placeholder & opt-out addresses** — `optout@`, `noreply@`, `declined@`, `test@`, single-character or all-numeric local parts (`0@gmail.com`), disposable domains (`mailinator.com`), and very short addresses
- **CMS=1** flagged emails appear in the main Issues table
- **CMS=2** flagged emails are in a collapsible section at the bottom of the report (closed by default)

Also new: the report is now **print-friendly**. `Ctrl+P` produces a clean layout with all sections expanded and interactive elements hidden.

---

## Install / Update

1. Download **`OAS-CAHPS-Auditor-v1.3.0-Setup.exe`** below.
2. Run the installer — it will upgrade in place if you already have a previous version.

Default install location: `C:\OAS-CAHPS-Auditor`

---

## How to Use

**Audit a single file** — hold **Shift**, right-click the Excel file, and select **"Audit this OAS file"**.

**Audit an entire folder** — hold **Shift**, right-click empty space inside the folder, and select **"Audit All OAS Files"**.

Reports are saved in an **AUDITS** folder next to your files and open automatically in your browser.