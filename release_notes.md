# Version 1.2.0 — Contact Lookup

**Full Changelog**: https://github.com/ToonLunk/OAS-CAHPS-Auditor/compare/v1.1.1...v1.2.0

## What's New

The HTML report now includes a **CONTACT LOOKUP** section for CMS=1 patients who have contact issues. It appears automatically when problems are found — no extra steps needed.

- **Search links** are included for patients with no valid phone number or an invalid email. Links open pre-populated on WhitePages, TruePeopleSearch, and FastPeopleSearch — nothing is fetched until you click.
- **Reference rows** are shown for patients who have at least one valid phone number but an invalid entry in the other field — the reason is noted, but no lookup is needed.
- Phone validation catches numbers that look formatted correctly but aren't real assignable US numbers (e.g. `1234567890`).

---

## Installation
Installation is now way easier. Just download the file and install the program! Please read the instructions carefully.

1. Download **`OAS-CAHPS-Auditor-v1.0.0-Setup.exe`** below.
2. Run the setup wizard and follow the prompts.

The Auditor will be installed to `C:\OAS-CAHPS-Auditor`.

---

## How to Use

### Audit a Single File

Right-click your Excel file while **holding SHIFT**, then select **"Audit this OAS file"**.

Your audit report will open automatically in your browser.

### Audit All Files in a Folder

**Hold SHIFT** and right-click on empty space inside the folder, then select **"Audit All OAS Files"**.

> **Note:** You must hold SHIFT when right-clicking to see these options.

---

All reports are saved in an **AUDITS** folder next to your files.

Full instructions are also included in `Installation Instructions.txt` alongside `audit.exe`.