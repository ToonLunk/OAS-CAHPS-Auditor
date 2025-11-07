# OAS CAHPS Auditor

Audits OAS CAHPS Excel files for data integrity and compliance. Validates headers, sample sizes, addresses, CPT codes, cross-tab consistency, and more. Generates detailed HTML reports.

## Usage

```cmd
audit <excel_file>    # Audit a single file
audit --all           # Audit all Excel files in current directory
audit --help          # Show help
audit --version       # Show version
```

## Installation

**Option 1: Use Pre-Built Executable (Easiest)**

If you received `audit.exe` have, simply:

1. Place `audit.exe` in any folder
2. (Optional) Add that folder to your Windows PATH to use `audit` from anywhere

**Option 2: Build from Source**

If you have the source code files:

### Prerequisites

- **Python 3.8+** - [Download here](https://www.python.org/downloads/)
  - Check "Add Python to PATH" during installation

### Build Steps

1. **Navigate to the source folder**

   ```cmd
   cd path\to\source\folder
   ```

2. **Install dependencies**

   ```cmd
   pip install -r requirements.txt
   ```

3. **Build the executable**

   ```cmd
   build_exe.bat
   ```

   This creates `audit.exe` in the `dist` folder (first build takes 2-5 minutes).

4. **Install (optional)**

   **Quick install:** Copy `deploy.bat` into the same folder as `audit.exe`, then run `deploy.bat` as administrator.

   **Manual:** Place `audit.exe` anywhere and add that folder to your Windows PATH.

### Troubleshooting

- **"Python is not recognized"** → Reinstall Python with "Add to PATH" checked
- **"pip is not recognized"** → Use `py -m pip install -r requirements.txt`
- **Antivirus blocks .exe** → Add exception (PyInstaller can trigger false positives)
- **Build fails** → Delete `build` and `dist` folders, then rebuild

### Development

Edit source files (`audit.py`, `audit_lib_funcs.py`, `audit_printer.py`, `audit_report.css`), test with `python audit.py <file>`, then run `build_exe.bat` to rebuild.
