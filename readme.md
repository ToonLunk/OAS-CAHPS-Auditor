# CAHPS Auditor

Audits OAS CAHPS and HCAHPS Excel files and outputs an HTML report. Detects header issues, sample size problems, data quality errors, cross-tab mismatches, and much more.

Download the latest release from the [Releases page](https://github.com/ToonLunk/OAS-CAHPS-Auditor/releases).

## Installation

1. Download and run `OAS-CAHPS-Auditor-v{VERSION}-Setup.exe`
2. Follow the setup wizard - it installs to `%LOCALAPPDATA%\OAS-CAHPS-Auditor` and adds it to your PATH
3. Optionally register right-click context menu entries for Explorer (recommended)

**To Uninstall:** Settings -> Apps -> Installed Apps -> CAHPS Auditor -> Uninstall

## Usage

**Command line:**
```cmd
audit filename.xlsx    # Audit a specific file
audit --all            # Audit all Excel files in the current directory
audit --version        # Show version number
```

**Context menu (recommended):**
- **Right-click** any `.xlsx` file -> **"Audit This CAHPS File"**
- **Right-click** inside a folder -> **"Audit All CAHPS Files"**

On Windows 11, hold **Shift** while right-clicking, or choose **"Show more options"**.

## Links

- [Changelog](CHANGELOG.md)
- [To-Do](todo.md)
- [License](LICENSE)
- [Development / Packaging](docs/PACKAGING_README.md)
- [Releases](https://github.com/ToonLunk/OAS-CAHPS-Auditor/releases)
