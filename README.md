# MBOX Converter v2.0

A powerful, cross-platform command-line tool for converting Mozilla MBOX files to CSV, PST, EML, or TXT formats. Features batch processing, email filtering, progress reporting, and full shell integration.

**Works on:** Windows, Linux, macOS, and cloud environments

## Features

- **Multiple Output Formats:** CSV, PST (Windows+Outlook), EML (individual files), TXT
- **Batch Processing:** Convert multiple files with wildcards (`*.mbox`) or directory scanning
- **Email Filtering:** Filter by date range, sender, recipient, subject (regex), body content
- **Progress Reporting:** Real-time progress bars with `tqdm`
- **Logging:** Configurable verbosity and file logging
- **Cross-Platform:** Python-based with shell wrappers for Windows (BAT/PowerShell) and Unix (Bash)
- **Backward Compatible:** Supports legacy command syntax
- **Structured Output:** JSON reports and exit codes for scripting

## Requirements

- Python 3.8+
- For PST conversion: Windows with Microsoft Outlook installed + `pywin32` package

## Installation

```bash
# Clone or download the repository
cd mbox-converter

# Install dependencies
pip install -r requirements.txt

# Make scripts executable (Linux/macOS)
chmod +x mbox_converter.py mbox_converter.sh
```

## Quick Start

### Convert to CSV
```bash
python mbox_converter.py convert inbox.mbox --format csv
```

### Convert to EML (individual email files)
```bash
python mbox_converter.py convert inbox.mbox --format eml --output-dir ./emails
```

### Batch Convert All MBOX Files
```bash
python mbox_converter.py convert ./archive/*.mbox --format csv --output-dir ./converted
```

### Filter by Date and Sender
```bash
python mbox_converter.py convert inbox.mbox --format csv \
    --date-after 2023-01-01 \
    --date-before 2024-01-01 \
    --from-pattern "@company.com"
```

## Command Reference

### Subcommands

| Command | Description |
|---------|-------------|
| `convert` | Convert MBOX files to CSV/PST/EML/TXT |
| `info` | Display MBOX file statistics |
| `list` | List emails with optional filtering |
| `config` | Generate sample configuration file |

### Convert Options

```
Usage: mbox_converter.py convert [OPTIONS] INPUT [INPUT ...]

Arguments:
  INPUT                    MBOX file(s), glob patterns (*.mbox), or directories

Options:
  -f, --format {csv,pst,eml,txt}  Output format (default: csv)
  -o, --output-dir PATH           Output directory
  -e, --encoding ENCODING         Email encoding (default: utf-8)
  --dry-run                       Preview without writing files
  -p, --progress                  Show progress bar
  -q, --quiet                     Suppress output (errors only)
  -v, --verbose                   Increase verbosity (-v, -vv)
  --log-file PATH                 Write logs to file
  --report PATH                   Write JSON conversion report

Filtering Options:
  --date-after DATE               Only include emails after date (YYYY-MM-DD)
  --date-before DATE              Only include emails before date
  --from-pattern REGEX            Filter by sender (regex)
  --to-pattern REGEX              Filter by recipient (regex)
  --subject-pattern REGEX         Filter by subject (regex)
  --body-contains TEXT            Filter by body content
  --has-attachment [true|false]   Filter by attachment presence
```

### Info Command

```bash
# Get statistics about MBOX file
python mbox_converter.py info inbox.mbox

# Output as JSON
python mbox_converter.py info inbox.mbox --json
```

Output:
```
=== inbox.mbox ===
  Total emails: 1,234
  Unique senders: 89
  With attachments: 156
  File size: 45.6 MB
  Date range: 2020-01-15 to 2024-03-22
```

### List Command

```bash
# List recent emails
python mbox_converter.py list inbox.mbox --limit 20

# Search for specific emails
python mbox_converter.py list inbox.mbox --subject-pattern "invoice" --from-pattern "@vendor.com"

# Output as JSON for scripting
python mbox_converter.py list inbox.mbox --json --limit 50
```

## Platform-Specific Usage

### Windows (CMD)

```cmd
REM Using batch script (auto-installs Python if needed)
run_converter.bat convert inbox.mbox --format csv

REM Direct Python
python mbox_converter.py convert inbox.mbox --format pst
```

### Windows (PowerShell)

```powershell
# Import module
Import-Module ./MboxConverter.ps1

# Convert with pipeline
Get-ChildItem *.mbox | Convert-MBox -Format csv -OutputDirectory ./output

# Get info
Get-MBoxInfo -Path inbox.mbox

# Search emails
Get-MBoxEmails -Path inbox.mbox -SubjectPattern "invoice" -Limit 50
```

### Linux / macOS (Bash)

```bash
# Using shell wrapper
./mbox_converter.sh convert inbox.mbox --format csv

# Direct Python
python3 mbox_converter.py convert ./emails/*.mbox --format eml -p

# Pipeline with find
find /archives -name "*.mbox" -exec python3 mbox_converter.py convert {} --format csv \;
```

### Cloud / CI/CD

```yaml
# GitHub Actions example
- name: Convert MBOX archives
  run: |
    pip install tqdm python-dateutil
    python mbox_converter.py convert ./data/*.mbox --format csv --output-dir ./csv-exports --report report.json
```

## Output Formats

### CSV
Standard comma-separated values with headers:
```
Index,From,To,Cc,Bcc,Subject,Date,Body,HasAttachment
```

### EML
Individual `.eml` files per email, named: `000001_Subject_Here.eml`

### TXT
Plain text file with all emails separated by delimiters, includes headers and body.

### PST (Windows Only)
Outlook Personal Storage Table format. Requires:
- Microsoft Outlook installed
- `pywin32` package

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | Error |
| 2 | Invalid arguments |
| 3 | Partial success (some files failed) |

## Examples

### 1. Basic CSV Conversion
```bash
python mbox_converter.py convert inbox.mbox --format csv
```

### 2. Batch Processing with Progress
```bash
python mbox_converter.py convert ./archives/*.mbox --format eml --output-dir ./exported -p
```

### 3. Filter Recent Emails from Specific Sender
```bash
python mbox_converter.py convert inbox.mbox --format csv \
    --date-after 2024-01-01 \
    --from-pattern "boss@company.com|ceo@company.com"
```

### 4. Extract Emails with Attachments
```bash
python mbox_converter.py convert inbox.mbox --format eml \
    --has-attachment true \
    --output-dir ./with-attachments
```

### 5. Search and Export Invoices
```bash
python mbox_converter.py convert inbox.mbox --format csv \
    --subject-pattern "invoice|receipt|payment" \
    --date-after 2023-01-01
```

### 6. Dry Run Preview
```bash
python mbox_converter.py convert ./emails/*.mbox --format csv --dry-run
```
Output:
```
=== DRY RUN MODE ===
Would process 5 file(s):
  - emails/inbox.mbox: 1234 emails (45.6 MB)
  - emails/sent.mbox: 567 emails (12.3 MB)
  ...
```

### 7. Generate Conversion Report
```bash
python mbox_converter.py convert inbox.mbox --format csv --report conversion_report.json
```

## Legacy Mode

For backward compatibility, the original syntax still works:
```bash
python mbox_converter.py inbox.mbox --csv output.csv
python mbox_converter.py inbox.mbox --pst output.pst
python mbox_converter.py inbox.mbox --eml ./eml-output
python mbox_converter.py inbox.mbox --txt output.txt
```

## Building Standalone Executable (Windows)

Create a portable `.exe` that runs without Python:

```bash
python build_exe.py
```

This creates `dist/mbox_converter.exe` using PyInstaller.

## Configuration File

Generate a sample config for reusable settings:

```bash
python mbox_converter.py config --generate my_config.json
```

## Troubleshooting

### "tqdm not found"
```bash
pip install tqdm python-dateutil
```

### PST conversion fails
- Ensure Microsoft Outlook is installed
- Install pywin32: `pip install pywin32`
- Run as administrator if needed

### Unicode errors
Use the `--encoding` flag:
```bash
python mbox_converter.py convert inbox.mbox --format csv --encoding latin-1
```

### Large files are slow
Use `--progress` to monitor and consider filtering to reduce volume:
```bash
python mbox_converter.py convert large.mbox --format csv -p --date-after 2023-01-01
```

## License

MIT License - See [LICENSE](LICENSE) for details.

## Contributing

Contributions welcome! Please submit issues and pull requests on GitHub.
