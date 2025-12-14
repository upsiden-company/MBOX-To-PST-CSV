# Email Converter v4.0

**Universal Email Migration Tool** - Convert between ALL major email formats with easy-to-use migration presets.

**Works on:** Windows, Linux, macOS, and cloud environments

## Supported Formats (8 Total)

| Format | Extension | Read | Write | Description |
|--------|-----------|------|-------|-------------|
| **MBOX** | .mbox | ✓ | ✓ | Mozilla Thunderbird, Google Takeout |
| **CSV** | .csv | ✓ | ✓ | Spreadsheet (Excel, Google Sheets) |
| **EML** | .eml/* | ✓ | ✓ | RFC 822 Standard Email Files |
| **MSG** | .msg/* | ✓ | ✓ | Microsoft Outlook Individual Messages |
| **TXT** | .txt | ✓ | ✓ | Plain Text (Human Readable) |
| **PST** | .pst | ✓ | ✓ | Outlook Personal Storage (Windows) |
| **JSON** | .json | ✓ | ✓ | Structured Data (API/Programming) |
| **Maildir** | folder/ | ✓ | ✓ | Unix Mail Directory (Dovecot/Postfix) |

**All formats can convert to any other format!**

## Easy Migration Presets

Built-in presets for common email migrations:

| Preset | Description |
|--------|-------------|
| `google-to-365` | Google Workspace → Microsoft 365 |
| `thunderbird-to-outlook` | Thunderbird → Outlook |
| `outlook-to-thunderbird` | Outlook → Thunderbird |
| `apple-to-outlook` | Apple Mail → Outlook |
| `yahoo-to-gmail` | Yahoo/AOL → Gmail |
| `outlook-to-gmail` | Outlook → Gmail |
| `eml-to-outlook` | EML Files → Outlook |
| `backup-to-csv` | Any Format → CSV for archiving |
| `maildir-to-mbox` | Unix Maildir → MBOX |
| `mbox-to-maildir` | MBOX → Unix Maildir |

## Quick Start

### Easy Migrations (Recommended)

```bash
# Google Workspace to Microsoft 365
python mbox_converter.py migrate google-to-365 ./Takeout/*.mbox -o ./for_outlook

# Thunderbird to Outlook
python mbox_converter.py migrate thunderbird-to-outlook ./Inbox -o ./outlook_import

# Backup all emails to CSV
python mbox_converter.py migrate backup-to-csv ./emails/* -o ./backup
```

### Direct Conversions

```bash
# MBOX to CSV
python mbox_converter.py convert inbox.mbox --format csv

# CSV to MBOX
python mbox_converter.py convert emails.csv --format mbox

# EML files to JSON
python mbox_converter.py convert ./eml_folder/ --format json

# JSON to EML files
python mbox_converter.py convert emails.json --format eml

# MBOX to Maildir (for Unix servers)
python mbox_converter.py convert inbox.mbox --format maildir
```

### Show Available Options

```bash
# Show all supported formats
python mbox_converter.py formats

# Show all migration presets
python mbox_converter.py presets

# Show help
python mbox_converter.py --help
```

## Features

- **8 Email Formats** - MBOX, CSV, EML, MSG, TXT, PST, JSON, Maildir
- **Bidirectional Conversion** - Any format to any other format
- **Easy Migration Presets** - One command for common migrations
- **Batch Processing** - Convert multiple files with wildcards
- **Email Filtering** - Filter by date, sender, subject, body content
- **Progress Bars** - Real-time feedback with tqdm
- **Format Auto-Detection** - Automatically detects input format
- **Cross-Platform** - Works on Windows, Linux, macOS, cloud

## Installation

```bash
# Clone or download
cd email-converter

# Install dependencies
pip install -r requirements.txt

# (Linux/macOS) Make executable
chmod +x mbox_converter.py
```

### Optional Dependencies

```bash
# For MSG file support
pip install extract-msg

# For PST support (Windows only)
pip install pywin32
```

## Command Reference

### Convert Command

```bash
python mbox_converter.py convert <input> --format <format> [options]

Options:
  --format, -f    Output format (mbox/csv/eml/msg/txt/pst/json/maildir)
  --output-dir    Output directory
  --progress, -p  Show progress bar
  --dry-run       Preview without converting
  --encoding      Character encoding (default: utf-8)

Filtering:
  --date-after    Only emails after date (YYYY-MM-DD)
  --date-before   Only emails before date
  --from-pattern  Filter by sender (regex)
  --to-pattern    Filter by recipient (regex)
  --subject-pattern  Filter by subject (regex)
  --body-contains Filter by body content
  --exclude-pattern  Exclude emails matching pattern
  --has-attachment   Filter by attachment presence
```

### Migrate Command

```bash
python mbox_converter.py migrate <preset> <input> [options]

Presets:
  google-to-365, thunderbird-to-outlook, outlook-to-thunderbird,
  apple-to-outlook, yahoo-to-gmail, outlook-to-gmail, eml-to-outlook,
  backup-to-csv, maildir-to-mbox, mbox-to-maildir
```

### Info Command

```bash
# Get file statistics
python mbox_converter.py info inbox.mbox

# JSON output
python mbox_converter.py info inbox.mbox --json
```

### List Command

```bash
# List emails
python mbox_converter.py list inbox.mbox --limit 50

# Search emails
python mbox_converter.py list inbox.mbox --subject-pattern "invoice"
```

## Examples

### 1. Google Workspace → Microsoft 365

```bash
# Step 1: Export from Google Takeout (creates MBOX files)
# Step 2: Convert to Outlook-compatible format
python mbox_converter.py migrate google-to-365 ./Takeout/*.mbox -o ./for_outlook
```

### 2. Thunderbird → Outlook

```bash
# Find Thunderbird profile folder and convert
python mbox_converter.py migrate thunderbird-to-outlook \
    "~/.thunderbird/profile/ImapMail/imap.gmail.com/INBOX" \
    -o ./outlook_import
```

### 3. Backup to CSV for Analysis

```bash
# Convert any format to CSV for spreadsheet analysis
python mbox_converter.py migrate backup-to-csv ./emails/* -o ./backup

# With filtering - only last year's emails
python mbox_converter.py convert inbox.mbox --format csv \
    --date-after 2024-01-01
```

### 4. Unix Server Migration (Maildir)

```bash
# Convert MBOX to Maildir for Dovecot
python mbox_converter.py convert inbox.mbox --format maildir -o ./maildir

# Convert Maildir back to MBOX
python mbox_converter.py convert ./maildir --format mbox
```

### 5. Batch Processing

```bash
# Convert all MBOX files in directory
python mbox_converter.py convert ./archive/*.mbox --format csv -o ./converted -p

# Convert with filtering
python mbox_converter.py convert ./emails/* --format json \
    --from-pattern "@company.com" \
    --date-after 2023-01-01
```

### 6. JSON for Programming

```bash
# Export to JSON for API/programmatic use
python mbox_converter.py convert inbox.mbox --format json

# Import JSON back to MBOX
python mbox_converter.py convert emails.json --format mbox
```

## Platform Notes

### Windows
- Full support for all formats including PST and MSG
- Requires Microsoft Outlook for PST operations
- Install pywin32: `pip install pywin32`

### Linux/macOS
- All formats except PST/MSG native write
- Can read MSG files with `extract-msg` package
- EML, MBOX, CSV, TXT, JSON, Maildir fully supported

### Cloud/CI
- Use for automated email processing pipelines
- JSON format ideal for API integration
- CSV for data analysis workflows

## Troubleshooting

### "pywin32 not found"
```bash
pip install pywin32
```
Note: Only needed for PST/MSG on Windows

### "extract-msg not found"
```bash
pip install extract-msg
```
Note: Only needed for reading MSG files on non-Windows

### Unicode errors
```bash
python mbox_converter.py convert inbox.mbox --format csv --encoding latin-1
```

### Large files
Use progress bar and filtering:
```bash
python mbox_converter.py convert large.mbox --format csv -p --date-after 2023-01-01
```

## Legacy Mode

For backward compatibility, old syntax still works:

```bash
python mbox_converter.py inbox.mbox --csv output.csv --mbox copy.mbox --json data.json
```

## License

MIT License - See [LICENSE](LICENSE) for details.

## Contributing

Contributions welcome! Please submit issues and pull requests on GitHub.
