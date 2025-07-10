# MBOX Converter

This project provides a simple command-line tool for converting Mozilla
`mbox` files to either CSV or Outlook PST format.

The converter is fully open source and distributed under the MIT license.
There are no limitations or subscriptions required.

## Requirements

* Python 3.8+
* For PST conversion you must run on Windows with Microsoft Outlook installed
  and have the `pywin32` package available.

## Installation

Create a virtual environment (optional) and install the required Python
packages:

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows use `.venv\Scripts\activate`
pip install -r requirements.txt
```

## Usage

Convert an MBOX file to CSV:

```bash
python mbox_converter.py input.mbox --csv output.csv
```

Convert an MBOX file to a PST (Windows + Outlook only):

```bash
python mbox_converter.py input.mbox --pst output.pst
```

## Building on Windows

The PST conversion relies on Outlook COM automation. Ensure Outlook is installed
and Python has access to the `pywin32` package. The script creates a new PST
file and imports all messages from the MBOX file.


### Building a Windows executable

To create a standalone `mbox_converter.exe` run the helper script
`build_exe.py`:

```bash
python build_exe.py
```

The script installs PyInstaller if necessary and produces
`dist\mbox_converter.exe`, which runs without requiring Python or
PowerShell on the target machine.

### Using the convenience batch script

Windows users can run `run_converter.bat` directly. The script uses the
built-in `curl` command to install Python if needed and then invokes
`mbox_converter.py` with any arguments you supply:

```cmd
run_converter.bat input.mbox --csv output.csv
```

