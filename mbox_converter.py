#!/usr/bin/env python3
"""
Email Converter v4.0 - Universal Email Migration Tool

Supports bidirectional conversion between all major email formats:
- MBOX (Mozilla Thunderbird, Google Takeout)
- PST (Microsoft Outlook)
- EML (Standard RFC 822)
- MSG (Outlook individual messages)
- CSV (Spreadsheet format)
- TXT (Plain text)
- Maildir (Unix/Linux mail directories)
- JSON (API/programmatic format)

Built-in Migration Presets:
- Google Workspace → Microsoft 365
- Thunderbird → Outlook
- Apple Mail → Outlook
- Yahoo/AOL → Gmail
- Any IMAP → Any format

Cross-platform: Windows, Linux, macOS, and cloud environments.
"""

import argparse
import csv
import glob
import json
import logging
import mailbox
import os
import re
import shutil
import struct
import sys
from datetime import datetime
from email import policy
from email.message import EmailMessage
from email.parser import Parser, BytesParser
from email.utils import parsedate_to_datetime, formatdate, make_msgid
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Tuple

try:
    from tqdm import tqdm
    TQDM_AVAILABLE = True
except ImportError:
    TQDM_AVAILABLE = False

try:
    from dateutil import parser as date_parser
    DATEUTIL_AVAILABLE = True
except ImportError:
    DATEUTIL_AVAILABLE = False

# Version
__version__ = "4.0.0"

# Exit codes
EXIT_SUCCESS = 0
EXIT_ERROR = 1
EXIT_INVALID_ARGS = 2
EXIT_PARTIAL_SUCCESS = 3

# Supported formats
SUPPORTED_FORMATS = ["mbox", "csv", "eml", "txt", "pst", "msg", "maildir", "json"]

# Migration presets
MIGRATION_PRESETS = {
    "google-to-365": {
        "name": "Google Workspace → Microsoft 365",
        "description": "Convert Google Takeout MBOX to Outlook-compatible PST/EML",
        "input_format": "mbox",
        "output_format": "pst",
        "notes": "Use Google Takeout to export your Gmail, then convert here"
    },
    "thunderbird-to-outlook": {
        "name": "Thunderbird → Outlook",
        "description": "Convert Thunderbird MBOX profiles to Outlook PST",
        "input_format": "mbox",
        "output_format": "pst",
        "notes": "Find MBOX files in Thunderbird profile folder"
    },
    "outlook-to-thunderbird": {
        "name": "Outlook → Thunderbird",
        "description": "Convert Outlook PST to Thunderbird-compatible MBOX",
        "input_format": "pst",
        "output_format": "mbox",
        "notes": "Import resulting MBOX into Thunderbird"
    },
    "apple-to-outlook": {
        "name": "Apple Mail → Outlook",
        "description": "Convert Apple Mail MBOX/EMLX to Outlook",
        "input_format": "mbox",
        "output_format": "pst",
        "notes": "Export from Apple Mail as MBOX first"
    },
    "yahoo-to-gmail": {
        "name": "Yahoo/AOL → Gmail",
        "description": "Convert Yahoo/AOL mail export to Gmail-importable format",
        "input_format": "mbox",
        "output_format": "mbox",
        "notes": "Use POP3 or export tool to get MBOX, then import to Gmail"
    },
    "outlook-to-gmail": {
        "name": "Outlook → Gmail",
        "description": "Convert Outlook PST to Gmail-importable MBOX",
        "input_format": "pst",
        "output_format": "mbox",
        "notes": "Import resulting MBOX using Gmail's import feature"
    },
    "eml-to-outlook": {
        "name": "EML Files → Outlook",
        "description": "Batch convert EML files to Outlook PST",
        "input_format": "eml",
        "output_format": "pst",
        "notes": "Point to folder containing .eml files"
    },
    "backup-to-csv": {
        "name": "Email Backup → CSV",
        "description": "Convert any email format to CSV for analysis/archiving",
        "input_format": "auto",
        "output_format": "csv",
        "notes": "Great for email analytics, searching, or record-keeping"
    },
    "maildir-to-mbox": {
        "name": "Maildir → MBOX",
        "description": "Convert Unix Maildir format to portable MBOX",
        "input_format": "maildir",
        "output_format": "mbox",
        "notes": "Common for Dovecot, Postfix, Courier IMAP servers"
    },
    "mbox-to-maildir": {
        "name": "MBOX → Maildir",
        "description": "Convert MBOX to Unix Maildir format",
        "input_format": "mbox",
        "output_format": "maildir",
        "notes": "For migration to Dovecot or similar mail servers"
    },
}

# Logger setup
logger = logging.getLogger("email_converter")


def setup_logging(verbose: int = 0, log_file: Optional[str] = None, quiet: bool = False) -> None:
    """Configure logging based on verbosity level."""
    if quiet:
        level = logging.ERROR
    elif verbose >= 2:
        level = logging.DEBUG
    elif verbose >= 1:
        level = logging.INFO
    else:
        level = logging.WARNING

    formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    logger.handlers = []
    console_handler = logging.StreamHandler(sys.stderr)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    if log_file:
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    logger.setLevel(level)


# =============================================================================
# EMAIL MESSAGE UTILITIES
# =============================================================================

def get_body(message, encoding: str = "utf-8") -> str:
    """Extract plain text body from an email message."""
    if message.is_multipart():
        parts = []
        for part in message.walk():
            if part.get_content_type() == "text/plain" and not part.get_filename():
                data = part.get_payload(decode=True)
                if data:
                    charset = part.get_content_charset() or encoding
                    parts.append(data.decode(charset, errors="ignore"))
        return "\n".join(parts)
    payload = message.get_payload(decode=True)
    if isinstance(payload, bytes):
        return payload.decode(message.get_content_charset() or encoding, errors="ignore")
    return payload or ""


def get_html_body(message, encoding: str = "utf-8") -> str:
    """Extract HTML body from an email message."""
    if message.is_multipart():
        for part in message.walk():
            if part.get_content_type() == "text/html" and not part.get_filename():
                data = part.get_payload(decode=True)
                if data:
                    charset = part.get_content_charset() or encoding
                    return data.decode(charset, errors="ignore")
    return ""


def get_attachments(message) -> List[Dict[str, Any]]:
    """Extract attachment info from an email message."""
    attachments = []
    if message.is_multipart():
        for part in message.walk():
            filename = part.get_filename()
            if filename:
                attachments.append({
                    "filename": filename,
                    "content_type": part.get_content_type(),
                    "size": len(part.get_payload(decode=True) or b""),
                })
    return attachments


def parse_date(date_str: Optional[str]) -> Optional[datetime]:
    """Parse email date string to datetime object."""
    if not date_str:
        return None
    try:
        return parsedate_to_datetime(date_str)
    except (ValueError, TypeError):
        if DATEUTIL_AVAILABLE:
            try:
                return date_parser.parse(date_str, fuzzy=True)
            except (ValueError, TypeError):
                pass
        return None


def parse_filter_date(date_str: str) -> datetime:
    """Parse user-provided date filter string."""
    formats = [
        "%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    
    if DATEUTIL_AVAILABLE:
        try:
            return date_parser.parse(date_str)
        except (ValueError, TypeError):
            pass
    
    raise ValueError(f"Cannot parse date: {date_str}")


def create_email_message(
    from_addr: str = "",
    to_addr: str = "",
    cc_addr: str = "",
    bcc_addr: str = "",
    subject: str = "",
    date_str: str = "",
    body: str = "",
    html_body: str = "",
    encoding: str = "utf-8"
) -> EmailMessage:
    """Create an EmailMessage from components."""
    msg = EmailMessage()
    msg["From"] = from_addr
    msg["To"] = to_addr
    if cc_addr:
        msg["Cc"] = cc_addr
    if bcc_addr:
        msg["Bcc"] = bcc_addr
    msg["Subject"] = subject
    msg["Date"] = date_str if date_str else formatdate(localtime=True)
    msg["Message-ID"] = make_msgid()
    
    if html_body and body:
        msg.make_alternative()
        msg.add_alternative(body, subtype="plain")
        msg.add_alternative(html_body, subtype="html")
    elif html_body:
        msg.set_content(html_body, subtype="html")
    else:
        msg.set_content(body, charset=encoding)
    
    return msg


def message_to_dict(message, encoding: str = "utf-8") -> Dict[str, Any]:
    """Convert email message to dictionary."""
    attachments = get_attachments(message)
    return {
        "from": message.get("from", ""),
        "to": message.get("to", ""),
        "cc": message.get("cc", ""),
        "bcc": message.get("bcc", ""),
        "subject": message.get("subject", ""),
        "date": message.get("date", ""),
        "message_id": message.get("message-id", ""),
        "body": get_body(message, encoding),
        "html_body": get_html_body(message, encoding),
        "has_attachments": len(attachments) > 0,
        "attachments": attachments,
        "headers": dict(message.items()),
    }


# =============================================================================
# EMAIL FILTER
# =============================================================================

class EmailFilter:
    """Filter emails based on various criteria."""

    def __init__(
        self,
        date_after: Optional[datetime] = None,
        date_before: Optional[datetime] = None,
        from_pattern: Optional[str] = None,
        to_pattern: Optional[str] = None,
        subject_pattern: Optional[str] = None,
        body_contains: Optional[str] = None,
        has_attachment: Optional[bool] = None,
        exclude_pattern: Optional[str] = None,
    ):
        self.date_after = date_after
        self.date_before = date_before
        self.from_regex = re.compile(from_pattern, re.IGNORECASE) if from_pattern else None
        self.to_regex = re.compile(to_pattern, re.IGNORECASE) if to_pattern else None
        self.subject_regex = re.compile(subject_pattern, re.IGNORECASE) if subject_pattern else None
        self.body_contains = body_contains.lower() if body_contains else None
        self.has_attachment = has_attachment
        self.exclude_regex = re.compile(exclude_pattern, re.IGNORECASE) if exclude_pattern else None

    def matches(self, message, encoding: str = "utf-8") -> bool:
        """Check if message matches all filter criteria."""
        # Exclusion filter
        if self.exclude_regex:
            subject = message.get("subject", "")
            from_field = message.get("from", "")
            if self.exclude_regex.search(subject) or self.exclude_regex.search(from_field):
                return False

        # Date filters
        if self.date_after or self.date_before:
            msg_date = parse_date(message.get("date"))
            if msg_date:
                msg_date_naive = msg_date.replace(tzinfo=None) if msg_date.tzinfo else msg_date
                date_after_naive = self.date_after.replace(tzinfo=None) if self.date_after and self.date_after.tzinfo else self.date_after
                date_before_naive = self.date_before.replace(tzinfo=None) if self.date_before and self.date_before.tzinfo else self.date_before
                
                if date_after_naive and msg_date_naive < date_after_naive:
                    return False
                if date_before_naive and msg_date_naive > date_before_naive:
                    return False
            elif self.date_after or self.date_before:
                return False

        if self.from_regex:
            from_field = message.get("from", "")
            if not self.from_regex.search(from_field):
                return False

        if self.to_regex:
            to_field = message.get("to", "")
            if not self.to_regex.search(to_field):
                return False

        if self.subject_regex:
            subject = message.get("subject", "")
            if not self.subject_regex.search(subject):
                return False

        if self.body_contains:
            body = get_body(message, encoding).lower()
            if self.body_contains not in body:
                return False

        if self.has_attachment is not None:
            has_attach = any(
                part.get_filename() for part in message.walk()
            ) if message.is_multipart() else False
            if has_attach != self.has_attachment:
                return False

        return True

    def matches_dict(self, email_dict: Dict[str, Any], encoding: str = "utf-8") -> bool:
        """Check if email dictionary matches filter criteria."""
        # Exclusion filter
        if self.exclude_regex:
            subject = email_dict.get("subject", "") or email_dict.get("Subject", "")
            from_field = email_dict.get("from", "") or email_dict.get("From", "")
            if self.exclude_regex.search(subject) or self.exclude_regex.search(from_field):
                return False

        if self.date_after or self.date_before:
            date_str = email_dict.get("date", "") or email_dict.get("Date", "")
            msg_date = parse_date(date_str)
            if msg_date:
                msg_date_naive = msg_date.replace(tzinfo=None) if msg_date.tzinfo else msg_date
                date_after_naive = self.date_after.replace(tzinfo=None) if self.date_after and self.date_after.tzinfo else self.date_after
                date_before_naive = self.date_before.replace(tzinfo=None) if self.date_before and self.date_before.tzinfo else self.date_before
                
                if date_after_naive and msg_date_naive < date_after_naive:
                    return False
                if date_before_naive and msg_date_naive > date_before_naive:
                    return False
            elif self.date_after or self.date_before:
                return False

        if self.from_regex:
            from_field = email_dict.get("from", "") or email_dict.get("From", "")
            if not self.from_regex.search(from_field):
                return False

        if self.to_regex:
            to_field = email_dict.get("to", "") or email_dict.get("To", "")
            if not self.to_regex.search(to_field):
                return False

        if self.subject_regex:
            subject = email_dict.get("subject", "") or email_dict.get("Subject", "")
            if not self.subject_regex.search(subject):
                return False

        if self.body_contains:
            body = (email_dict.get("body", "") or email_dict.get("Body", "")).lower()
            if self.body_contains not in body:
                return False

        return True


def build_filter_from_args(args) -> Optional[EmailFilter]:
    """Build EmailFilter from command line arguments."""
    has_filters = any([
        getattr(args, 'date_after', None),
        getattr(args, 'date_before', None),
        getattr(args, 'from_pattern', None),
        getattr(args, 'to_pattern', None),
        getattr(args, 'subject_pattern', None),
        getattr(args, 'body_contains', None),
        getattr(args, 'has_attachment', None) is not None,
        getattr(args, 'exclude_pattern', None),
    ])
    
    if not has_filters:
        return None
    
    date_after = parse_filter_date(args.date_after) if getattr(args, 'date_after', None) else None
    date_before = parse_filter_date(args.date_before) if getattr(args, 'date_before', None) else None
    
    return EmailFilter(
        date_after=date_after,
        date_before=date_before,
        from_pattern=getattr(args, 'from_pattern', None),
        to_pattern=getattr(args, 'to_pattern', None),
        subject_pattern=getattr(args, 'subject_pattern', None),
        body_contains=getattr(args, 'body_contains', None),
        has_attachment=getattr(args, 'has_attachment', None),
        exclude_pattern=getattr(args, 'exclude_pattern', None),
    )


# =============================================================================
# FORMAT READERS
# =============================================================================

def read_from_mbox(
    path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> Iterator[Tuple[int, Any]]:
    """Read emails from MBOX file."""
    mbox = mailbox.mbox(path)
    total = len(mbox)
    
    items = list(enumerate(mbox))
    if show_progress and TQDM_AVAILABLE and not quiet:
        items = tqdm(items, desc=f"Reading {Path(path).name}", unit="emails", ncols=80)
    
    filtered_count = 0
    for idx, msg in items:
        if email_filter is None or email_filter.matches(msg, encoding):
            filtered_count += 1
            yield idx, msg
    
    mbox.close()
    logger.info(f"Read {total} emails from MBOX, {filtered_count} matched filters")


def read_from_maildir(
    path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> Iterator[Tuple[int, Any]]:
    """Read emails from Maildir directory."""
    md = mailbox.Maildir(path)
    total = len(md)
    
    items = list(enumerate(md))
    if show_progress and TQDM_AVAILABLE and not quiet:
        items = tqdm(items, desc=f"Reading {Path(path).name}", unit="emails", ncols=80)
    
    filtered_count = 0
    for idx, msg in items:
        if email_filter is None or email_filter.matches(msg, encoding):
            filtered_count += 1
            yield idx, msg
    
    md.close()
    logger.info(f"Read {total} emails from Maildir, {filtered_count} matched filters")


def read_from_csv(
    path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> Iterator[Tuple[int, Any]]:
    """Read emails from CSV file."""
    with open(path, "r", encoding=encoding, newline="") as fh:
        reader = csv.DictReader(fh)
        rows = list(reader)
    
    if show_progress and TQDM_AVAILABLE and not quiet:
        rows = tqdm(rows, desc=f"Reading {Path(path).name}", unit="emails", ncols=80)
    
    filtered_count = 0
    for idx, row in enumerate(rows):
        normalized = {k.lower(): v for k, v in row.items()}
        
        if email_filter is not None and not email_filter.matches_dict(normalized, encoding):
            continue
        
        msg = create_email_message(
            from_addr=normalized.get("from", ""),
            to_addr=normalized.get("to", ""),
            cc_addr=normalized.get("cc", ""),
            bcc_addr=normalized.get("bcc", ""),
            subject=normalized.get("subject", ""),
            date_str=normalized.get("date", ""),
            body=normalized.get("body", ""),
            html_body=normalized.get("html_body", ""),
            encoding=encoding,
        )
        filtered_count += 1
        yield idx, msg
    
    logger.info(f"Read emails from CSV, {filtered_count} matched filters")


def read_from_json(
    path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> Iterator[Tuple[int, Any]]:
    """Read emails from JSON file."""
    with open(path, "r", encoding=encoding) as fh:
        data = json.load(fh)
    
    if isinstance(data, dict) and "emails" in data:
        emails = data["emails"]
    elif isinstance(data, list):
        emails = data
    else:
        raise ValueError("JSON must be a list of emails or object with 'emails' key")
    
    if show_progress and TQDM_AVAILABLE and not quiet:
        emails = tqdm(emails, desc=f"Reading {Path(path).name}", unit="emails", ncols=80)
    
    filtered_count = 0
    for idx, email_dict in enumerate(emails):
        if email_filter is not None and not email_filter.matches_dict(email_dict, encoding):
            continue
        
        msg = create_email_message(
            from_addr=email_dict.get("from", ""),
            to_addr=email_dict.get("to", ""),
            cc_addr=email_dict.get("cc", ""),
            bcc_addr=email_dict.get("bcc", ""),
            subject=email_dict.get("subject", ""),
            date_str=email_dict.get("date", ""),
            body=email_dict.get("body", ""),
            html_body=email_dict.get("html_body", ""),
            encoding=encoding,
        )
        filtered_count += 1
        yield idx, msg
    
    logger.info(f"Read emails from JSON, {filtered_count} matched filters")


def read_from_eml_directory(
    path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> Iterator[Tuple[int, Any]]:
    """Read emails from directory of EML files."""
    if os.path.isfile(path) and path.lower().endswith(".eml"):
        eml_files = [path]
    elif os.path.isdir(path):
        eml_files = sorted(glob.glob(os.path.join(path, "*.eml")))
    else:
        eml_files = sorted(glob.glob(path))
    
    if show_progress and TQDM_AVAILABLE and not quiet:
        eml_files = tqdm(eml_files, desc="Reading EML files", unit="files", ncols=80)
    
    filtered_count = 0
    for idx, eml_path in enumerate(eml_files):
        try:
            with open(eml_path, "rb") as fh:
                msg = BytesParser(policy=policy.default).parse(fh)
            
            if email_filter is None or email_filter.matches(msg, encoding):
                filtered_count += 1
                yield idx, msg
        except Exception as e:
            logger.warning(f"Failed to read {eml_path}: {e}")
    
    logger.info(f"Read EML files, {filtered_count} matched filters")


def read_from_msg(
    path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> Iterator[Tuple[int, Any]]:
    """Read emails from MSG files (Outlook format)."""
    if os.path.isfile(path) and path.lower().endswith(".msg"):
        msg_files = [path]
    elif os.path.isdir(path):
        msg_files = sorted(glob.glob(os.path.join(path, "*.msg")))
    else:
        msg_files = sorted(glob.glob(path))
    
    if show_progress and TQDM_AVAILABLE and not quiet:
        msg_files = tqdm(msg_files, desc="Reading MSG files", unit="files", ncols=80)
    
    filtered_count = 0
    for idx, msg_path in enumerate(msg_files):
        try:
            msg = parse_msg_file(msg_path, encoding)
            if msg and (email_filter is None or email_filter.matches(msg, encoding)):
                filtered_count += 1
                yield idx, msg
        except Exception as e:
            logger.warning(f"Failed to read {msg_path}: {e}")
    
    logger.info(f"Read MSG files, {filtered_count} matched filters")


def parse_msg_file(path: str, encoding: str = "utf-8") -> Optional[EmailMessage]:
    """Parse Outlook MSG file (simplified parser)."""
    # Try to use extract_msg if available
    try:
        import extract_msg
        msg = extract_msg.Message(path)
        email_msg = create_email_message(
            from_addr=msg.sender or "",
            to_addr=msg.to or "",
            cc_addr=msg.cc or "",
            subject=msg.subject or "",
            date_str=str(msg.date) if msg.date else "",
            body=msg.body or "",
            html_body=msg.htmlBody or "",
            encoding=encoding,
        )
        msg.close()
        return email_msg
    except ImportError:
        pass
    
    # Fallback: basic OLE parsing
    try:
        import olefile
        ole = olefile.OleFileIO(path)
        
        def get_stream(name):
            try:
                if ole.exists(name):
                    return ole.openstream(name).read()
            except Exception:
                pass
            return b""
        
        subject = get_stream("__substg1.0_0037001F").decode("utf-16-le", errors="ignore").rstrip("\x00")
        sender = get_stream("__substg1.0_0C1F001F").decode("utf-16-le", errors="ignore").rstrip("\x00")
        to = get_stream("__substg1.0_0E04001F").decode("utf-16-le", errors="ignore").rstrip("\x00")
        body = get_stream("__substg1.0_1000001F").decode("utf-16-le", errors="ignore").rstrip("\x00")
        
        ole.close()
        
        return create_email_message(
            from_addr=sender,
            to_addr=to,
            subject=subject,
            body=body,
            encoding=encoding,
        )
    except ImportError:
        logger.warning("Install 'extract-msg' or 'olefile' for MSG support: pip install extract-msg")
        return None
    except Exception as e:
        logger.warning(f"Failed to parse MSG: {e}")
        return None


def read_from_txt(
    path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> Iterator[Tuple[int, Any]]:
    """Read emails from TXT file."""
    with open(path, "r", encoding=encoding) as fh:
        content = fh.read()
    
    email_blocks = re.split(r"={80}\nEmail #\d+\n-{40}\n", content)
    email_blocks = [b.strip() for b in email_blocks if b.strip()]
    
    if show_progress and TQDM_AVAILABLE and not quiet:
        email_blocks = tqdm(email_blocks, desc=f"Reading {Path(path).name}", unit="emails", ncols=80)
    
    filtered_count = 0
    for idx, block in enumerate(email_blocks):
        try:
            lines = block.split("\n")
            headers = {}
            body_start = 0
            
            for i, line in enumerate(lines):
                if line.startswith("-" * 40):
                    body_start = i + 1
                    break
                if ": " in line:
                    key, value = line.split(": ", 1)
                    headers[key.lower()] = value
            
            body = "\n".join(lines[body_start:]).strip()
            
            msg = create_email_message(
                from_addr=headers.get("from", ""),
                to_addr=headers.get("to", ""),
                cc_addr=headers.get("cc", ""),
                subject=headers.get("subject", ""),
                date_str=headers.get("date", ""),
                body=body,
                encoding=encoding,
            )
            
            if email_filter is None or email_filter.matches(msg, encoding):
                filtered_count += 1
                yield idx, msg
        except Exception as e:
            logger.warning(f"Failed to parse email block {idx}: {e}")
    
    logger.info(f"Read emails from TXT, {filtered_count} matched filters")


def read_from_pst(
    path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> Iterator[Tuple[int, Any]]:
    """Read emails from PST file (Windows + Outlook only)."""
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "pywin32 is required for PST reading. "
            "Install with: pip install pywin32\n"
            "Note: PST operations only work on Windows with Outlook installed."
        )
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    outlook.AddStoreEx(os.path.abspath(path), 3)
    
    pst_root = None
    for folder in outlook.Folders:
        try:
            if hasattr(folder, "Store") and getattr(folder.Store, "FilePath", None):
                if os.path.abspath(folder.Store.FilePath) == os.path.abspath(path):
                    pst_root = folder
                    break
        except Exception:
            continue
    
    if pst_root is None:
        raise RuntimeError("Failed to open PST file")
    
    def iterate_folder(folder, idx_start=0):
        idx = idx_start
        try:
            for item in folder.Items:
                try:
                    msg = create_email_message(
                        from_addr=getattr(item, "SenderEmailAddress", "") or "",
                        to_addr=getattr(item, "To", "") or "",
                        cc_addr=getattr(item, "CC", "") or "",
                        subject=getattr(item, "Subject", "") or "",
                        date_str=str(getattr(item, "ReceivedTime", "")) or "",
                        body=getattr(item, "Body", "") or "",
                        html_body=getattr(item, "HTMLBody", "") or "",
                        encoding=encoding,
                    )
                    if email_filter is None or email_filter.matches(msg, encoding):
                        yield idx, msg
                    idx += 1
                except Exception:
                    continue
        except Exception:
            pass
        
        try:
            for subfolder in folder.Folders:
                yield from iterate_folder(subfolder, idx)
        except Exception:
            pass
    
    yield from iterate_folder(pst_root)
    outlook.RemoveStore(pst_root)


# =============================================================================
# FORMAT WRITERS
# =============================================================================

def write_to_mbox(
    emails: Iterator[Tuple[int, Any]],
    output_path: str,
    encoding: str = "utf-8",
) -> int:
    """Write emails to MBOX file."""
    mbox = mailbox.mbox(output_path)
    mbox.lock()
    count = 0
    
    try:
        for idx, msg in emails:
            mbox.add(msg)
            count += 1
    finally:
        mbox.unlock()
        mbox.close()
    
    return count


def write_to_maildir(
    emails: Iterator[Tuple[int, Any]],
    output_path: str,
    encoding: str = "utf-8",
) -> int:
    """Write emails to Maildir directory."""
    # Ensure Maildir structure exists
    for subdir in ["cur", "new", "tmp"]:
        os.makedirs(os.path.join(output_path, subdir), exist_ok=True)
    
    md = mailbox.Maildir(output_path, create=True)
    count = 0
    
    try:
        for idx, msg in emails:
            md.add(msg)
            count += 1
    finally:
        md.close()
    
    return count


def write_to_csv(
    emails: Iterator[Tuple[int, Any]],
    output_path: str,
    encoding: str = "utf-8",
) -> int:
    """Write emails to CSV file."""
    count = 0
    
    with open(output_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(["Index", "From", "To", "Cc", "Bcc", "Subject", "Date", "Body", "HTML_Body", "HasAttachment", "Attachments"])
        
        for idx, msg in emails:
            attachments = get_attachments(msg)
            attachment_names = "; ".join([a["filename"] for a in attachments])
            writer.writerow([
                idx,
                msg.get("from", ""),
                msg.get("to", ""),
                msg.get("cc", ""),
                msg.get("bcc", ""),
                msg.get("subject", ""),
                msg.get("date", ""),
                get_body(msg, encoding).replace("\r", "").replace("\n", " "),
                get_html_body(msg, encoding).replace("\r", "").replace("\n", " ")[:500],  # Truncate HTML
                len(attachments) > 0,
                attachment_names,
            ])
            count += 1
    
    return count


def write_to_json(
    emails: Iterator[Tuple[int, Any]],
    output_path: str,
    encoding: str = "utf-8",
) -> int:
    """Write emails to JSON file."""
    email_list = []
    
    for idx, msg in emails:
        email_list.append(message_to_dict(msg, encoding))
    
    with open(output_path, "w", encoding="utf-8") as fh:
        json.dump({"emails": email_list, "count": len(email_list)}, fh, indent=2, default=str)
    
    return len(email_list)


def write_to_eml(
    emails: Iterator[Tuple[int, Any]],
    output_dir: str,
    encoding: str = "utf-8",
) -> int:
    """Write emails to individual EML files."""
    os.makedirs(output_dir, exist_ok=True)
    count = 0
    
    for idx, msg in emails:
        subject = msg.get("subject", "no_subject") or "no_subject"
        safe_subject = re.sub(r'[<>:"/\\|?*]', '_', subject)[:50]
        filename = f"{idx:06d}_{safe_subject}.eml"
        filepath = os.path.join(output_dir, filename)
        
        with open(filepath, "wb") as fh:
            if hasattr(msg, 'as_bytes'):
                fh.write(msg.as_bytes())
            else:
                fh.write(str(msg).encode(encoding))
        count += 1
    
    return count


def write_to_msg(
    emails: Iterator[Tuple[int, Any]],
    output_dir: str,
    encoding: str = "utf-8",
) -> int:
    """Write emails to MSG files (requires Windows + Outlook)."""
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError("pywin32 required for MSG output. Install: pip install pywin32")
    
    os.makedirs(output_dir, exist_ok=True)
    outlook = win32com.client.Dispatch("Outlook.Application")
    count = 0
    
    for idx, msg in emails:
        try:
            mail = outlook.CreateItem(0)  # 0 = Mail item
            mail.Subject = msg.get("subject", "")
            mail.To = msg.get("to", "")
            mail.CC = msg.get("cc", "")
            mail.Body = get_body(msg, encoding)
            
            subject = msg.get("subject", "no_subject") or "no_subject"
            safe_subject = re.sub(r'[<>:"/\\|?*]', '_', subject)[:50]
            filename = f"{idx:06d}_{safe_subject}.msg"
            filepath = os.path.join(output_dir, filename)
            
            mail.SaveAs(filepath, 3)  # 3 = olMSG format
            count += 1
        except Exception as e:
            logger.warning(f"Failed to create MSG: {e}")
    
    return count


def write_to_txt(
    emails: Iterator[Tuple[int, Any]],
    output_path: str,
    encoding: str = "utf-8",
) -> int:
    """Write emails to plain text file."""
    count = 0
    
    with open(output_path, "w", encoding="utf-8") as fh:
        for idx, msg in emails:
            fh.write("=" * 80 + "\n")
            fh.write(f"Email #{idx}\n")
            fh.write("-" * 40 + "\n")
            fh.write(f"From: {msg.get('from', '')}\n")
            fh.write(f"To: {msg.get('to', '')}\n")
            fh.write(f"Cc: {msg.get('cc', '')}\n")
            fh.write(f"Subject: {msg.get('subject', '')}\n")
            fh.write(f"Date: {msg.get('date', '')}\n")
            fh.write("-" * 40 + "\n")
            fh.write(get_body(msg, encoding))
            fh.write("\n\n")
            count += 1
    
    return count


def write_to_pst(
    emails: Iterator[Tuple[int, Any]],
    output_path: str,
    encoding: str = "utf-8",
) -> int:
    """Write emails to PST file (Windows + Outlook only)."""
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "pywin32 is required for PST conversion. "
            "Install with: pip install pywin32\n"
            "Note: PST conversion only works on Windows with Outlook installed."
        )

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    outlook.AddStoreEx(os.path.abspath(output_path), 3)
    
    pst_root = None
    for folder in outlook.Folders:
        try:
            if hasattr(folder, "Store") and getattr(folder.Store, "FilePath", None):
                if os.path.abspath(folder.Store.FilePath) == os.path.abspath(output_path):
                    pst_root = folder
                    break
        except Exception:
            continue

    if pst_root is None:
        raise RuntimeError("Failed to create PST store")

    target = pst_root.Folders.Add("Imported")
    count = 0
    
    for idx, msg in emails:
        item = target.Items.Add()
        item.Subject = msg.get("subject", "")
        item.To = msg.get("to", "")
        item.CC = msg.get("cc", "")
        item.BCC = msg.get("bcc", "")
        item.Body = get_body(msg, encoding)
        html = get_html_body(msg, encoding)
        if html:
            item.HTMLBody = html
        item.Save()
        count += 1
    
    outlook.RemoveStore(pst_root)
    return count


# =============================================================================
# FORMAT DETECTION AND ROUTING
# =============================================================================

def detect_format(path: str) -> str:
    """Detect format from file extension or content."""
    path_lower = path.lower()
    
    if path_lower.endswith(".mbox"):
        return "mbox"
    elif path_lower.endswith(".csv"):
        return "csv"
    elif path_lower.endswith(".eml"):
        return "eml"
    elif path_lower.endswith(".msg"):
        return "msg"
    elif path_lower.endswith(".txt"):
        return "txt"
    elif path_lower.endswith(".pst"):
        return "pst"
    elif path_lower.endswith(".json"):
        return "json"
    elif os.path.isdir(path):
        # Check directory type
        if glob.glob(os.path.join(path, "*.eml")):
            return "eml"
        if glob.glob(os.path.join(path, "*.msg")):
            return "msg"
        # Check for Maildir structure
        if all(os.path.isdir(os.path.join(path, d)) for d in ["cur", "new", "tmp"] if os.path.exists(os.path.join(path, d))):
            return "maildir"
    
    # Content-based detection
    if os.path.isfile(path):
        try:
            with open(path, "rb") as fh:
                header = fh.read(200)
            
            if header.startswith(b"From "):
                return "mbox"
            if b"Index,From,To" in header or b"from,to,subject" in header.lower():
                return "csv"
            if header.strip().startswith(b"{") or header.strip().startswith(b"["):
                return "json"
        except Exception:
            pass
    
    return "unknown"


def get_reader(fmt: str):
    """Get reader function for format."""
    readers = {
        "mbox": read_from_mbox,
        "csv": read_from_csv,
        "eml": read_from_eml_directory,
        "msg": read_from_msg,
        "txt": read_from_txt,
        "pst": read_from_pst,
        "json": read_from_json,
        "maildir": read_from_maildir,
    }
    return readers.get(fmt)


def get_writer(fmt: str):
    """Get writer function for format."""
    writers = {
        "mbox": write_to_mbox,
        "csv": write_to_csv,
        "eml": write_to_eml,
        "msg": write_to_msg,
        "txt": write_to_txt,
        "pst": write_to_pst,
        "json": write_to_json,
        "maildir": write_to_maildir,
    }
    return writers.get(fmt)


# =============================================================================
# CONVERSION FUNCTION
# =============================================================================

def convert(
    input_path: str,
    output_path: str,
    input_format: Optional[str] = None,
    output_format: Optional[str] = None,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> int:
    """Universal conversion between any formats."""
    
    if input_format is None:
        input_format = detect_format(input_path)
        if input_format == "unknown":
            raise ValueError(f"Cannot detect input format: {input_path}")
    
    if output_format is None:
        output_format = detect_format(output_path)
        if output_format == "unknown":
            ext = Path(output_path).suffix.lower().lstrip(".")
            if ext in SUPPORTED_FORMATS:
                output_format = ext
            else:
                raise ValueError(f"Cannot detect output format: {output_path}")
    
    logger.info(f"Converting {input_format.upper()} -> {output_format.upper()}")
    
    reader = get_reader(input_format)
    writer = get_writer(output_format)
    
    if reader is None:
        raise ValueError(f"Unsupported input format: {input_format}")
    if writer is None:
        raise ValueError(f"Unsupported output format: {output_format}")
    
    # For directory outputs
    if output_format in ["eml", "msg", "maildir"]:
        os.makedirs(output_path, exist_ok=True)
    
    emails = reader(input_path, email_filter, encoding, show_progress, quiet)
    count = writer(emails, output_path, encoding)
    
    return count


# =============================================================================
# FILE INFO
# =============================================================================

def get_file_info(path: str, encoding: str = "utf-8") -> Dict[str, Any]:
    """Get information about an email file/directory."""
    fmt = detect_format(path)
    
    info = {
        "path": path,
        "format": fmt,
        "format_name": get_format_name(fmt),
        "file_size_bytes": 0,
        "file_size_mb": 0.0,
        "total_emails": 0,
        "unique_senders": 0,
        "date_range": None,
    }
    
    if os.path.isfile(path):
        info["file_size_bytes"] = os.path.getsize(path)
        info["file_size_mb"] = round(info["file_size_bytes"] / (1024 * 1024), 2)
    elif os.path.isdir(path):
        total_size = sum(
            os.path.getsize(os.path.join(root, f))
            for root, dirs, files in os.walk(path)
            for f in files
        )
        info["file_size_bytes"] = total_size
        info["file_size_mb"] = round(total_size / (1024 * 1024), 2)
    
    reader = get_reader(fmt)
    if reader is None:
        return info
    
    senders = set()
    dates = []
    count = 0
    
    try:
        for idx, msg in reader(path, None, encoding, False, True):
            count += 1
            from_addr = msg.get("from", "")
            if from_addr:
                senders.add(from_addr)
            date = parse_date(msg.get("date"))
            if date:
                dates.append(date)
    except Exception as e:
        logger.warning(f"Error reading {path}: {e}")
    
    info["total_emails"] = count
    info["unique_senders"] = len(senders)
    
    if dates:
        info["date_range"] = {
            "earliest": min(dates).isoformat(),
            "latest": max(dates).isoformat(),
        }
    
    return info


def get_format_name(fmt: str) -> str:
    """Get human-readable format name."""
    names = {
        "mbox": "MBOX (Thunderbird/Gmail)",
        "csv": "CSV (Spreadsheet)",
        "eml": "EML (RFC 822 Email)",
        "msg": "MSG (Outlook Message)",
        "txt": "TXT (Plain Text)",
        "pst": "PST (Outlook Storage)",
        "json": "JSON (Structured Data)",
        "maildir": "Maildir (Unix Mail)",
    }
    return names.get(fmt, fmt.upper())


# =============================================================================
# PATH EXPANSION
# =============================================================================

def expand_paths(patterns: List[str]) -> List[str]:
    """Expand glob patterns and directories."""
    paths = []
    
    for pattern in patterns:
        if "*" in pattern or "?" in pattern:
            expanded = glob.glob(pattern, recursive=True)
            paths.extend(expanded)
        elif os.path.isdir(pattern):
            for ext in [".mbox", ".csv", ".txt", ".pst", ".json"]:
                paths.extend(glob.glob(os.path.join(pattern, f"*{ext}")))
            if glob.glob(os.path.join(pattern, "*.eml")):
                paths.append(pattern)
            if glob.glob(os.path.join(pattern, "*.msg")):
                paths.append(pattern)
            # Check for Maildir
            if os.path.isdir(os.path.join(pattern, "cur")) or os.path.isdir(os.path.join(pattern, "new")):
                paths.append(pattern)
        elif os.path.isfile(pattern):
            paths.append(pattern)
        else:
            logger.warning(f"Path not found: {pattern}")
    
    return sorted(set(paths))


# =============================================================================
# COMMAND HANDLERS
# =============================================================================

def cmd_convert(args) -> int:
    """Handle convert subcommand."""
    setup_logging(args.verbose, args.log_file, args.quiet)
    
    input_paths = expand_paths(args.input)
    
    if not input_paths:
        logger.error("No input files found")
        return EXIT_ERROR
    
    logger.info(f"Found {len(input_paths)} file(s) to process")
    
    email_filter = build_filter_from_args(args)
    if email_filter:
        logger.info("Email filters applied")
    
    if args.dry_run:
        print("\n=== DRY RUN MODE ===")
        print(f"Would process {len(input_paths)} file(s):")
        for f in input_paths:
            info = get_file_info(f, args.encoding)
            print(f"  - {f} [{info['format'].upper()}]: {info['total_emails']} emails ({info['file_size_mb']} MB)")
        print(f"\nOutput format: {args.format.upper()}")
        print(f"Output directory: {args.output_dir or 'current directory'}")
        return EXIT_SUCCESS
    
    results = []
    total_converted = 0
    errors = 0
    
    for input_path in input_paths:
        logger.info(f"Processing: {input_path}")
        
        try:
            input_format = detect_format(input_path)
            base_name = Path(input_path).stem
            output_dir = args.output_dir or os.path.dirname(input_path) or "."
            os.makedirs(output_dir, exist_ok=True)
            
            if args.format in ["eml", "msg", "maildir"]:
                output_path = os.path.join(output_dir, f"{base_name}_{args.format}")
            else:
                output_path = os.path.join(output_dir, f"{base_name}.{args.format}")
            
            count = convert(
                input_path=input_path,
                output_path=output_path,
                input_format=input_format,
                output_format=args.format,
                email_filter=email_filter,
                encoding=args.encoding,
                show_progress=args.progress,
                quiet=args.quiet,
            )
            
            total_converted += count
            results.append({
                "input": input_path,
                "input_format": input_format,
                "output": output_path,
                "output_format": args.format,
                "emails_converted": count,
                "status": "success",
            })
            
            if not args.quiet:
                print(f"✓ {input_path} [{input_format.upper()}] -> {output_path} [{args.format.upper()}] ({count} emails)")
        
        except Exception as e:
            logger.error(f"Error: {e}")
            errors += 1
            results.append({
                "input": input_path,
                "output": None,
                "emails_converted": 0,
                "status": "error",
                "error": str(e),
            })
    
    if not args.quiet:
        print("\n=== CONVERSION COMPLETE ===")
        print(f"Files processed: {len(input_paths)}")
        print(f"Total emails converted: {total_converted}")
        if errors:
            print(f"Errors: {errors}")
    
    if args.report:
        with open(args.report, "w", encoding="utf-8") as fh:
            json.dump({"summary": {"files_processed": len(input_paths), "total_converted": total_converted, "errors": errors}, "results": results}, fh, indent=2)
    
    return EXIT_ERROR if errors == len(input_paths) else (EXIT_PARTIAL_SUCCESS if errors > 0 else EXIT_SUCCESS)


def cmd_migrate(args) -> int:
    """Handle migrate subcommand with presets."""
    setup_logging(args.verbose, None, args.quiet)
    
    preset = MIGRATION_PRESETS.get(args.preset)
    if not preset:
        logger.error(f"Unknown preset: {args.preset}")
        return EXIT_ERROR
    
    print(f"\n=== {preset['name']} ===")
    print(f"Description: {preset['description']}")
    print(f"Note: {preset['notes']}\n")
    
    if not args.input:
        print("Usage: email_converter migrate <preset> <input_files> [--output-dir <dir>]")
        print(f"\nExample: email_converter migrate {args.preset} ./inbox.mbox --output-dir ./migrated")
        return EXIT_SUCCESS
    
    input_paths = expand_paths(args.input)
    if not input_paths:
        logger.error("No input files found")
        return EXIT_ERROR
    
    output_format = preset["output_format"]
    output_dir = args.output_dir or "./migrated"
    os.makedirs(output_dir, exist_ok=True)
    
    email_filter = build_filter_from_args(args)
    total_converted = 0
    
    for input_path in input_paths:
        try:
            base_name = Path(input_path).stem
            if output_format in ["eml", "msg", "maildir"]:
                output_path = os.path.join(output_dir, f"{base_name}_{output_format}")
            else:
                output_path = os.path.join(output_dir, f"{base_name}.{output_format}")
            
            count = convert(
                input_path=input_path,
                output_path=output_path,
                output_format=output_format,
                email_filter=email_filter,
                encoding=args.encoding,
                show_progress=args.progress,
                quiet=args.quiet,
            )
            
            total_converted += count
            if not args.quiet:
                print(f"✓ {input_path} -> {output_path} ({count} emails)")
        
        except Exception as e:
            logger.error(f"Error: {e}")
    
    print("\n=== Migration Complete ===")
    print(f"Total emails migrated: {total_converted}")
    print(f"Output directory: {output_dir}")
    
    return EXIT_SUCCESS


def cmd_info(args) -> int:
    """Handle info subcommand."""
    setup_logging(args.verbose, None, args.quiet)
    
    input_paths = expand_paths(args.input)
    
    if not input_paths:
        logger.error("No files found")
        return EXIT_ERROR
    
    all_info = []
    
    for path in input_paths:
        try:
            info = get_file_info(path, args.encoding)
            all_info.append(info)
            
            if not args.json:
                print(f"\n=== {path} ===")
                print(f"  Format: {info['format_name']}")
                print(f"  Total emails: {info['total_emails']}")
                print(f"  Unique senders: {info['unique_senders']}")
                print(f"  File size: {info['file_size_mb']} MB")
                if info.get("date_range"):
                    print(f"  Date range: {info['date_range']['earliest']} to {info['date_range']['latest']}")
        
        except Exception as e:
            logger.error(f"Error reading {path}: {e}")
    
    if args.json:
        print(json.dumps(all_info, indent=2))
    
    return EXIT_SUCCESS


def cmd_list(args) -> int:
    """Handle list subcommand."""
    setup_logging(args.verbose, None, args.quiet)
    
    input_paths = expand_paths(args.input)
    
    if not input_paths:
        logger.error("No files found")
        return EXIT_ERROR
    
    email_filter = build_filter_from_args(args)
    all_emails = []
    
    for path in input_paths:
        try:
            fmt = detect_format(path)
            reader = get_reader(fmt)
            
            if reader is None:
                continue
            
            count = 0
            for idx, msg in reader(path, email_filter, args.encoding, False, args.quiet):
                email_info = {
                    "file": path,
                    "format": fmt,
                    "index": idx,
                    "from": msg.get("from", ""),
                    "to": msg.get("to", ""),
                    "subject": msg.get("subject", ""),
                    "date": msg.get("date", ""),
                }
                all_emails.append(email_info)
                count += 1
                
                if not args.json and count <= args.limit:
                    print(f"[{idx}] {msg.get('date', 'No date')[:25]} | {msg.get('from', 'Unknown')[:30]} | {msg.get('subject', 'No subject')[:50]}")
                
                if count >= args.limit:
                    break
        
        except Exception as e:
            logger.error(f"Error: {e}")
    
    if args.json:
        print(json.dumps(all_emails[:args.limit], indent=2))
    
    return EXIT_SUCCESS


def cmd_formats(args) -> int:
    """Show supported formats and conversion matrix."""
    print("""
╔══════════════════════════════════════════════════════════════════════════════╗
║                    EMAIL CONVERTER v4.0 - SUPPORTED FORMATS                   ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║  FORMAT    EXT        READ  WRITE  DESCRIPTION                               ║
║  ────────────────────────────────────────────────────────────────────────    ║
║  MBOX      .mbox       ✓     ✓     Mozilla Thunderbird, Google Takeout      ║
║  CSV       .csv        ✓     ✓     Spreadsheet (Excel, Google Sheets)       ║
║  EML       .eml/*      ✓     ✓     RFC 822 Standard Email Files             ║
║  MSG       .msg/*      ✓     ✓     Microsoft Outlook Individual Messages    ║
║  TXT       .txt        ✓     ✓     Plain Text (Human Readable)              ║
║  PST       .pst        ✓     ✓     Outlook Personal Storage (Windows)       ║
║  JSON      .json       ✓     ✓     Structured Data (API/Programming)        ║
║  Maildir   folder/     ✓     ✓     Unix Mail Directory (Dovecot/Postfix)    ║
║                                                                              ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                           CONVERSION MATRIX                                   ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   FROM ↓ → TO    MBOX  CSV  EML  MSG  TXT  PST  JSON Maildir                ║
║   ─────────────────────────────────────────────────────────                  ║
║   MBOX            -     ✓    ✓    ✓*   ✓    ✓*   ✓    ✓                     ║
║   CSV             ✓     -    ✓    ✓*   ✓    ✓*   ✓    ✓                     ║
║   EML             ✓     ✓    -    ✓*   ✓    ✓*   ✓    ✓                     ║
║   MSG*            ✓     ✓    ✓    -    ✓    ✓*   ✓    ✓                     ║
║   TXT             ✓     ✓    ✓    ✓*   -    ✓*   ✓    ✓                     ║
║   PST*            ✓     ✓    ✓    ✓*   ✓    -    ✓    ✓                     ║
║   JSON            ✓     ✓    ✓    ✓*   ✓    ✓*   -    ✓                     ║
║   Maildir         ✓     ✓    ✓    ✓*   ✓    ✓*   ✓    -                     ║
║                                                                              ║
║   * = Requires Windows with Microsoft Outlook installed                      ║
║                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
""")
    return EXIT_SUCCESS


def cmd_presets(args) -> int:
    """Show available migration presets."""
    print("""
╔══════════════════════════════════════════════════════════════════════════════╗
║                      EASY MIGRATION PRESETS                                   ║
╠══════════════════════════════════════════════════════════════════════════════╣
""")
    
    for key, preset in MIGRATION_PRESETS.items():
        print(f"║  {key:25} │ {preset['name'][:45]}")
        print(f"║  {'':25} │ {preset['description'][:45]}")
        print(f"║  {'':25} │ Note: {preset['notes'][:39]}")
        print("║")
    
    print("""╠══════════════════════════════════════════════════════════════════════════════╣
║  USAGE:                                                                       ║
║    email_converter migrate <preset> <input_files> [--output-dir <dir>]        ║
║                                                                              ║
║  EXAMPLES:                                                                    ║
║    # Google to Microsoft 365                                                  ║
║    email_converter migrate google-to-365 ./Takeout/*.mbox -o ./for_outlook   ║
║                                                                              ║
║    # Thunderbird to Outlook                                                   ║
║    email_converter migrate thunderbird-to-outlook ./Inbox -o ./outlook_import║
║                                                                              ║
║    # Backup all emails to CSV                                                 ║
║    email_converter migrate backup-to-csv ./emails/* -o ./backup              ║
║                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
""")
    return EXIT_SUCCESS


def cmd_config(args) -> int:
    """Handle config subcommand."""
    if args.generate:
        sample = {
            "encoding": "utf-8",
            "verbose": 1,
            "progress": True,
            "filters": {
                "date_after": "2023-01-01",
                "date_before": None,
                "from_pattern": None,
                "to_pattern": None,
                "subject_pattern": None,
                "body_contains": None,
                "has_attachment": None,
                "exclude_pattern": None,
            },
            "output": {"format": "csv", "directory": "./output"}
        }
        
        with open(args.generate, "w", encoding="utf-8") as fh:
            json.dump(sample, fh, indent=2)
        print(f"Sample config written to {args.generate}")
        return EXIT_SUCCESS
    
    print("Use --generate <path> to create a sample config file")
    return EXIT_SUCCESS


# =============================================================================
# MAIN
# =============================================================================

def add_filter_args(parser):
    """Add filter arguments to a parser."""
    group = parser.add_argument_group("Filtering Options")
    group.add_argument("--date-after", help="Only emails after date (YYYY-MM-DD)")
    group.add_argument("--date-before", help="Only emails before date (YYYY-MM-DD)")
    group.add_argument("--from-pattern", help="Filter by sender (regex)")
    group.add_argument("--to-pattern", help="Filter by recipient (regex)")
    group.add_argument("--subject-pattern", help="Filter by subject (regex)")
    group.add_argument("--body-contains", help="Filter by body content")
    group.add_argument("--exclude-pattern", help="Exclude emails matching pattern")
    group.add_argument("--has-attachment", type=lambda x: x.lower() in ("true", "1", "yes"), nargs="?", const=True, help="Filter by attachment")


def main() -> int:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        prog="email_converter",
        description="Universal Email Converter - Convert between MBOX, CSV, EML, MSG, TXT, PST, JSON, Maildir",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
QUICK START:
  # Convert formats
  %(prog)s convert inbox.mbox --format csv
  %(prog)s convert emails.csv --format mbox
  
  # Easy migrations
  %(prog)s migrate google-to-365 ./Takeout/*.mbox
  %(prog)s migrate thunderbird-to-outlook ./Inbox
  
  # Show all presets
  %(prog)s presets
  
  # Show all formats
  %(prog)s formats
"""
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    
    subparsers = parser.add_subparsers(dest="command", help="Commands")
    
    # CONVERT
    convert_p = subparsers.add_parser("convert", help="Convert between email formats")
    convert_p.add_argument("input", nargs="+", help="Input file(s)")
    convert_p.add_argument("--format", "-f", choices=SUPPORTED_FORMATS, default="csv", help="Output format")
    convert_p.add_argument("--output-dir", "-o", help="Output directory")
    convert_p.add_argument("--encoding", "-e", default="utf-8", help="Encoding")
    convert_p.add_argument("--dry-run", action="store_true", help="Preview only")
    convert_p.add_argument("--progress", "-p", action="store_true", help="Show progress")
    convert_p.add_argument("--quiet", "-q", action="store_true", help="Quiet mode")
    convert_p.add_argument("--verbose", "-v", action="count", default=0, help="Verbose")
    convert_p.add_argument("--log-file", help="Log file")
    convert_p.add_argument("--report", help="JSON report file")
    add_filter_args(convert_p)
    
    # MIGRATE
    migrate_p = subparsers.add_parser("migrate", help="Easy migration presets")
    migrate_p.add_argument("preset", choices=list(MIGRATION_PRESETS.keys()), help="Migration preset")
    migrate_p.add_argument("input", nargs="*", help="Input file(s)")
    migrate_p.add_argument("--output-dir", "-o", help="Output directory")
    migrate_p.add_argument("--encoding", "-e", default="utf-8", help="Encoding")
    migrate_p.add_argument("--progress", "-p", action="store_true", help="Show progress")
    migrate_p.add_argument("--quiet", "-q", action="store_true", help="Quiet mode")
    migrate_p.add_argument("--verbose", "-v", action="count", default=0, help="Verbose")
    add_filter_args(migrate_p)
    
    # INFO
    info_p = subparsers.add_parser("info", help="Show file info")
    info_p.add_argument("input", nargs="+", help="Input file(s)")
    info_p.add_argument("--encoding", "-e", default="utf-8")
    info_p.add_argument("--json", action="store_true")
    info_p.add_argument("--quiet", "-q", action="store_true")
    info_p.add_argument("--verbose", "-v", action="count", default=0)
    
    # LIST
    list_p = subparsers.add_parser("list", help="List emails")
    list_p.add_argument("input", nargs="+", help="Input file(s)")
    list_p.add_argument("--limit", "-n", type=int, default=100)
    list_p.add_argument("--encoding", "-e", default="utf-8")
    list_p.add_argument("--json", action="store_true")
    list_p.add_argument("--quiet", "-q", action="store_true")
    list_p.add_argument("--verbose", "-v", action="count", default=0)
    add_filter_args(list_p)
    
    # FORMATS
    subparsers.add_parser("formats", help="Show supported formats")
    
    # PRESETS
    subparsers.add_parser("presets", help="Show migration presets")
    
    # CONFIG
    config_p = subparsers.add_parser("config", help="Configuration")
    config_p.add_argument("--generate", metavar="PATH", help="Generate sample config")
    
    # LEGACY MODE
    if len(sys.argv) > 1 and sys.argv[1] not in ["convert", "migrate", "info", "list", "formats", "presets", "config", "-h", "--help", "--version"]:
        legacy_parser = argparse.ArgumentParser(description="Email Converter (Legacy)")
        legacy_parser.add_argument("input", help="Input file")
        legacy_parser.add_argument("--csv", help="Output CSV")
        legacy_parser.add_argument("--mbox", help="Output MBOX")
        legacy_parser.add_argument("--eml", help="Output EML dir")
        legacy_parser.add_argument("--txt", help="Output TXT")
        legacy_parser.add_argument("--pst", help="Output PST")
        legacy_parser.add_argument("--json", dest="json_out", help="Output JSON")
        legacy_parser.add_argument("--msg", help="Output MSG dir")
        legacy_parser.add_argument("--maildir", help="Output Maildir")
        
        legacy_args = legacy_parser.parse_args()
        
        outputs = [
            ("csv", legacy_args.csv), ("mbox", legacy_args.mbox), ("eml", legacy_args.eml),
            ("txt", legacy_args.txt), ("pst", legacy_args.pst), ("json", legacy_args.json_out),
            ("msg", legacy_args.msg), ("maildir", legacy_args.maildir),
        ]
        outputs = [(fmt, path) for fmt, path in outputs if path]
        
        if not outputs:
            legacy_parser.error("Specify at least one output format")
        
        setup_logging(0, None, False)
        
        try:
            for fmt, output_path in outputs:
                count = convert(input_path=legacy_args.input, output_path=output_path, output_format=fmt)
                print(f"{fmt.upper()} written to {output_path} ({count} emails)")
            return EXIT_SUCCESS
        except Exception as e:
            logger.error(str(e))
            return EXIT_ERROR
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return EXIT_INVALID_ARGS
    
    commands = {
        "convert": cmd_convert,
        "migrate": cmd_migrate,
        "info": cmd_info,
        "list": cmd_list,
        "formats": cmd_formats,
        "presets": cmd_presets,
        "config": cmd_config,
    }
    
    return commands[args.command](args)


if __name__ == "__main__":
    sys.exit(main())
