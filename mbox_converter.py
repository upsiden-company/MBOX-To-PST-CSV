#!/usr/bin/env python3
"""
MBOX Converter v3.0 - Universal Email Format Converter

Supports bidirectional conversion between:
- MBOX (Mozilla Thunderbird)
- CSV (Spreadsheet)
- EML (Individual email files)
- TXT (Plain text)
- PST (Outlook - Windows only)

Features batch processing, filtering, progress reporting, and logging.
Cross-platform compatible: Windows, Linux, macOS, and cloud environments.
"""

import argparse
import csv
import glob
import json
import logging
import mailbox
import os
import re
import sys
from datetime import datetime
from email import policy
from email.message import EmailMessage
from email.parser import Parser, BytesParser
from email.utils import parsedate_to_datetime, formatdate, make_msgid
from io import StringIO, BytesIO
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
__version__ = "3.0.0"

# Exit codes
EXIT_SUCCESS = 0
EXIT_ERROR = 1
EXIT_INVALID_ARGS = 2
EXIT_PARTIAL_SUCCESS = 3

# Supported formats
SUPPORTED_FORMATS = ["mbox", "csv", "eml", "txt", "pst"]

# Logger setup
logger = logging.getLogger("mbox_converter")


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

    # Clear existing handlers
    logger.handlers = []

    # Console handler
    console_handler = logging.StreamHandler(sys.stderr)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # File handler if specified
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
    msg.set_content(body, charset=encoding)
    return msg


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
    ):
        self.date_after = date_after
        self.date_before = date_before
        self.from_regex = re.compile(from_pattern, re.IGNORECASE) if from_pattern else None
        self.to_regex = re.compile(to_pattern, re.IGNORECASE) if to_pattern else None
        self.subject_regex = re.compile(subject_pattern, re.IGNORECASE) if subject_pattern else None
        self.body_contains = body_contains.lower() if body_contains else None
        self.has_attachment = has_attachment

    def matches(self, message, encoding: str = "utf-8") -> bool:
        """Check if message matches all filter criteria."""
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

        # From filter
        if self.from_regex:
            from_field = message.get("from", "")
            if not self.from_regex.search(from_field):
                return False

        # To filter
        if self.to_regex:
            to_field = message.get("to", "")
            if not self.to_regex.search(to_field):
                return False

        # Subject filter
        if self.subject_regex:
            subject = message.get("subject", "")
            if not self.subject_regex.search(subject):
                return False

        # Body contains filter
        if self.body_contains:
            body = get_body(message, encoding).lower()
            if self.body_contains not in body:
                return False

        # Attachment filter
        if self.has_attachment is not None:
            has_attach = any(
                part.get_filename() for part in message.walk()
            ) if message.is_multipart() else False
            if has_attach != self.has_attachment:
                return False

        return True

    def matches_dict(self, email_dict: Dict[str, Any], encoding: str = "utf-8") -> bool:
        """Check if email dictionary matches all filter criteria."""
        # Date filters
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

        # From filter
        if self.from_regex:
            from_field = email_dict.get("from", "") or email_dict.get("From", "")
            if not self.from_regex.search(from_field):
                return False

        # To filter
        if self.to_regex:
            to_field = email_dict.get("to", "") or email_dict.get("To", "")
            if not self.to_regex.search(to_field):
                return False

        # Subject filter
        if self.subject_regex:
            subject = email_dict.get("subject", "") or email_dict.get("Subject", "")
            if not self.subject_regex.search(subject):
                return False

        # Body contains filter
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
    )


# =============================================================================
# FORMAT READERS - Read emails from various formats
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


def read_from_csv(
    path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> Iterator[Tuple[int, Any]]:
    """Read emails from CSV file and yield EmailMessage objects."""
    with open(path, "r", encoding=encoding, newline="") as fh:
        reader = csv.DictReader(fh)
        rows = list(reader)
    
    if show_progress and TQDM_AVAILABLE and not quiet:
        rows = tqdm(rows, desc=f"Reading {Path(path).name}", unit="emails", ncols=80)
    
    filtered_count = 0
    for idx, row in enumerate(rows):
        # Normalize column names (case-insensitive)
        normalized = {k.lower(): v for k, v in row.items()}
        
        if email_filter is not None and not email_filter.matches_dict(normalized, encoding):
            continue
        
        # Create EmailMessage from CSV row
        msg = create_email_message(
            from_addr=normalized.get("from", ""),
            to_addr=normalized.get("to", ""),
            cc_addr=normalized.get("cc", ""),
            bcc_addr=normalized.get("bcc", ""),
            subject=normalized.get("subject", ""),
            date_str=normalized.get("date", ""),
            body=normalized.get("body", ""),
            encoding=encoding,
        )
        filtered_count += 1
        yield idx, msg
    
    logger.info(f"Read {len(rows) if isinstance(rows, list) else 'N'} emails from CSV, {filtered_count} matched filters")


def read_from_eml_directory(
    path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> Iterator[Tuple[int, Any]]:
    """Read emails from directory of EML files."""
    if os.path.isfile(path) and path.endswith(".eml"):
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
    
    logger.info(f"Read {len(eml_files) if isinstance(eml_files, list) else 'N'} EML files, {filtered_count} matched filters")


def read_from_txt(
    path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> Iterator[Tuple[int, Any]]:
    """Read emails from TXT file (our format with delimiters)."""
    with open(path, "r", encoding=encoding) as fh:
        content = fh.read()
    
    # Split by our delimiter pattern
    email_blocks = re.split(r"={80}\nEmail #\d+\n-{40}\n", content)
    email_blocks = [b.strip() for b in email_blocks if b.strip()]
    
    if show_progress and TQDM_AVAILABLE and not quiet:
        email_blocks = tqdm(email_blocks, desc=f"Reading {Path(path).name}", unit="emails", ncols=80)
    
    filtered_count = 0
    for idx, block in enumerate(email_blocks):
        try:
            # Parse header section
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
    
    logger.info(f"Read {len(email_blocks) if isinstance(email_blocks, list) else 'N'} emails from TXT, {filtered_count} matched filters")


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
# FORMAT WRITERS - Write emails to various formats
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


def write_to_csv(
    emails: Iterator[Tuple[int, Any]],
    output_path: str,
    encoding: str = "utf-8",
) -> int:
    """Write emails to CSV file."""
    count = 0
    
    with open(output_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(["Index", "From", "To", "Cc", "Bcc", "Subject", "Date", "Body", "HasAttachment"])
        
        for idx, msg in emails:
            has_attach = any(part.get_filename() for part in msg.walk()) if msg.is_multipart() else False
            writer.writerow([
                idx,
                msg.get("from", ""),
                msg.get("to", ""),
                msg.get("cc", ""),
                msg.get("bcc", ""),
                msg.get("subject", ""),
                msg.get("date", ""),
                get_body(msg, encoding).replace("\r", "").replace("\n", " "),
                has_attach,
            ])
            count += 1
    
    return count


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
        item.Save()
        count += 1
    
    outlook.RemoveStore(pst_root)
    return count


# =============================================================================
# FORMAT DETECTION
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
    elif path_lower.endswith(".txt"):
        return "txt"
    elif path_lower.endswith(".pst"):
        return "pst"
    elif os.path.isdir(path):
        # Check if directory contains EML files
        if glob.glob(os.path.join(path, "*.eml")):
            return "eml"
    
    # Try to detect by content
    if os.path.isfile(path):
        with open(path, "rb") as fh:
            header = fh.read(100)
        
        if header.startswith(b"From "):
            return "mbox"
        elif b"Index,From,To" in header or b"from,to,subject" in header.lower():
            return "csv"
    
    return "unknown"


def get_reader(fmt: str):
    """Get the appropriate reader function for a format."""
    readers = {
        "mbox": read_from_mbox,
        "csv": read_from_csv,
        "eml": read_from_eml_directory,
        "txt": read_from_txt,
        "pst": read_from_pst,
    }
    return readers.get(fmt)


def get_writer(fmt: str):
    """Get the appropriate writer function for a format."""
    writers = {
        "mbox": write_to_mbox,
        "csv": write_to_csv,
        "eml": write_to_eml,
        "txt": write_to_txt,
        "pst": write_to_pst,
    }
    return writers.get(fmt)


# =============================================================================
# UNIVERSAL CONVERT FUNCTION
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
    """Universal conversion function between any supported formats."""
    
    # Auto-detect input format
    if input_format is None:
        input_format = detect_format(input_path)
        if input_format == "unknown":
            raise ValueError(f"Cannot detect input format for: {input_path}")
    
    # Auto-detect output format
    if output_format is None:
        output_format = detect_format(output_path)
        if output_format == "unknown":
            # Default to extension-based detection
            ext = Path(output_path).suffix.lower().lstrip(".")
            if ext in SUPPORTED_FORMATS:
                output_format = ext
            else:
                raise ValueError(f"Cannot detect output format for: {output_path}")
    
    logger.info(f"Converting {input_format.upper()} -> {output_format.upper()}")
    
    # Get reader and writer
    reader = get_reader(input_format)
    writer = get_writer(output_format)
    
    if reader is None:
        raise ValueError(f"Unsupported input format: {input_format}")
    if writer is None:
        raise ValueError(f"Unsupported output format: {output_format}")
    
    # For EML output, output_path is a directory
    if output_format == "eml":
        os.makedirs(output_path, exist_ok=True)
    
    # Read and write
    emails = reader(input_path, email_filter, encoding, show_progress, quiet)
    
    # We need to handle the iterator properly for writers
    # Some writers need to iterate multiple times, so we'll collect if needed
    if output_format in ["mbox", "pst"]:
        # These need the iterator directly
        count = writer(emails, output_path, encoding)
    else:
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
            os.path.getsize(os.path.join(path, f))
            for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))
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


# =============================================================================
# PATH EXPANSION
# =============================================================================

def expand_paths(patterns: List[str]) -> List[str]:
    """Expand glob patterns and directories to list of files."""
    paths = []
    
    for pattern in patterns:
        if "*" in pattern or "?" in pattern:
            expanded = glob.glob(pattern, recursive=True)
            paths.extend(expanded)
        elif os.path.isdir(pattern):
            for ext in [".mbox", ".csv", ".txt", ".pst"]:
                paths.extend(glob.glob(os.path.join(pattern, f"*{ext}")))
            # Also check for EML directories
            eml_files = glob.glob(os.path.join(pattern, "*.eml"))
            if eml_files:
                paths.append(pattern)  # Add directory for EML processing
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
        logger.error("No input files found matching the pattern(s)")
        return EXIT_ERROR
    
    logger.info(f"Found {len(input_paths)} file(s) to process")
    
    email_filter = build_filter_from_args(args)
    if email_filter:
        logger.info("Email filters applied")
    
    # Dry run mode
    if args.dry_run:
        print("\n=== DRY RUN MODE ===")
        print(f"Would process {len(input_paths)} file(s):")
        for f in input_paths:
            info = get_file_info(f, args.encoding)
            print(f"  - {f} [{info['format'].upper()}]: {info['total_emails']} emails ({info['file_size_mb']} MB)")
        print(f"\nOutput format: {args.format}")
        print(f"Output directory: {args.output_dir or 'current directory'}")
        return EXIT_SUCCESS
    
    # Process files
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
            
            # Determine output path
            if args.format == "eml":
                output_path = os.path.join(output_dir, f"{base_name}_eml")
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
            logger.error(f"Error processing {input_path}: {e}")
            errors += 1
            results.append({
                "input": input_path,
                "output": None,
                "emails_converted": 0,
                "status": "error",
                "error": str(e),
            })
    
    # Summary
    if not args.quiet:
        print(f"\n=== CONVERSION COMPLETE ===")
        print(f"Files processed: {len(input_paths)}")
        print(f"Total emails converted: {total_converted}")
        if errors:
            print(f"Errors: {errors}")
    
    # Write report
    if args.report:
        with open(args.report, "w", encoding="utf-8") as fh:
            json.dump({
                "summary": {
                    "files_processed": len(input_paths),
                    "total_converted": total_converted,
                    "errors": errors,
                },
                "results": results,
            }, fh, indent=2)
        logger.info(f"Report written to {args.report}")
    
    if errors == len(input_paths):
        return EXIT_ERROR
    elif errors > 0:
        return EXIT_PARTIAL_SUCCESS
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
                print(f"  Format: {info['format'].upper()}")
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
    """Handle list subcommand - list emails."""
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
                logger.warning(f"Unsupported format: {path}")
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
                    print(f"[{idx}] {msg.get('date', 'No date')} | {msg.get('from', 'Unknown')} | {msg.get('subject', 'No subject')}")
                
                if count >= args.limit:
                    break
            
            if not args.json and count > args.limit:
                print(f"... showing {args.limit} of {count} emails (use --limit to show more)")
        
        except Exception as e:
            logger.error(f"Error reading {path}: {e}")
    
    if args.json:
        print(json.dumps(all_emails[:args.limit], indent=2))
    
    return EXIT_SUCCESS


def cmd_formats(args) -> int:
    """Show supported formats and conversion matrix."""
    print("""
=== MBOX Converter - Supported Formats ===

FORMAT   EXTENSION   READ   WRITE   DESCRIPTION
------   ---------   ----   -----   -----------
MBOX     .mbox       ✓      ✓       Mozilla Thunderbird mailbox
CSV      .csv        ✓      ✓       Comma-separated values (spreadsheet)
EML      .eml/*      ✓      ✓       Individual email files (RFC 822)
TXT      .txt        ✓      ✓       Plain text (human-readable)
PST      .pst        ✓      ✓       Outlook Personal Storage (Windows only)

=== Conversion Matrix ===

All formats can convert to any other format:

  FROM ↓  →  TO →   MBOX   CSV    EML    TXT    PST
  ─────────────────────────────────────────────────
  MBOX              -      ✓      ✓      ✓      ✓*
  CSV               ✓      -      ✓      ✓      ✓*
  EML               ✓      ✓      -      ✓      ✓*
  TXT               ✓      ✓      ✓      -      ✓*
  PST*              ✓      ✓      ✓      ✓      -

* PST requires Windows with Microsoft Outlook installed

=== Examples ===

# MBOX to CSV
mbox_converter convert inbox.mbox --format csv

# CSV to MBOX
mbox_converter convert emails.csv --format mbox

# EML directory to CSV
mbox_converter convert ./eml_folder/ --format csv

# CSV to EML files
mbox_converter convert emails.csv --format eml --output-dir ./eml_output

# TXT to MBOX
mbox_converter convert emails.txt --format mbox
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
            },
            "output": {
                "format": "csv",
                "directory": "./output",
            }
        }
        
        with open(args.generate, "w", encoding="utf-8") as fh:
            json.dump(sample, fh, indent=2)
        print(f"Sample configuration written to {args.generate}")
        return EXIT_SUCCESS
    
    print("Use --generate <path> to create a sample config file")
    return EXIT_SUCCESS


# =============================================================================
# MAIN & ARGUMENT PARSING
# =============================================================================

def main() -> int:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        prog="mbox_converter",
        description="Universal Email Format Converter - Convert between MBOX, CSV, EML, TXT, PST",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert MBOX to CSV
  %(prog)s convert inbox.mbox --format csv

  # Convert CSV to MBOX
  %(prog)s convert emails.csv --format mbox

  # Convert EML files to CSV
  %(prog)s convert ./emails/*.eml --format csv

  # Convert with filtering
  %(prog)s convert inbox.mbox --format csv --date-after 2023-01-01 --from-pattern "@company.com"

  # Batch convert all formats
  %(prog)s convert ./archive/* --format csv --output-dir ./converted

  # Show supported formats
  %(prog)s formats

  # Get file info
  %(prog)s info inbox.mbox emails.csv

Exit Codes:
  0 - Success
  1 - Error
  2 - Invalid arguments
  3 - Partial success (some files failed)
"""
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    
    subparsers = parser.add_subparsers(dest="command", help="Available commands")
    
    # === CONVERT SUBCOMMAND ===
    convert_parser = subparsers.add_parser("convert", help="Convert between email formats")
    convert_parser.add_argument(
        "input", nargs="+",
        help="Input file(s), glob patterns, or directories"
    )
    convert_parser.add_argument(
        "--format", "-f", choices=SUPPORTED_FORMATS, default="csv",
        help="Output format (default: csv)"
    )
    convert_parser.add_argument("--output-dir", "-o", help="Output directory")
    convert_parser.add_argument("--encoding", "-e", default="utf-8", help="Email encoding (default: utf-8)")
    convert_parser.add_argument("--dry-run", action="store_true", help="Preview without writing")
    convert_parser.add_argument("--progress", "-p", action="store_true", help="Show progress bar")
    convert_parser.add_argument("--quiet", "-q", action="store_true", help="Suppress output")
    convert_parser.add_argument("--verbose", "-v", action="count", default=0, help="Increase verbosity")
    convert_parser.add_argument("--log-file", help="Write logs to file")
    convert_parser.add_argument("--report", help="Write JSON report")
    
    # Filtering options
    filter_group = convert_parser.add_argument_group("Filtering Options")
    filter_group.add_argument("--date-after", help="Only include emails after date (YYYY-MM-DD)")
    filter_group.add_argument("--date-before", help="Only include emails before date (YYYY-MM-DD)")
    filter_group.add_argument("--from-pattern", help="Filter by sender (regex)")
    filter_group.add_argument("--to-pattern", help="Filter by recipient (regex)")
    filter_group.add_argument("--subject-pattern", help="Filter by subject (regex)")
    filter_group.add_argument("--body-contains", help="Filter by body content")
    filter_group.add_argument(
        "--has-attachment",
        type=lambda x: x.lower() in ("true", "1", "yes"),
        nargs="?", const=True,
        help="Filter by attachment"
    )
    
    # === INFO SUBCOMMAND ===
    info_parser = subparsers.add_parser("info", help="Show file information")
    info_parser.add_argument("input", nargs="+", help="Input file(s)")
    info_parser.add_argument("--encoding", "-e", default="utf-8", help="Encoding")
    info_parser.add_argument("--json", action="store_true", help="Output as JSON")
    info_parser.add_argument("--quiet", "-q", action="store_true")
    info_parser.add_argument("--verbose", "-v", action="count", default=0)
    
    # === LIST SUBCOMMAND ===
    list_parser = subparsers.add_parser("list", help="List emails")
    list_parser.add_argument("input", nargs="+", help="Input file(s)")
    list_parser.add_argument("--limit", "-n", type=int, default=100, help="Max emails to list")
    list_parser.add_argument("--encoding", "-e", default="utf-8", help="Encoding")
    list_parser.add_argument("--json", action="store_true", help="Output as JSON")
    list_parser.add_argument("--quiet", "-q", action="store_true")
    list_parser.add_argument("--verbose", "-v", action="count", default=0)
    
    # List filters
    list_filter = list_parser.add_argument_group("Filtering")
    list_filter.add_argument("--date-after", help="Filter by date")
    list_filter.add_argument("--date-before", help="Filter by date")
    list_filter.add_argument("--from-pattern", help="Filter by sender")
    list_filter.add_argument("--to-pattern", help="Filter by recipient")
    list_filter.add_argument("--subject-pattern", help="Filter by subject")
    list_filter.add_argument("--body-contains", help="Filter by body")
    list_filter.add_argument("--has-attachment", type=lambda x: x.lower() in ("true", "1", "yes"), nargs="?", const=True)
    
    # === FORMATS SUBCOMMAND ===
    subparsers.add_parser("formats", help="Show supported formats and conversion matrix")
    
    # === CONFIG SUBCOMMAND ===
    config_parser = subparsers.add_parser("config", help="Configuration utilities")
    config_parser.add_argument("--generate", metavar="PATH", help="Generate sample config")
    
    # === LEGACY MODE ===
    if len(sys.argv) > 1 and sys.argv[1] not in ["convert", "info", "list", "config", "formats", "-h", "--help", "--version"]:
        legacy_parser = argparse.ArgumentParser(description="MBOX Converter (Legacy Mode)")
        legacy_parser.add_argument("input", help="Input file")
        legacy_parser.add_argument("--csv", help="Output CSV file")
        legacy_parser.add_argument("--pst", help="Output PST file")
        legacy_parser.add_argument("--eml", help="Output EML directory")
        legacy_parser.add_argument("--txt", help="Output TXT file")
        legacy_parser.add_argument("--mbox", help="Output MBOX file")
        
        legacy_args = legacy_parser.parse_args()
        
        outputs = [
            ("csv", legacy_args.csv),
            ("pst", legacy_args.pst),
            ("eml", legacy_args.eml),
            ("txt", legacy_args.txt),
            ("mbox", legacy_args.mbox),
        ]
        outputs = [(fmt, path) for fmt, path in outputs if path]
        
        if not outputs:
            legacy_parser.error("Specify at least one output: --csv, --pst, --eml, --txt, or --mbox")
        
        setup_logging(0, None, False)
        
        try:
            for fmt, output_path in outputs:
                count = convert(
                    input_path=legacy_args.input,
                    output_path=output_path,
                    output_format=fmt,
                )
                print(f"{fmt.upper()} written to {output_path} ({count} emails)")
            return EXIT_SUCCESS
        except Exception as e:
            logger.error(str(e))
            return EXIT_ERROR
    
    # Parse arguments
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return EXIT_INVALID_ARGS
    
    # Dispatch
    if args.command == "convert":
        return cmd_convert(args)
    elif args.command == "info":
        return cmd_info(args)
    elif args.command == "list":
        return cmd_list(args)
    elif args.command == "formats":
        return cmd_formats(args)
    elif args.command == "config":
        return cmd_config(args)
    
    return EXIT_SUCCESS


if __name__ == "__main__":
    sys.exit(main())
