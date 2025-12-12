#!/usr/bin/env python3
"""
MBOX Converter - A comprehensive CLI tool for converting MBOX files.

Supports conversion to CSV, PST (Windows+Outlook), EML, and TXT formats.
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
from email.generator import Generator
from email.utils import parsedate_to_datetime
from io import StringIO
from pathlib import Path
from typing import Any, Callable, Dict, Iterator, List, Optional, Tuple

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
__version__ = "2.0.0"

# Exit codes
EXIT_SUCCESS = 0
EXIT_ERROR = 1
EXIT_INVALID_ARGS = 2
EXIT_PARTIAL_SUCCESS = 3

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
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
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
                # Normalize both dates to naive (remove timezone) for comparison
                msg_date_naive = msg_date.replace(tzinfo=None) if msg_date.tzinfo else msg_date
                date_after_naive = self.date_after.replace(tzinfo=None) if self.date_after and self.date_after.tzinfo else self.date_after
                date_before_naive = self.date_before.replace(tzinfo=None) if self.date_before and self.date_before.tzinfo else self.date_before
                
                if date_after_naive and msg_date_naive < date_after_naive:
                    return False
                if date_before_naive and msg_date_naive > date_before_naive:
                    return False
            elif self.date_after or self.date_before:
                # No date and we have date filters - skip
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
                if message.is_multipart()
            )
            if has_attach != self.has_attachment:
                return False

        return True


def count_messages(mbox_path: str) -> int:
    """Count messages in an MBOX file."""
    mbox = mailbox.mbox(mbox_path)
    count = len(mbox)
    mbox.close()
    return count


def iter_filtered_messages(
    mbox_path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> Iterator[Tuple[int, Any]]:
    """Iterate over filtered messages with optional progress bar."""
    mbox = mailbox.mbox(mbox_path)
    total = len(mbox)
    
    iterator = enumerate(mbox)
    
    if show_progress and TQDM_AVAILABLE and not quiet:
        iterator = tqdm(
            list(iterator),
            desc=f"Processing {Path(mbox_path).name}",
            unit="emails",
            ncols=80,
        )
    
    filtered_count = 0
    for idx, msg in iterator:
        if email_filter is None or email_filter.matches(msg, encoding):
            filtered_count += 1
            yield idx, msg
    
    mbox.close()
    logger.info(f"Processed {total} emails, {filtered_count} matched filters")


def convert_to_csv(
    mbox_path: str,
    csv_path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> int:
    """Convert an MBOX file to CSV. Returns count of converted emails."""
    count = 0
    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(["Index", "From", "To", "Cc", "Bcc", "Subject", "Date", "Body", "HasAttachment"])
        
        for idx, msg in iter_filtered_messages(mbox_path, email_filter, encoding, show_progress, quiet):
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


def convert_to_eml(
    mbox_path: str,
    output_dir: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> int:
    """Convert MBOX to individual EML files. Returns count of converted emails."""
    os.makedirs(output_dir, exist_ok=True)
    count = 0
    
    for idx, msg in iter_filtered_messages(mbox_path, email_filter, encoding, show_progress, quiet):
        # Create filename from subject or index
        subject = msg.get("subject", "no_subject") or "no_subject"
        # Sanitize filename
        safe_subject = re.sub(r'[<>:"/\\|?*]', '_', subject)[:50]
        filename = f"{idx:06d}_{safe_subject}.eml"
        filepath = os.path.join(output_dir, filename)
        
        with open(filepath, "wb") as fh:
            fh.write(msg.as_bytes())
        count += 1
    
    return count


def convert_to_txt(
    mbox_path: str,
    txt_path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> int:
    """Convert MBOX to plain text file. Returns count of converted emails."""
    count = 0
    
    with open(txt_path, "w", encoding="utf-8") as fh:
        for idx, msg in iter_filtered_messages(mbox_path, email_filter, encoding, show_progress, quiet):
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


def convert_to_pst(
    mbox_path: str,
    pst_path: str,
    email_filter: Optional[EmailFilter] = None,
    encoding: str = "utf-8",
    show_progress: bool = False,
    quiet: bool = False,
) -> int:
    """Convert MBOX to PST using Outlook COM (Windows only). Returns count of converted emails."""
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for PST conversion. "
            "Install with: pip install pywin32\n"
            "Note: PST conversion only works on Windows with Outlook installed."
        ) from exc

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    outlook.AddStoreEx(os.path.abspath(pst_path), 3)  # 3 = olStoreUnicode
    
    pst_root = None
    for folder in outlook.Folders:
        try:
            if hasattr(folder, "Store") and getattr(folder.Store, "FilePath", None):
                if os.path.abspath(folder.Store.FilePath) == os.path.abspath(pst_path):
                    pst_root = folder
                    break
        except Exception:
            continue

    if pst_root is None:
        raise RuntimeError("Failed to create PST store")

    target = pst_root.Folders.Add("Imported")
    count = 0
    
    for idx, msg in iter_filtered_messages(mbox_path, email_filter, encoding, show_progress, quiet):
        item = target.Items.Add()
        item.Subject = msg.get("subject", "")
        item.To = msg.get("to", "")
        item.CC = msg.get("cc", "")
        item.BCC = msg.get("bcc", "")
        item.Sender = msg.get("from", "")
        item.Body = get_body(msg, encoding)
        item.Save()
        count += 1
    
    outlook.RemoveStore(pst_root)
    return count


def get_mbox_info(mbox_path: str, encoding: str = "utf-8") -> Dict[str, Any]:
    """Get information about an MBOX file."""
    mbox = mailbox.mbox(mbox_path)
    total = len(mbox)
    
    dates = []
    senders = set()
    subjects_with_attachments = 0
    
    for msg in mbox:
        date = parse_date(msg.get("date"))
        if date:
            dates.append(date)
        
        from_addr = msg.get("from", "")
        if from_addr:
            senders.add(from_addr)
        
        if msg.is_multipart() and any(part.get_filename() for part in msg.walk()):
            subjects_with_attachments += 1
    
    mbox.close()
    
    info = {
        "path": mbox_path,
        "total_emails": total,
        "unique_senders": len(senders),
        "emails_with_attachments": subjects_with_attachments,
        "file_size_bytes": os.path.getsize(mbox_path),
        "file_size_mb": round(os.path.getsize(mbox_path) / (1024 * 1024), 2),
    }
    
    if dates:
        info["date_range"] = {
            "earliest": min(dates).isoformat(),
            "latest": max(dates).isoformat(),
        }
    
    return info


def expand_paths(patterns: List[str]) -> List[str]:
    """Expand glob patterns and directories to list of MBOX files."""
    paths = []
    
    for pattern in patterns:
        # Check if it's a glob pattern
        if "*" in pattern or "?" in pattern:
            expanded = glob.glob(pattern, recursive=True)
            paths.extend(expanded)
        elif os.path.isdir(pattern):
            # Scan directory for .mbox files
            for root, _, files in os.walk(pattern):
                for f in files:
                    if f.endswith(".mbox"):
                        paths.append(os.path.join(root, f))
        elif os.path.isfile(pattern):
            paths.append(pattern)
        else:
            logger.warning(f"Path not found: {pattern}")
    
    return sorted(set(paths))


def load_config(config_path: str) -> Dict[str, Any]:
    """Load configuration from JSON file."""
    with open(config_path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def create_sample_config(output_path: str) -> None:
    """Create a sample configuration file."""
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
    
    with open(output_path, "w", encoding="utf-8") as fh:
        json.dump(sample, fh, indent=2)


def build_filter_from_args(args) -> Optional[EmailFilter]:
    """Build EmailFilter from command line arguments."""
    has_filters = any([
        args.date_after,
        args.date_before,
        args.from_pattern,
        args.to_pattern,
        args.subject_pattern,
        args.body_contains,
        args.has_attachment is not None,
    ])
    
    if not has_filters:
        return None
    
    date_after = parse_filter_date(args.date_after) if args.date_after else None
    date_before = parse_filter_date(args.date_before) if args.date_before else None
    
    return EmailFilter(
        date_after=date_after,
        date_before=date_before,
        from_pattern=args.from_pattern,
        to_pattern=args.to_pattern,
        subject_pattern=args.subject_pattern,
        body_contains=args.body_contains,
        has_attachment=args.has_attachment,
    )


def cmd_convert(args) -> int:
    """Handle convert subcommand."""
    setup_logging(args.verbose, args.log_file, args.quiet)
    
    # Expand input paths
    mbox_files = expand_paths(args.input)
    
    if not mbox_files:
        logger.error("No MBOX files found matching the input pattern(s)")
        return EXIT_ERROR
    
    logger.info(f"Found {len(mbox_files)} MBOX file(s) to process")
    
    # Build filter
    email_filter = build_filter_from_args(args)
    if email_filter:
        logger.info("Email filters applied")
    
    # Dry run mode
    if args.dry_run:
        print("\n=== DRY RUN MODE ===")
        print(f"Would process {len(mbox_files)} file(s):")
        for f in mbox_files:
            info = get_mbox_info(f, args.encoding)
            print(f"  - {f}: {info['total_emails']} emails ({info['file_size_mb']} MB)")
        print(f"\nOutput format: {args.format}")
        print(f"Output directory: {args.output_dir or 'current directory'}")
        return EXIT_SUCCESS
    
    # Process files
    results = []
    total_converted = 0
    errors = 0
    
    for mbox_path in mbox_files:
        logger.info(f"Processing: {mbox_path}")
        
        try:
            # Determine output path
            base_name = Path(mbox_path).stem
            output_dir = args.output_dir or os.path.dirname(mbox_path) or "."
            os.makedirs(output_dir, exist_ok=True)
            
            if args.format == "csv":
                output_path = os.path.join(output_dir, f"{base_name}.csv")
                count = convert_to_csv(
                    mbox_path, output_path, email_filter,
                    args.encoding, args.progress, args.quiet
                )
            elif args.format == "eml":
                output_path = os.path.join(output_dir, f"{base_name}_eml")
                count = convert_to_eml(
                    mbox_path, output_path, email_filter,
                    args.encoding, args.progress, args.quiet
                )
            elif args.format == "txt":
                output_path = os.path.join(output_dir, f"{base_name}.txt")
                count = convert_to_txt(
                    mbox_path, output_path, email_filter,
                    args.encoding, args.progress, args.quiet
                )
            elif args.format == "pst":
                output_path = os.path.join(output_dir, f"{base_name}.pst")
                count = convert_to_pst(
                    mbox_path, output_path, email_filter,
                    args.encoding, args.progress, args.quiet
                )
            else:
                logger.error(f"Unknown format: {args.format}")
                errors += 1
                continue
            
            total_converted += count
            results.append({
                "input": mbox_path,
                "output": output_path,
                "emails_converted": count,
                "status": "success",
            })
            
            if not args.quiet:
                print(f"✓ {mbox_path} -> {output_path} ({count} emails)")
        
        except Exception as e:
            logger.error(f"Error processing {mbox_path}: {e}")
            errors += 1
            results.append({
                "input": mbox_path,
                "output": None,
                "emails_converted": 0,
                "status": "error",
                "error": str(e),
            })
    
    # Summary
    if not args.quiet:
        print(f"\n=== CONVERSION COMPLETE ===")
        print(f"Files processed: {len(mbox_files)}")
        print(f"Total emails converted: {total_converted}")
        if errors:
            print(f"Errors: {errors}")
    
    # Write results to JSON if requested
    if args.report:
        with open(args.report, "w", encoding="utf-8") as fh:
            json.dump({
                "summary": {
                    "files_processed": len(mbox_files),
                    "total_converted": total_converted,
                    "errors": errors,
                },
                "results": results,
            }, fh, indent=2)
        logger.info(f"Report written to {args.report}")
    
    if errors == len(mbox_files):
        return EXIT_ERROR
    elif errors > 0:
        return EXIT_PARTIAL_SUCCESS
    return EXIT_SUCCESS


def cmd_info(args) -> int:
    """Handle info subcommand."""
    setup_logging(args.verbose, None, args.quiet)
    
    mbox_files = expand_paths(args.input)
    
    if not mbox_files:
        logger.error("No MBOX files found")
        return EXIT_ERROR
    
    all_info = []
    
    for mbox_path in mbox_files:
        try:
            info = get_mbox_info(mbox_path, args.encoding)
            all_info.append(info)
            
            if not args.json:
                print(f"\n=== {mbox_path} ===")
                print(f"  Total emails: {info['total_emails']}")
                print(f"  Unique senders: {info['unique_senders']}")
                print(f"  With attachments: {info['emails_with_attachments']}")
                print(f"  File size: {info['file_size_mb']} MB")
                if "date_range" in info:
                    print(f"  Date range: {info['date_range']['earliest']} to {info['date_range']['latest']}")
        
        except Exception as e:
            logger.error(f"Error reading {mbox_path}: {e}")
    
    if args.json:
        print(json.dumps(all_info, indent=2))
    
    return EXIT_SUCCESS


def cmd_list(args) -> int:
    """Handle list subcommand - list emails in MBOX."""
    setup_logging(args.verbose, None, args.quiet)
    
    mbox_files = expand_paths(args.input)
    
    if not mbox_files:
        logger.error("No MBOX files found")
        return EXIT_ERROR
    
    email_filter = build_filter_from_args(args)
    
    all_emails = []
    
    for mbox_path in mbox_files:
        try:
            for idx, msg in iter_filtered_messages(mbox_path, email_filter, args.encoding, False, args.quiet):
                email_info = {
                    "file": mbox_path,
                    "index": idx,
                    "from": msg.get("from", ""),
                    "to": msg.get("to", ""),
                    "subject": msg.get("subject", ""),
                    "date": msg.get("date", ""),
                }
                all_emails.append(email_info)
                
                if not args.json and len(all_emails) <= args.limit:
                    print(f"[{idx}] {msg.get('date', 'No date')} | {msg.get('from', 'Unknown')} | {msg.get('subject', 'No subject')}")
            
            if not args.json and len(all_emails) > args.limit:
                print(f"... and {len(all_emails) - args.limit} more (use --limit to show more)")
        
        except Exception as e:
            logger.error(f"Error reading {mbox_path}: {e}")
    
    if args.json:
        print(json.dumps(all_emails[:args.limit], indent=2))
    
    return EXIT_SUCCESS


def cmd_config(args) -> int:
    """Handle config subcommand."""
    if args.generate:
        create_sample_config(args.generate)
        print(f"Sample configuration written to {args.generate}")
        return EXIT_SUCCESS
    
    print("Use --generate <path> to create a sample config file")
    return EXIT_SUCCESS


def main() -> int:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        prog="mbox_converter",
        description="MBOX Converter - Convert MBOX files to CSV, PST, EML, or TXT",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert single file to CSV
  %(prog)s convert inbox.mbox --format csv

  # Batch convert all .mbox files in directory
  %(prog)s convert ./emails/*.mbox --format eml --output-dir ./converted

  # Filter by date and sender
  %(prog)s convert inbox.mbox --format csv --date-after 2023-01-01 --from-pattern "@company.com"

  # Preview conversion (dry run)
  %(prog)s convert inbox.mbox --format csv --dry-run

  # Get info about MBOX file(s)
  %(prog)s info inbox.mbox

  # List emails matching filters
  %(prog)s list inbox.mbox --subject-pattern "invoice" --limit 50

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
    convert_parser = subparsers.add_parser("convert", help="Convert MBOX files")
    convert_parser.add_argument(
        "input",
        nargs="+",
        help="Input MBOX file(s), glob patterns (*.mbox), or directories"
    )
    convert_parser.add_argument(
        "--format", "-f",
        choices=["csv", "pst", "eml", "txt"],
        default="csv",
        help="Output format (default: csv)"
    )
    convert_parser.add_argument(
        "--output-dir", "-o",
        help="Output directory (default: same as input)"
    )
    convert_parser.add_argument(
        "--encoding", "-e",
        default="utf-8",
        help="Email encoding (default: utf-8)"
    )
    convert_parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview conversion without writing files"
    )
    convert_parser.add_argument(
        "--progress", "-p",
        action="store_true",
        help="Show progress bar"
    )
    convert_parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Suppress output (errors only)"
    )
    convert_parser.add_argument(
        "--verbose", "-v",
        action="count",
        default=0,
        help="Increase verbosity (-v for info, -vv for debug)"
    )
    convert_parser.add_argument(
        "--log-file",
        help="Write logs to file"
    )
    convert_parser.add_argument(
        "--report",
        help="Write conversion report to JSON file"
    )
    
    # Filtering options
    filter_group = convert_parser.add_argument_group("Filtering Options")
    filter_group.add_argument(
        "--date-after",
        help="Only include emails after this date (YYYY-MM-DD)"
    )
    filter_group.add_argument(
        "--date-before",
        help="Only include emails before this date (YYYY-MM-DD)"
    )
    filter_group.add_argument(
        "--from-pattern",
        help="Filter by sender (regex pattern)"
    )
    filter_group.add_argument(
        "--to-pattern",
        help="Filter by recipient (regex pattern)"
    )
    filter_group.add_argument(
        "--subject-pattern",
        help="Filter by subject (regex pattern)"
    )
    filter_group.add_argument(
        "--body-contains",
        help="Filter by body content (substring match)"
    )
    filter_group.add_argument(
        "--has-attachment",
        type=lambda x: x.lower() in ("true", "1", "yes"),
        nargs="?",
        const=True,
        help="Filter by attachment presence (true/false)"
    )
    
    # === INFO SUBCOMMAND ===
    info_parser = subparsers.add_parser("info", help="Show MBOX file information")
    info_parser.add_argument(
        "input",
        nargs="+",
        help="Input MBOX file(s) or patterns"
    )
    info_parser.add_argument(
        "--encoding", "-e",
        default="utf-8",
        help="Email encoding (default: utf-8)"
    )
    info_parser.add_argument(
        "--json",
        action="store_true",
        help="Output as JSON"
    )
    info_parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Suppress warnings"
    )
    info_parser.add_argument(
        "--verbose", "-v",
        action="count",
        default=0,
        help="Increase verbosity"
    )
    
    # === LIST SUBCOMMAND ===
    list_parser = subparsers.add_parser("list", help="List emails in MBOX")
    list_parser.add_argument(
        "input",
        nargs="+",
        help="Input MBOX file(s)"
    )
    list_parser.add_argument(
        "--limit", "-n",
        type=int,
        default=100,
        help="Maximum emails to list (default: 100)"
    )
    list_parser.add_argument(
        "--encoding", "-e",
        default="utf-8",
        help="Email encoding (default: utf-8)"
    )
    list_parser.add_argument(
        "--json",
        action="store_true",
        help="Output as JSON"
    )
    list_parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Suppress warnings"
    )
    list_parser.add_argument(
        "--verbose", "-v",
        action="count",
        default=0,
        help="Increase verbosity"
    )
    
    # List filtering options
    list_filter_group = list_parser.add_argument_group("Filtering Options")
    list_filter_group.add_argument("--date-after", help="Filter by date (YYYY-MM-DD)")
    list_filter_group.add_argument("--date-before", help="Filter by date (YYYY-MM-DD)")
    list_filter_group.add_argument("--from-pattern", help="Filter by sender (regex)")
    list_filter_group.add_argument("--to-pattern", help="Filter by recipient (regex)")
    list_filter_group.add_argument("--subject-pattern", help="Filter by subject (regex)")
    list_filter_group.add_argument("--body-contains", help="Filter by body content")
    list_filter_group.add_argument(
        "--has-attachment",
        type=lambda x: x.lower() in ("true", "1", "yes"),
        nargs="?",
        const=True,
        help="Filter by attachment"
    )
    
    # === CONFIG SUBCOMMAND ===
    config_parser = subparsers.add_parser("config", help="Configuration utilities")
    config_parser.add_argument(
        "--generate",
        metavar="PATH",
        help="Generate sample config file"
    )
    
    # === LEGACY MODE (backward compatibility) ===
    # Support old-style: mbox_converter.py input.mbox --csv output.csv
    if len(sys.argv) > 1 and not sys.argv[1] in ["convert", "info", "list", "config", "-h", "--help", "--version"]:
        # Legacy mode
        legacy_parser = argparse.ArgumentParser(description="MBOX Converter (Legacy Mode)")
        legacy_parser.add_argument("mbox", help="Input MBOX file")
        legacy_parser.add_argument("--csv", help="Output CSV file")
        legacy_parser.add_argument("--pst", help="Output PST file")
        legacy_parser.add_argument("--eml", help="Output EML directory")
        legacy_parser.add_argument("--txt", help="Output TXT file")
        
        legacy_args = legacy_parser.parse_args()
        
        if not any([legacy_args.csv, legacy_args.pst, legacy_args.eml, legacy_args.txt]):
            legacy_parser.error("Specify at least one output: --csv, --pst, --eml, or --txt")
        
        setup_logging(0, None, False)
        
        try:
            if legacy_args.csv:
                count = convert_to_csv(legacy_args.mbox, legacy_args.csv)
                print(f"CSV written to {legacy_args.csv} ({count} emails)")
            
            if legacy_args.pst:
                count = convert_to_pst(legacy_args.mbox, legacy_args.pst)
                print(f"PST written to {legacy_args.pst} ({count} emails)")
            
            if legacy_args.eml:
                count = convert_to_eml(legacy_args.mbox, legacy_args.eml)
                print(f"EML files written to {legacy_args.eml}/ ({count} emails)")
            
            if legacy_args.txt:
                count = convert_to_txt(legacy_args.mbox, legacy_args.txt)
                print(f"TXT written to {legacy_args.txt} ({count} emails)")
            
            return EXIT_SUCCESS
        
        except Exception as e:
            logger.error(str(e))
            return EXIT_ERROR
    
    # Parse arguments
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return EXIT_INVALID_ARGS
    
    # Dispatch to subcommand
    if args.command == "convert":
        return cmd_convert(args)
    elif args.command == "info":
        return cmd_info(args)
    elif args.command == "list":
        return cmd_list(args)
    elif args.command == "config":
        return cmd_config(args)
    
    return EXIT_SUCCESS


if __name__ == "__main__":
    sys.exit(main())
