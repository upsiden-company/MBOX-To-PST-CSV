#!/usr/bin/env python3
"""Simple MBOX to CSV/PST converter."""

import argparse
import csv
import mailbox
import os
import sys


def get_body(message):
    """Extract plain text body from an email message."""
    if message.is_multipart():
        parts = []
        for part in message.walk():
            if part.get_content_type() == "text/plain" and not part.get_filename():
                data = part.get_payload(decode=True)
                if data:
                    charset = part.get_content_charset() or "utf-8"
                    parts.append(data.decode(charset, errors="ignore"))
        return "\n".join(parts)
    payload = message.get_payload(decode=True)
    if isinstance(payload, bytes):
        return payload.decode(message.get_content_charset() or "utf-8", errors="ignore")
    return payload


def convert_to_csv(mbox_path: str, csv_path: str) -> None:
    """Convert an MBOX file to CSV."""
    mbox = mailbox.mbox(mbox_path)
    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(["From", "To", "Cc", "Bcc", "Subject", "Date", "Body"])
        for msg in mbox:
            writer.writerow([
                msg.get("from", ""),
                msg.get("to", ""),
                msg.get("cc", ""),
                msg.get("bcc", ""),
                msg.get("subject", ""),
                msg.get("date", ""),
                get_body(msg).replace("\r", "").replace("\n", " "),
            ])
    mbox.close()


def convert_to_pst(mbox_path: str, pst_path: str) -> None:
    """Convert MBOX to PST using Outlook COM (Windows only)."""
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError("pywin32 is required for PST conversion") from exc

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
    mbox = mailbox.mbox(mbox_path)
    for msg in mbox:
        item = target.Items.Add()
        item.Subject = msg.get("subject", "")
        item.To = msg.get("to", "")
        item.CC = msg.get("cc", "")
        item.BCC = msg.get("bcc", "")
        item.Sender = msg.get("from", "")
        item.Body = get_body(msg)
        item.Save()
    mbox.close()
    outlook.RemoveStore(pst_root)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert MBOX to CSV or PST")
    parser.add_argument("mbox", help="Path to the input mbox file")
    parser.add_argument("--csv", help="Output CSV file")
    parser.add_argument("--pst", help="Output PST file (Windows + Outlook)")
    args = parser.parse_args()

    if not args.csv and not args.pst:
        parser.error("Specify --csv or --pst output")

    if args.csv:
        convert_to_csv(args.mbox, args.csv)
        print(f"CSV written to {args.csv}")

    if args.pst:
        convert_to_pst(args.mbox, args.pst)
        print(f"PST written to {args.pst}")
