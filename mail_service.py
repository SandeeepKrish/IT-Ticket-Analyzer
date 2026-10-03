"""
Mail Service Module
-------------------
Handles email directory parsing, creator email matching,
reminder message templating, mailto link generation, and SMTP dispatch.
"""

from __future__ import annotations

import smtplib
import urllib.parse
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Any, Dict, List, Optional, Tuple
import os

import pandas as pd


def get_smtp_env_config() -> Dict[str, Any]:
    """Retrieve SMTP & Sender configuration across Streamlit secrets and environment variables."""
    cfg = {
        "smtp_host": "smtp.office365.com",
        "smtp_port": 587,
        "smtp_password": "",
        "use_tls": True,
        "sender_email": "ttkerp@outlook.com",
        "sender_name": "Sandeep Yadav",
        "sender_company": "United Tyrekrafts Pvt. Ltd.",
        "sender_phone": "9682548514",
    }
    
    # 1. Check Streamlit Cloud Secrets (st.secrets)
    try:
        import streamlit as st
        try:
            sec = st.secrets
            if sec is not None:
                if "SMTP_HOST" in sec and sec["SMTP_HOST"]:
                    cfg["smtp_host"] = str(sec["SMTP_HOST"]).strip()
                if "SMTP_PORT" in sec and sec["SMTP_PORT"]:
                    cfg["smtp_port"] = int(sec["SMTP_PORT"])
                if "SMTP_PASSWORD" in sec and sec["SMTP_PASSWORD"]:
                    cfg["smtp_password"] = str(sec["SMTP_PASSWORD"]).strip()
                if "SMTP_USE_TLS" in sec and sec["SMTP_USE_TLS"] is not None:
                    cfg["use_tls"] = bool(sec["SMTP_USE_TLS"])
                if "SENDER_EMAIL" in sec and sec["SENDER_EMAIL"]:
                    cfg["sender_email"] = str(sec["SENDER_EMAIL"]).strip()
                if "SENDER_NAME" in sec and sec["SENDER_NAME"]:
                    cfg["sender_name"] = str(sec["SENDER_NAME"]).strip()
                if "COMPANY_NAME" in sec and sec["COMPANY_NAME"]:
                    cfg["sender_company"] = str(sec["COMPANY_NAME"]).strip()
                if "SENDER_PHONE" in sec and sec["SENDER_PHONE"]:
                    cfg["sender_phone"] = str(sec["SENDER_PHONE"]).strip()
        except Exception:
            pass
    except Exception:
        pass

    # 2. Check os.environ
    if os.environ.get("SMTP_HOST"):
        cfg["smtp_host"] = os.environ.get("SMTP_HOST", "").strip()
    if os.environ.get("SMTP_PORT"):
        try:
            cfg["smtp_port"] = int(os.environ.get("SMTP_PORT", "587").strip())
        except Exception:
            pass
    if os.environ.get("SMTP_PASSWORD"):
        cfg["smtp_password"] = os.environ.get("SMTP_PASSWORD", "").strip()
    if os.environ.get("SMTP_USE_TLS") is not None:
        cfg["use_tls"] = os.environ.get("SMTP_USE_TLS", "").strip().lower() in ("true", "1", "yes")
    if os.environ.get("SENDER_EMAIL"):
        cfg["sender_email"] = os.environ.get("SENDER_EMAIL", "").strip()
    if os.environ.get("SENDER_NAME"):
        cfg["sender_name"] = os.environ.get("SENDER_NAME", "").strip()
    if os.environ.get("COMPANY_NAME"):
        cfg["sender_company"] = os.environ.get("COMPANY_NAME", "").strip()
    if os.environ.get("SENDER_PHONE"):
        cfg["sender_phone"] = os.environ.get("SENDER_PHONE", "").strip()
    
    return cfg


def build_email_directory(mail_to_df: Optional[pd.DataFrame]) -> Dict[str, str]:
    """
    Extract a mapping from person/ticket creator name -> email address.
    Supports flexible column names:
    - Creator/Name: 'Ticket Creator', 'Name', 'User', 'Employee Name', etc.
    - Email: 'Email', 'Email ID', 'Mail', 'Mail ID', 'Mail To', etc.
    """
    if mail_to_df is None or mail_to_df.empty:
        return {}
    
    col_names = list(mail_to_df.columns)
    name_col = None
    email_col = None
    
    # 1. Search for name column
    name_patterns = ["ticket creator", "creator", "employee name", "user name", "name", "user", "person", "staff", "employee"]
    for pattern in name_patterns:
        for c in col_names:
            if pattern in str(c).strip().lower():
                name_col = c
                break
        if name_col:
            break
            
    # 2. Search for email column
    email_patterns = ["email id", "mail id", "mail to", "email address", "email", "mail", "e-mail"]
    for pattern in email_patterns:
        for c in col_names:
            if pattern in str(c).strip().lower():
                email_col = c
                break
        if email_col:
            break
            
    # Fallback heuristics
    if not name_col and len(col_names) > 0:
        name_col = col_names[0]
    if not email_col and len(col_names) > 1:
        for c in col_names:
            sample_vals = mail_to_df[c].dropna().astype(str).tolist()[:10]
            if any('@' in v for v in sample_vals):
                email_col = c
                break
        if not email_col:
            email_col = col_names[1]

    if not name_col or not email_col:
        return {}

    directory: Dict[str, str] = {}
    for _, row in mail_to_df.iterrows():
        raw_name = str(row.get(name_col, "")).strip()
        raw_email = str(row.get(email_col, "")).strip()
        if raw_name and raw_email and "@" in raw_email:
            directory[raw_name.lower()] = raw_email
            directory[raw_name] = raw_email
            
    return directory


def find_creator_email(creator_name: str, directory: Dict[str, str]) -> Optional[str]:
    """Find email for a creator name using exact, lowercase, or partial token matching."""
    if not creator_name or not directory:
        return None
        
    c_clean = creator_name.strip()
    if c_clean in directory:
        return directory[c_clean]
    if c_clean.lower() in directory:
        return directory[c_clean.lower()]
        
    # Check partial matching (e.g., 'Sanket Tambe' matching 'Sanket' or 'sanket.tambe')
    c_lower = c_clean.lower()
    for key, email in directory.items():
        if key and len(key) >= 3:
            if key in c_lower or c_lower in key:
                return email
                
    # Token check
    tokens = [t for t in c_lower.replace('.', ' ').split() if len(t) >= 3]
    for key, email in directory.items():
        if any(tok in key for tok in tokens):
            return email
            
    return None


def generate_reminder_email(
    ticket_row: pd.Series, 
    sender_name: str = "Sandeep Yadav", 
    company_name: str = "United Tyrekrafts Pvt. Ltd.", 
    sender_email: str = "sandeep.yadav@unitread.co.in",
    sender_phone: str = "9682548514",
    *args: Any,
    **kwargs: Any
) -> Tuple[str, str]:
    """Generate subject and professional body for reminder email to ticket creator."""
    phone = kwargs.get("sender_phone", sender_phone)
    if args and len(args) > 0 and isinstance(args[0], str):
        phone = args[0]

    ticket_no = str(ticket_row.get("Ticket Number", "")).strip()
    subject_text = str(ticket_row.get("Subject", "")).strip()
    creator_name = str(ticket_row.get("Ticket Creator", "User")).strip()
    aging = ticket_row.get("Ticket Aging", 0)
    dept = ticket_row.get("Ticket Department", "Support")
    start_date = ticket_row.get("Start Date", "N/A")

    email_subject = f"Follow-up: Ticket #{ticket_no} - {subject_text}"
    
    email_body = (
        f"Dear {creator_name},\n\n"
        f"Hi, this is {sender_name} from {company_name}.\n\n"
        f"We are following up on your support ticket, which is currently marked 'Awaiting User Info':\n\n"
        f"  • Ticket Number: {ticket_no}\n"
        f"  • Subject: {subject_text}\n"
        f"  • Department: {dept}\n"
        f"  • Start Date: {start_date}\n"
        f"  • Aging: {aging} days\n\n"
        f"Please reply to your ticket as soon as possible as we have to clear awaiting user info, so our support team can take the next steps:\n"
        f"  1. If your problem has been RESOLVED, please confirm so we can close this ticket.\n"
        f"  2. If the problem is NOT SOLVED, please let us know your pending questions/issues so we can solve it for you immediately.\n\n"
        f"Please reply to this on ticket as this has to be cleared. If any doubt, contact us on {phone}.\n\n"
        f"Thank you,\n"
        f"{sender_name}\n"
        f"{company_name}\n"
        f"Email: {sender_email}\n"
    )
    return email_subject, email_body


def generate_consolidated_creator_reminder(
    creator_name: str,
    ticket_items: List[Any],
    sender_name: str = "Sandeep Yadav",
    company_name: str = "United Tyrekrafts Pvt. Ltd.",
    sender_email: str = "sandeep.yadav@unitread.co.in",
    sender_phone: str = "9682548514",
    *args: Any,
    **kwargs: Any
) -> Tuple[str, str]:
    """
    Generate subject and consolidated body for an email to a ticket creator 
    who has one or multiple tickets in 'Awaiting User Info'.
    """
    phone = kwargs.get("sender_phone", sender_phone)
    if args and len(args) > 0 and isinstance(args[0], str):
        phone = args[0]

    c_clean = creator_name.strip() if creator_name else "User"
    ticket_count = len(ticket_items)

    if ticket_count <= 1:
        if ticket_items:
            first_item = ticket_items[0]
            t_num = str(first_item.get("Ticket Number", "") if hasattr(first_item, "get") else getattr(first_item, "ticket_no", "")).strip()
            t_subj = str(first_item.get("Subject", "") if hasattr(first_item, "get") else getattr(first_item, "subject", "")).strip()
            email_subject = f"Follow-up: Ticket #{t_num} - {t_subj}"
        else:
            email_subject = f"Follow-up: Support Ticket Information - {company_name}"
    else:
        email_subject = f"Follow-up: Pending Information on Your {ticket_count} Support Tickets - {company_name}"

    ticket_lines = []
    for idx, item in enumerate(ticket_items, 1):
        t_no = str(item.get("Ticket Number", "") if hasattr(item, "get") else getattr(item, "ticket_no", "")).strip()
        subj = str(item.get("Subject", "") if hasattr(item, "get") else getattr(item, "subject", "")).strip()
        dept = str(item.get("Ticket Department", "Support") if hasattr(item, "get") else getattr(item, "department", "Support")).strip()
        start = str(item.get("Start Date", "N/A") if hasattr(item, "get") else getattr(item, "start_date", "N/A")).strip()
        aging = item.get("Ticket Aging", 0) if hasattr(item, "get") else getattr(item, "aging", 0)

        ticket_lines.append(
            f"  {idx}. Ticket #{t_no}\n"
            f"     • Subject: {subj}\n"
            f"     • Department: {dept}\n"
            f"     • Start Date: {start}\n"
            f"     • Aging: {aging} days"
        )

    tickets_block = "\n\n".join(ticket_lines)

    if ticket_count > 1:
        header_intro = (
            f"We are following up on your support tickets, which are currently marked 'Awaiting User Info'. "
            f"You currently have {ticket_count} pending tickets waiting for your feedback:\n\n"
            f"{tickets_block}\n\n"
            f"Please reply to your tickets as soon as possible as we have to clear awaiting user info, so our support team can take the next steps:\n"
            f"  1. If the problem has been RESOLVED, please confirm the ticket number(s) so we can close them.\n"
            f"  2. If the problem is NOT SOLVED, please let us know your pending questions/issues so we can solve it for you immediately."
        )
    else:
        header_intro = (
            f"We are following up on your support ticket, which is currently marked 'Awaiting User Info':\n\n"
            f"{tickets_block}\n\n"
            f"Please reply to your ticket as soon as possible as we have to clear awaiting user info, so our support team can take the next steps:\n"
            f"  1. If your problem has been RESOLVED, please confirm so we can close this ticket.\n"
            f"  2. If the problem is NOT SOLVED, please let us know your pending questions/issues so we can solve it for you immediately."
        )

    email_body = (
        f"Dear {c_clean},\n\n"
        f"Hi, this is {sender_name} from {company_name}.\n\n"
        f"{header_intro}\n\n"
        f"Please reply to this on ticket as this has to be cleared. If any doubt, contact us on {phone}.\n\n"
        f"Thank you,\n"
        f"{sender_name}\n"
        f"{company_name}\n"
        f"Email: {sender_email}\n"
    )

    return email_subject, email_body


def generate_mailto_url(recipient_email: str, subject: str, body: str) -> str:
    """Generate an encoded mailto: URL for one-click email client dispatch."""
    if not recipient_email:
        return ""
    encoded_sub = urllib.parse.quote(subject)
    encoded_body = urllib.parse.quote(body)
    return f"mailto:{recipient_email}?subject={encoded_sub}&body={encoded_body}"


def send_smtp_email(
    smtp_host: str, 
    smtp_port: int, 
    sender_email: str, 
    sender_password: str, 
    recipient_email: str, 
    subject: str, 
    body: str,
    use_tls: bool = True
) -> Tuple[bool, str]:
    """Send an email using SMTP."""
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        if smtp_port == 465:
            with smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=15) as server:
                server.login(sender_email, sender_password)
                server.send_message(msg)
        else:
            with smtplib.SMTP(smtp_host, smtp_port, timeout=15) as server:
                if use_tls:
                    server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(msg)
        return True, "Sent successfully"
    except Exception as e:
        return False, str(e)
