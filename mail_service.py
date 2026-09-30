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

import pandas as pd


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
        
    c_clean = str(creator_name).strip()
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
    sender_email: str = "utkerp@outlook.com"
) -> Tuple[str, str]:
    """Generate subject and professional body for reminder email to ticket creator."""
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
        f"Please reply to your ticket so our support team can take the next steps:\n"
        f"  1. If your problem has been RESOLVED, please confirm so we can close this ticket.\n"
        f"  2. If the problem is NOT SOLVED, please let us know your pending questions/issues so we can solve it for you immediately.\n\n"
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
