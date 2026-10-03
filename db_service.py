"""
SQLite Database and Demo Data Management Service
------------------------------------------------
Provides persistent SQLite storage for ticket datasets and demo files,
enabling instant demo mode without requiring manual Excel file uploads.
"""

from __future__ import annotations

import os
import sqlite3
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "tickets.db")
DEMO_FOLDER_PATH = os.path.join(BASE_DIR, "demo_files")

REQUIRED_COLS = [
    "Ticket Number", "Ticket Creator", "Subject", "Customer Name",
    "Ticket Department", "Status", "Ticket Owner", "Ticket Priority",
    "Ticket Aging", "Start Date", "Required Date",
    "Estimate Resolution Date", "Description",
]

TABLE_KEYS = ["master", "wip", "dev", "wait", "hold", "open", "pending", "closed", "mail_to"]

TABLE_NAME_MAP = {
    "master": "master_tickets",
    "wip": "wip_tickets",
    "dev": "dev_tickets",
    "wait": "wait_tickets",
    "hold": "hold_tickets",
    "open": "open_tickets",
    "pending": "pending_tickets",
    "closed": "closed_tickets",
    "mail_to": "mail_directory",
}

EXCEL_FILE_NAMES = {
    "master": "1_Master_Ticket_List.xlsx",
    "wip": "2_Work_In_Progress.xlsx",
    "dev": "3_Under_Development.xlsx",
    "wait": "4_Awaiting_User_Info.xlsx",
    "hold": "5_Hold_Tickets.xlsx",
    "open": "6_Open_Tickets.xlsx",
    "pending": "7_Pending_Tickets.xlsx",
    "closed": "8_Closed_Tickets.xlsx",
    "mail_to": "9_Mail_Directory_Mapping.xlsx",
}


def get_connection(db_path: str = DB_PATH) -> sqlite3.Connection:
    """Ensure directory exists and return an SQLite connection."""
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    # Ensure sent_reminders log table always exists
    conn.execute("""
        CREATE TABLE IF NOT EXISTS sent_reminders (
            ticket_number TEXT PRIMARY KEY,
            creator_name TEXT,
            recipient_email TEXT,
            subject TEXT,
            sent_by TEXT,
            sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            delivery_mode TEXT DEFAULT 'Outlook'
        )
    """)
    conn.commit()
    return conn


def record_sent_reminder(
    ticket_number: str,
    creator_name: str,
    recipient_email: str,
    subject: str,
    sent_by: str = "utkerp@outlook.com",
    delivery_mode: str = "Outlook",
    db_path: str = DB_PATH
) -> bool:
    """Record or update a ticket email reminder dispatch in SQLite."""
    if not ticket_number:
        return False
    try:
        conn = get_connection(db_path)
        with conn:
            conn.execute("""
                INSERT OR REPLACE INTO sent_reminders (
                    ticket_number, creator_name, recipient_email, subject, sent_by, sent_at, delivery_mode
                ) VALUES (?, ?, ?, ?, ?, datetime('now', 'localtime'), ?)
            """, (str(ticket_number).strip(), creator_name, recipient_email, subject, sent_by, delivery_mode))
        conn.close()
        return True
    except Exception:
        return False


def get_sent_reminders_map(db_path: str = DB_PATH) -> Dict[str, Dict[str, Any]]:
    """Return dictionary of ticket_number -> reminder record for quick lookup."""
    try:
        conn = get_connection(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT ticket_number, creator_name, recipient_email, subject, sent_by, sent_at, delivery_mode FROM sent_reminders")
        rows = cursor.fetchall()
        conn.close()
        return {
            str(r[0]).strip(): {
                "ticket_number": str(r[0]).strip(),
                "creator_name": r[1],
                "recipient_email": r[2],
                "subject": r[3],
                "sent_by": r[4],
                "sent_at": r[5],
                "delivery_mode": r[6]
            }
            for r in rows
        }
    except Exception:
        return {}


def get_sent_reminders_dataframe(db_path: str = DB_PATH) -> pd.DataFrame:
    """Return sent reminders history as a clean pandas DataFrame."""
    try:
        conn = get_connection(db_path)
        df = pd.read_sql_query(
            "SELECT ticket_number AS 'Ticket Number', creator_name AS 'Creator Name', "
            "recipient_email AS 'Recipient Email', subject AS 'Email Subject', "
            "sent_by AS 'Sent By', sent_at AS 'Sent At', delivery_mode AS 'Method' "
            "FROM sent_reminders ORDER BY sent_at DESC", 
            conn
        )
        conn.close()
        return df
    except Exception:
        return pd.DataFrame()


def remove_sent_reminder(ticket_number: str, db_path: str = DB_PATH) -> bool:
    """Remove a reminder record from SQLite (marks ticket as Not Sent)."""
    if not ticket_number:
        return False
    try:
        conn = get_connection(db_path)
        with conn:
            conn.execute("DELETE FROM sent_reminders WHERE ticket_number = ?", (str(ticket_number).strip(),))
        conn.close()
        return True
    except Exception:
        return False



def generate_seed_data() -> Dict[str, pd.DataFrame]:
    """
    Generate realistic IT/ERP enterprise ticket datasets for demonstration.
    Reflects the company's real workflows (ERP Support, IT Infrastructure, QAD/SAP,
    United Tyrekrafts Pvt. Ltd. domain context).
    """
    # Master list of realistic tickets
    tickets_data = [
        # --- WIP Tickets ---
        {
            "Ticket Number": "IT-2026-0101",
            "Ticket Creator": "Sanket Tambe",
            "Subject": "QAD ERP Session crash during Month-End closing report generation",
            "Customer Name": "United Tyrekrafts Plant 1",
            "Ticket Department": "ERP Support",
            "Status": "Work in Progress",
            "Ticket Owner": "Sandeep Yadav",
            "Ticket Priority": "Critical",
            "Ticket Aging": 4,
            "Start Date": "2026-09-24",
            "Required Date": "2026-09-26",
            "Estimate Resolution Date": "2026-09-30",
            "Description": "Finance department reporting unhandled memory exception when exporting batch GL trial balances. Investigating database lock on table gl_hist.",
        },
        {
            "Ticket Number": "IT-2026-0104",
            "Ticket Creator": "Rohan Desai",
            "Subject": "VPN Gateway authentication timeout for remote plant managers",
            "Customer Name": "Corporate IT",
            "Ticket Department": "IT Infrastructure",
            "Status": "Work in Progress",
            "Ticket Owner": "Rajesh Kumar",
            "Ticket Priority": "High",
            "Ticket Aging": 6,
            "Start Date": "2026-09-22",
            "Required Date": "2026-09-25",
            "Estimate Resolution Date": "2026-09-29",
            "Description": "Radius server connection drops during peak afternoon hours. Checking packet loss on primary ISP link.",
        },
        {
            "Ticket Number": "IT-2026-0108",
            "Ticket Creator": "Pooja Nair",
            "Subject": "Barcode scanner connectivity failure on Tyre Curing Line 3",
            "Customer Name": "Plant Operations Pune",
            "Ticket Department": "Plant IT Operations",
            "Status": "Work in Progress",
            "Ticket Owner": "Sneha Patel",
            "Ticket Priority": "High",
            "Ticket Aging": 8,
            "Start Date": "2026-09-20",
            "Required Date": "2026-09-23",
            "Estimate Resolution Date": "2026-09-30",
            "Description": "Wireless hand terminal disconnects from AP-04 every 20 minutes. Re-configuring roaming thresholds.",
        },
        {
            "Ticket Number": "IT-2026-0112",
            "Ticket Creator": "Vikram Malhotra",
            "Subject": "Slow query performance on Dispatch & Logistics daily summary",
            "Customer Name": "Supply Chain Logistics",
            "Ticket Department": "Database Admin",
            "Status": "Work in Progress",
            "Ticket Owner": "Amit Sharma",
            "Ticket Priority": "Medium",
            "Ticket Aging": 11,
            "Start Date": "2026-09-17",
            "Required Date": "2026-09-22",
            "Estimate Resolution Date": "2026-09-30",
            "Description": "Query takes 4.5 minutes to load 15k rows. Index optimization and partition maintenance in progress.",
        },
        {
            "Ticket Number": "IT-2026-0115",
            "Ticket Creator": "Anjali Joshi",
            "Subject": "Procurement Purchase Order approval email notification delay",
            "Customer Name": "Purchasing Dept",
            "Ticket Department": "ERP Support",
            "Status": "Work in Progress",
            "Ticket Owner": "Sandeep Yadav",
            "Ticket Priority": "Medium",
            "Ticket Aging": 14,
            "Start Date": "2026-09-14",
            "Required Date": "2026-09-20",
            "Estimate Resolution Date": "2026-09-30",
            "Description": "SMTP relay queue backlog observed during automated workflow trigger for POs exceeding threshold limit.",
        },
        
        # --- Under Development Tickets ---
        {
            "Ticket Number": "IT-2026-0118",
            "Ticket Creator": "Manish Tiwari",
            "Subject": "Automated GST e-Way Bill API integration enhancement",
            "Customer Name": "Finance & Taxation",
            "Ticket Department": "Development",
            "Status": "Under Development",
            "Ticket Owner": "Amit Sharma",
            "Ticket Priority": "High",
            "Ticket Aging": 16,
            "Start Date": "2026-09-12",
            "Required Date": "2026-09-28",
            "Estimate Resolution Date": "2026-10-05",
            "Description": "Building automated payload validation and error handling for government GST portal schema version 1.04 update.",
        },
        {
            "Ticket Number": "IT-2026-0121",
            "Ticket Creator": "Sanket Tambe",
            "Subject": "Custom QC inspection mobile web form for Tyre Tread testing",
            "Customer Name": "Quality Assurance",
            "Ticket Department": "Development",
            "Status": "Under Development",
            "Ticket Owner": "Priya Verma",
            "Ticket Priority": "Medium",
            "Ticket Aging": 19,
            "Start Date": "2026-09-09",
            "Required Date": "2026-09-25",
            "Estimate Resolution Date": "2026-10-02",
            "Description": "Designing responsive tablet input interface with mandatory photo upload and tolerance validation limits.",
        },
        {
            "Ticket Number": "IT-2026-0125",
            "Ticket Creator": "Rohan Desai",
            "Subject": "Inventory Batch Traceability module for Raw Rubber compounds",
            "Customer Name": "Warehouse & Materials",
            "Ticket Department": "Development",
            "Status": "Under Development",
            "Ticket Owner": "Sandeep Yadav",
            "Ticket Priority": "High",
            "Ticket Aging": 22,
            "Start Date": "2026-09-06",
            "Required Date": "2026-09-26",
            "Estimate Resolution Date": "2026-10-10",
            "Description": "Implementing FIFO lot allocation and expiry tracking backend stored procedures in ERP core.",
        },
        {
            "Ticket Number": "IT-2026-0128",
            "Ticket Creator": "Vikram Malhotra",
            "Subject": "Automated Daily Attendance SMS/Email Alert for shift supervisors",
            "Customer Name": "HR Department",
            "Ticket Department": "Development",
            "Status": "Under Development",
            "Ticket Owner": "Priya Verma",
            "Ticket Priority": "Low",
            "Ticket Aging": 25,
            "Start Date": "2026-09-03",
            "Required Date": "2026-09-20",
            "Estimate Resolution Date": "2026-10-01",
            "Description": "Cron script scheduled to aggregate biometric machine logs and fire morning shift exception summary.",
        },

        # --- Awaiting User Info Tickets ---
        {
            "Ticket Number": "IT-2026-0130",
            "Ticket Creator": "Sanket Tambe",
            "Subject": "Request for SAP Sub-ledger access and Cost Center 402 permissions",
            "Customer Name": "Accounts Dept",
            "Ticket Department": "ERP Support",
            "Status": "Awaiting User Info",
            "Ticket Owner": "Sandeep Yadav",
            "Ticket Priority": "High",
            "Ticket Aging": 12,
            "Start Date": "2026-09-16",
            "Required Date": "2026-09-19",
            "Estimate Resolution Date": "2026-09-28",
            "Description": "Need signed approval form from department head Mr. Joshi before granting modification rights to ledger.",
        },
        {
            "Ticket Number": "IT-2026-0133",
            "Ticket Creator": "Anjali Joshi",
            "Subject": "Vendor master bank details update rejection error",
            "Customer Name": "Accounts Payable",
            "Ticket Department": "ERP Support",
            "Status": "Awaiting User Info",
            "Ticket Owner": "Rajesh Kumar",
            "Ticket Priority": "Medium",
            "Ticket Aging": 15,
            "Start Date": "2026-09-13",
            "Required Date": "2026-09-18",
            "Estimate Resolution Date": "2026-09-27",
            "Description": "Waiting for verified cancelled cheque copy and vendor GST registration certificate from user.",
        },
        {
            "Ticket Number": "IT-2026-0137",
            "Ticket Creator": "Pooja Nair",
            "Subject": "Replacement of damaged weighing scale serial communication cable",
            "Customer Name": "Finished Goods Store",
            "Ticket Department": "IT Infrastructure",
            "Status": "Awaiting User Info",
            "Ticket Owner": "Sneha Patel",
            "Ticket Priority": "Medium",
            "Ticket Aging": 18,
            "Start Date": "2026-09-10",
            "Required Date": "2026-09-15",
            "Estimate Resolution Date": "2026-09-25",
            "Description": "Need exact hardware model and serial pin configuration (RS232 DB9 vs DB25) from store in-charge.",
        },
        {
            "Ticket Number": "IT-2026-0140",
            "Ticket Creator": "Manish Tiwari",
            "Subject": "Discrepancy in Q3 Sales commission rebate percentage calculation",
            "Customer Name": "Commercial Sales",
            "Ticket Department": "ERP Support",
            "Status": "Awaiting User Info",
            "Ticket Owner": "Sandeep Yadav",
            "Ticket Priority": "High",
            "Ticket Aging": 24,
            "Start Date": "2026-09-04",
            "Required Date": "2026-09-12",
            "Estimate Resolution Date": "2026-09-29",
            "Description": "Requested policy amendment circular number and approved tier formula from VP Sales.",
        },

        # --- Hold Tickets ---
        {
            "Ticket Number": "IT-2026-0145",
            "Ticket Creator": "Rohan Desai",
            "Subject": "Server Room UPS Battery replacement and thermal cooling overhaul",
            "Customer Name": "Infrastructure & Plant Maint",
            "Ticket Department": "IT Infrastructure",
            "Status": "Hold",
            "Ticket Owner": "Rajesh Kumar",
            "Ticket Priority": "High",
            "Ticket Aging": 35,
            "Start Date": "2026-08-24",
            "Required Date": "2026-09-10",
            "Estimate Resolution Date": "2026-10-15",
            "Description": "Awaiting capital expenditure budget approval from executive committee for Eaton 40kVA battery banks.",
        },
        {
            "Ticket Number": "IT-2026-0148",
            "Ticket Creator": "Vikram Malhotra",
            "Subject": "License renewal for 3D CAD Mold Design software",
            "Customer Name": "R&D Tyre Design",
            "Ticket Department": "IT Infrastructure",
            "Status": "Hold",
            "Ticket Owner": "Sneha Patel",
            "Ticket Priority": "Medium",
            "Ticket Aging": 40,
            "Start Date": "2026-08-19",
            "Required Date": "2026-09-05",
            "Estimate Resolution Date": "2026-10-20",
            "Description": "Vendor quote review in progress by commercial audit team. Negotiations on concurrent license counts.",
        },

        # --- Open Tickets ---
        {
            "Ticket Number": "IT-2026-0151",
            "Ticket Creator": "Sanket Tambe",
            "Subject": "New user desktop onboarding for Quality Inspector - Plant 2",
            "Customer Name": "Quality Assurance",
            "Ticket Department": "IT Infrastructure",
            "Status": "Open",
            "Ticket Owner": "Sneha Patel",
            "Ticket Priority": "Medium",
            "Ticket Aging": 2,
            "Start Date": "2026-09-27",
            "Required Date": "2026-09-30",
            "Estimate Resolution Date": "2026-10-01",
            "Description": "Standard Windows 11 installation with ERP client, antivirus, and local network printer mapping.",
        },
        {
            "Ticket Number": "IT-2026-0154",
            "Ticket Creator": "Pooja Nair",
            "Subject": "Production work order print misalignment on Line 1 Dot Matrix",
            "Customer Name": "Production Line 1",
            "Ticket Department": "Plant IT Operations",
            "Status": "Open",
            "Ticket Owner": "Rajesh Kumar",
            "Ticket Priority": "High",
            "Ticket Aging": 3,
            "Start Date": "2026-09-26",
            "Required Date": "2026-09-28",
            "Estimate Resolution Date": "2026-09-30",
            "Description": "Top margin offset shifted by 1.5 inches after continuous stationery paper roll change.",
        },
        {
            "Ticket Number": "IT-2026-0158",
            "Ticket Creator": "Manish Tiwari",
            "Subject": "Reset locked active directory account for regional sales executive",
            "Customer Name": "North Zone Sales",
            "Ticket Department": "IT Infrastructure",
            "Status": "Open",
            "Ticket Owner": "Sandeep Yadav",
            "Ticket Priority": "Low",
            "Ticket Aging": 1,
            "Start Date": "2026-09-28",
            "Required Date": "2026-09-29",
            "Estimate Resolution Date": "2026-09-29",
            "Description": "User entered wrong password on mobile Outlook client repeatedly. Account locked out.",
        },

        # --- Pending Tickets ---
        {
            "Ticket Number": "IT-2026-0162",
            "Ticket Creator": "Anjali Joshi",
            "Subject": "Quarterly physical inventory stock audit reconciliation mismatch",
            "Customer Name": "Finance & Audit",
            "Ticket Department": "ERP Support",
            "Status": "Pending",
            "Ticket Owner": "Sandeep Yadav",
            "Ticket Priority": "High",
            "Ticket Aging": 21,
            "Start Date": "2026-09-07",
            "Required Date": "2026-09-18",
            "Estimate Resolution Date": "2026-10-03",
            "Description": "Pending physical count sign-off from external auditors for Unitread Nashik warehouse.",
        },
        {
            "Ticket Number": "IT-2026-0166",
            "Ticket Creator": "Rohan Desai",
            "Subject": "Fiber link splicing approval from highway authority for Plant 3 expansion",
            "Customer Name": "Network Projects",
            "Ticket Department": "IT Infrastructure",
            "Status": "Pending",
            "Ticket Owner": "Rajesh Kumar",
            "Ticket Priority": "Medium",
            "Ticket Aging": 28,
            "Start Date": "2026-08-31",
            "Required Date": "2026-09-15",
            "Estimate Resolution Date": "2026-10-12",
            "Description": "Right of way clearance permission pending with state road development corporation.",
        },

        # --- Closed Tickets ---
        {
            "Ticket Number": "IT-2026-0171",
            "Ticket Creator": "Sanket Tambe",
            "Subject": "Thermal barcode label printer ribbon replacement in Dispatch",
            "Customer Name": "Dispatch Store",
            "Ticket Department": "Plant IT Operations",
            "Status": "Closed",
            "Ticket Owner": "Sneha Patel",
            "Ticket Priority": "Medium",
            "Ticket Aging": 7,
            "Start Date": "2026-09-18",
            "Required Date": "2026-09-21",
            "Estimate Resolution Date": "2026-09-22",
            "Description": "Replaced wax-resin ribbon and cleaned printhead with isopropyl alcohol. Test print verified.",
        },
        {
            "Ticket Number": "IT-2026-0175",
            "Ticket Creator": "Pooja Nair",
            "Subject": "Email forwarding rule configured for retired store supervisor",
            "Customer Name": "Human Resources",
            "Ticket Department": "IT Infrastructure",
            "Status": "Closed",
            "Ticket Owner": "Rajesh Kumar",
            "Ticket Priority": "Low",
            "Ticket Aging": 5,
            "Start Date": "2026-09-21",
            "Required Date": "2026-09-24",
            "Estimate Resolution Date": "2026-09-24",
            "Description": "Configured M365 auto-forwarding to current store in-charge as per HR exit checklist.",
        },
        {
            "Ticket Number": "IT-2026-0179",
            "Ticket Creator": "Vikram Malhotra",
            "Subject": "Monthly database vacuum and index rebuild on production cluster",
            "Customer Name": "Internal IT Ops",
            "Ticket Department": "Database Admin",
            "Status": "Closed",
            "Ticket Owner": "Amit Sharma",
            "Ticket Priority": "High",
            "Ticket Aging": 3,
            "Start Date": "2026-09-24",
            "Required Date": "2026-09-26",
            "Estimate Resolution Date": "2026-09-25",
            "Description": "Completed scheduled maintenance window on Sunday night. Freed 42GB unindexed space.",
        },
        {
            "Ticket Number": "IT-2026-0182",
            "Ticket Creator": "Anjali Joshi",
            "Subject": "Credit memo authorization workflow issue for returned tread rubber",
            "Customer Name": "Customer Billing",
            "Ticket Department": "ERP Support",
            "Status": "Closed",
            "Ticket Owner": "Sandeep Yadav",
            "Ticket Priority": "Critical",
            "Ticket Aging": 4,
            "Start Date": "2026-09-23",
            "Required Date": "2026-09-25",
            "Estimate Resolution Date": "2026-09-26",
            "Description": "Resolved table lock on credit note verification queue. Processed 14 held memos.",
        },
    ]

    master_df = pd.DataFrame(tickets_data)
    for col in REQUIRED_COLS:
        if col not in master_df.columns:
            master_df[col] = "—"

    # Category filters
    wip_df = pd.DataFrame(master_df[master_df["Status"] == "Work in Progress"].copy().reset_index(drop=True))
    dev_df = pd.DataFrame(master_df[master_df["Status"] == "Under Development"].copy().reset_index(drop=True))
    wait_df = pd.DataFrame(master_df[master_df["Status"] == "Awaiting User Info"].copy().reset_index(drop=True))
    hold_df = pd.DataFrame(master_df[master_df["Status"] == "Hold"].copy().reset_index(drop=True))
    open_df = pd.DataFrame(master_df[master_df["Status"] == "Open"].copy().reset_index(drop=True))
    pending_df = pd.DataFrame(master_df[master_df["Status"] == "Pending"].copy().reset_index(drop=True))
    closed_df = pd.DataFrame(master_df[master_df["Status"] == "Closed"].copy().reset_index(drop=True))

    # Mail Directory Mapping
    mail_to_data = [
        {"Ticket Creator": "Sanket Tambe", "Email ID": "sanket.tambe@unitread.co.in", "Department": "Quality & Operations", "Location": "Plant 1 - Pune"},
        {"Ticket Creator": "Rohan Desai", "Email ID": "rohan.desai@unitread.co.in", "Department": "Materials & Stores", "Location": "Warehouse - Nashik"},
        {"Ticket Creator": "Pooja Nair", "Email ID": "pooja.nair@unitread.co.in", "Department": "Plant Operations", "Location": "Plant 2 - Chakan"},
        {"Ticket Creator": "Vikram Malhotra", "Email ID": "vikram.malhotra@unitread.co.in", "Department": "R&D Tyre Design", "Location": "Tech Center - Mumbai"},
        {"Ticket Creator": "Anjali Joshi", "Email ID": "anjali.joshi@unitread.co.in", "Department": "Finance & Taxation", "Location": "Corporate Office"},
        {"Ticket Creator": "Manish Tiwari", "Email ID": "manish.tiwari@unitread.co.in", "Department": "Commercial Sales", "Location": "Regional Office"},
        {"Ticket Creator": "Sandeep Yadav", "Email ID": "sandeep.yadav@unitread.co.in", "Department": "ERP & IT Lead", "Location": "Corporate IT"},
        {"Ticket Creator": "Rajesh Kumar", "Email ID": "rajesh.kumar@unitread.co.in", "Department": "IT Infrastructure", "Location": "Corporate IT"},
        {"Ticket Creator": "Sneha Patel", "Email ID": "sneha.patel@unitread.co.in", "Department": "Systems Support", "Location": "Plant 1 - Pune"},
        {"Ticket Creator": "Amit Sharma", "Email ID": "amit.sharma@unitread.co.in", "Department": "Database Administration", "Location": "Corporate IT"},
        {"Ticket Creator": "Priya Verma", "Email ID": "priya.verma@unitread.co.in", "Department": "Software Engineering", "Location": "Tech Center - Mumbai"},
    ]
    mail_to_df = pd.DataFrame(mail_to_data)

    return {
        "master": master_df,
        "wip": wip_df,
        "dev": dev_df,
        "wait": wait_df,
        "hold": hold_df,
        "open": open_df,
        "pending": pending_df,
        "closed": closed_df,
        "mail_to": mail_to_df,
    }


def init_database(force_reseed: bool = False, db_path: str = DB_PATH) -> bool:
    """
    Initialize SQLite database tables and seed with enterprise demo data if empty.
    Returns True if initialized or seeded.
    """
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = get_connection(db_path)
    try:
        cursor = conn.cursor()
        # Check if master_tickets exists and has rows
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='master_tickets'")
        table_exists = cursor.fetchone() is not None
        
        has_rows = False
        if table_exists:
            cursor.execute("SELECT count(*) FROM master_tickets")
            count = cursor.fetchone()[0]
            has_rows = count > 0

        if not has_rows or force_reseed:
            seed_dict = generate_seed_data()
            save_all_to_sqlite(seed_dict, dataset_name="Default IT-ERP Demo Showcase", db_path=db_path)
            return True
        return False
    finally:
        conn.close()


def save_all_to_sqlite(
    data_dict: Dict[str, pd.DataFrame],
    dataset_name: str = "Uploaded Dataset",
    db_path: str = DB_PATH
) -> Tuple[bool, str]:
    """
    Save all 8 workflow stages + mail directory to SQLite.
    Overwrites previous demo/current tables with clean, indexed tables.
    """
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = get_connection(db_path)
    try:
        for key in TABLE_KEYS:
            df = data_dict.get(key)
            if df is not None and isinstance(df, pd.DataFrame):
                table_name = TABLE_NAME_MAP[key]
                # Ensure clean column names
                clean_df = df.copy()
                clean_df.columns = [str(c).strip() for c in clean_df.columns]
                clean_df.to_sql(table_name, conn, if_exists="replace", index=False)
        
        # Write metadata
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS meta_info (
                key TEXT PRIMARY KEY,
                value TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        total_tickets = len(data_dict.get("master", [])) if data_dict.get("master") is not None else 0
        cursor.execute("INSERT OR REPLACE INTO meta_info (key, value) VALUES ('dataset_name', ?)", (dataset_name,))
        cursor.execute("INSERT OR REPLACE INTO meta_info (key, value) VALUES ('ticket_count', ?)", (str(total_tickets),))
        cursor.execute("INSERT OR REPLACE INTO meta_info (key, value) VALUES ('last_updated', datetime('now'))", ())
        conn.commit()
        return True, f"Successfully saved dataset '{dataset_name}' with {total_tickets} tickets to SQLite!"
    except Exception as e:
        return False, f"Failed saving to SQLite: {str(e)}"
    finally:
        conn.close()


def load_all_from_sqlite(db_path: str = DB_PATH) -> Optional[Dict[str, pd.DataFrame]]:
    """
    Load all 8 workflow stages and mail directory from SQLite into DataFrames.
    Returns None if tables don't exist or are empty.
    """
    if not os.path.exists(db_path):
        init_database(force_reseed=True, db_path=db_path)

    conn = get_connection(db_path)
    try:
        cursor = conn.cursor()
        # Verify master_tickets exists
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='master_tickets'")
        if not cursor.fetchone():
            return None

        result: Dict[str, pd.DataFrame] = {}
        for key in TABLE_KEYS:
            table_name = TABLE_NAME_MAP[key]
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}'")
            if cursor.fetchone():
                df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
                df.columns = [str(c).strip() for c in df.columns]
                if "Ticket Number" in df.columns:
                    df["Ticket Number"] = df["Ticket Number"].astype(str).str.strip()
                if "Ticket Aging" in df.columns:
                    df["Ticket Aging"] = pd.to_numeric(df["Ticket Aging"], errors="coerce").fillna(0).astype(int)
                result[key] = df
            else:
                result[key] = pd.DataFrame()

        # Check that we have the 8 core ticket stages
        core_keys = ["master", "wip", "dev", "wait", "hold", "open", "pending", "closed"]
        if any(result.get(k) is None or result[k].empty for k in core_keys):
            return None

        return result
    except Exception:
        return None
    finally:
        conn.close()


def get_sqlite_stats(db_path: str = DB_PATH) -> Dict[str, Any]:
    """Get metadata and row counts for all tables in SQLite."""
    if not os.path.exists(db_path):
        return {"exists": False, "total_tickets": 0, "tables": {}}

    conn = get_connection(db_path)
    try:
        cursor = conn.cursor()
        stats: Dict[str, Any] = {"exists": True, "tables": {}, "total_tickets": 0, "dataset_name": "Demo Dataset"}
        
        # Check metadata
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='meta_info'")
        if cursor.fetchone():
            cursor.execute("SELECT key, value FROM meta_info")
            for k, v in cursor.fetchall():
                stats[k] = v

        for key, table_name in TABLE_NAME_MAP.items():
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}'")
            if cursor.fetchone():
                cursor.execute(f"SELECT count(*) FROM {table_name}")
                count = cursor.fetchone()[0]
                stats["tables"][key] = count
                if key == "master":
                    stats["total_tickets"] = count
            else:
                stats["tables"][key] = 0

        return stats
    except Exception:
        return {"exists": False, "total_tickets": 0, "tables": {}}
    finally:
        conn.close()


def export_database_to_excel_folder(folder_path: str = DEMO_FOLDER_PATH, db_path: str = DB_PATH) -> Tuple[bool, str]:
    """
    Export all 9 SQLite tables into standalone .xlsx files in the demo folder.
    This creates the folder so users can open it and select files or showcase folder upload.
    """
    data = load_all_from_sqlite(db_path=db_path)
    if not data:
        # initialize first
        init_database(force_reseed=True, db_path=db_path)
        data = load_all_from_sqlite(db_path=db_path)

    if not data:
        return False, "Failed to load database data for export."

    try:
        os.makedirs(folder_path, exist_ok=True)
        for key, filename in EXCEL_FILE_NAMES.items():
            df = data.get(key)
            if df is not None and isinstance(df, pd.DataFrame):
                target_file = os.path.join(folder_path, filename)
                df.to_excel(target_file, index=False, engine="openpyxl")

        return True, f"All 9 demo Excel files successfully created in '{os.path.basename(folder_path)}/' folder!"
    except Exception as e:
        return False, f"Export failed: {str(e)}"


def load_from_demo_folder(folder_path: str = DEMO_FOLDER_PATH) -> Optional[Dict[str, pd.DataFrame]]:
    """
    Load dataframes directly from the demo_files folder on disk.
    Auto-detects files matching the expected names.
    """
    if not os.path.exists(folder_path):
        return None

    files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith((".xlsx", ".xls", ".csv"))]
    if not files:
        return None

    from ai import detect_file_type  # Reuse fuzzy matching

    result: Dict[str, Any] = {
        'master': None, 'wip': None, 'dev': None, 'wait': None,
        'hold': None, 'open': None, 'pending': None, 'closed': None,
        'mail_to': None
    }

    for file_path in files:
        filename = os.path.basename(file_path)
        cat = detect_file_type(filename)
        if cat:
            try:
                if file_path.endswith(".csv"):
                    df = pd.read_csv(file_path)
                else:
                    df = pd.read_excel(file_path, engine="openpyxl")
                df.columns = [str(c).strip() for c in df.columns]
                if "Ticket Number" in df.columns:
                    df["Ticket Number"] = df["Ticket Number"].astype(str).str.strip()
                if "Ticket Aging" in df.columns:
                    df["Ticket Aging"] = pd.to_numeric(df["Ticket Aging"], errors="coerce").fillna(0).astype(int)
                result[cat] = df
            except Exception:
                pass

    core_keys = ["master", "wip", "dev", "wait", "hold", "open", "pending", "closed"]
    if any(result.get(k) is None or result[k].empty for k in core_keys):
        return None

    return result
