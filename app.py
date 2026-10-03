"""
Ticket Flow Tracker
--------------------
Upload your "Under Development" and "Awaiting User Info" ticket exports
(.xlsx) and search a ticket number to see its status and every other open
ticket for the same owner — so nothing sits stuck waiting on a reply.

Run with:  streamlit run app.py
"""

from __future__ import annotations

import os
import sys
from typing import Any, Dict, List, Optional, Set, Tuple

# Suppress harmless Windows asyncio Proactor WinError 10054 when a client tab closes
if sys.platform == "win32":
    from functools import wraps
    try:
        from asyncio.proactor_events import _ProactorBasePipeTransport
        _orig_call_connection_lost = _ProactorBasePipeTransport._call_connection_lost

        @wraps(_orig_call_connection_lost)
        def _silent_call_connection_lost(self, exc):
            try:
                _orig_call_connection_lost(self, exc)
            except (ConnectionResetError, OSError):
                pass

        _ProactorBasePipeTransport._call_connection_lost = _silent_call_connection_lost
    except Exception:
        pass

import pandas as pd

import streamlit as st
from dotenv import load_dotenv

# Ensure .env is explicitly loaded with absolute path before anything else
ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
load_dotenv(dotenv_path=ENV_PATH, override=True)
load_dotenv(override=True)

from ai import (
    draft_reply,
    generate_data_context,
    get_chatbot_response,
    get_folder_upload_summary,
    has_key_configured,
    process_folder_upload,
    validate_folder_upload,
)
from footer import render_footer
from mail_service import (
    build_email_directory,
    find_creator_email,
    generate_mailto_url,
    generate_reminder_email,
    send_smtp_email,
)
from db_service import (
    DB_PATH,
    DEMO_FOLDER_PATH,
    export_database_to_excel_folder,
    get_sent_reminders_dataframe,
    get_sent_reminders_map,
    get_sqlite_stats,
    init_database,
    load_all_from_sqlite,
    load_from_demo_folder,
    record_sent_reminder,
    remove_sent_reminder,
    save_all_to_sqlite,
)


# Auto-initialize SQLite database with demo data if empty
try:
    init_database()
except Exception:
    pass


def load_dataset_into_session(data_dict: Dict[str, Any], source_label: str = "SQLite Database") -> None:
    """Store loaded dataset (all 8 stages + mail_to) into session state."""
    st.session_state["folder_data_loaded"] = True
    st.session_state["data_source_mode"] = source_label
    st.session_state["master_df"] = data_dict.get("master")
    st.session_state["wip_df"] = data_dict.get("wip")
    st.session_state["dev_df"] = data_dict.get("dev")
    st.session_state["wait_df"] = data_dict.get("wait")
    st.session_state["hold_df"] = data_dict.get("hold")
    st.session_state["open_df"] = data_dict.get("open")
    st.session_state["pending_df"] = data_dict.get("pending")
    st.session_state["closed_df"] = data_dict.get("closed")
    st.session_state["mail_to_df"] = data_dict.get("mail_to", None)


def clear_loaded_session_data() -> None:
    """Clear all loaded dataset keys from session state."""
    keys_to_clear = [
        "folder_data_loaded", "data_source_mode", "master_df", "wip_df",
        "dev_df", "wait_df", "hold_df", "open_df", "pending_df", "closed_df", "mail_to_df"
    ]
    for k in keys_to_clear:
        if k in st.session_state:
            del st.session_state[k]


st.set_page_config(
    page_title="Ticket Flow Tracker", 
    page_icon="🎫", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load external CSS stylesheet
def _load_css(css_path: str) -> None:
    """Read a CSS file and inject it into the Streamlit page."""
    with open(css_path, "r", encoding="utf-8") as f:
        css = f.read()
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)

_load_css(os.path.join(os.path.dirname(__file__), "style.css"))


def display_filter_results(df: pd.DataFrame, category_name: str, owner_filter: Any, year_filter: Any, month_filter: Any, category_color: str, creator_filter: Any = "All creators"):
    """Display prominent filter results with ticket counts and breakdowns"""
    
    # Get filtered data
    filtered_df = get_filtered_data(df, owner_filter, year_filter, month_filter, creator_filter)
    total_tickets = len(filtered_df)
    
    # Build filter description
    filter_parts = []
    if owner_filter != "All owners":
        filter_parts.append(f"Owner: {owner_filter}")
    if creator_filter != "All creators":
        filter_parts.append(f"Creator: {creator_filter}")
    if year_filter != "All Years":
        filter_parts.append(f"Year: {year_filter}")
    if month_filter != "All Months":
        filter_parts.append(f"Month: {month_filter}")
    
    filter_description = " | ".join(filter_parts) if filter_parts else "All Tickets"
    
    # Display main results
    st.markdown(f"""
    <div class="filter-results" style="background: linear-gradient(135deg, {category_color} 0%, {category_color}dd 100%);">
        <div class="filter-results-count">{total_tickets:,}</div>
        <div class="filter-results-label">{category_name} Tickets Found</div>
        <div class="filter-results-details">{filter_description}</div>
    </div>
    """, unsafe_allow_html=True)
    
    # Show monthly breakdown if year is selected but not specific month
    if year_filter != "All Years" and month_filter == "All Months" and total_tickets > 0:
        display_monthly_breakdown(filtered_df, year_filter, category_name)
    
    # Show owner breakdown if month/year selected but not owner
    if (year_filter != "All Years" or month_filter != "All Months") and owner_filter == "All owners" and total_tickets > 0:
        display_owner_breakdown(filtered_df, category_name)


def display_monthly_breakdown(df: pd.DataFrame, year: str, category_name: str):
    """Display monthly breakdown for a specific year"""
    if df.empty or "Start Date" not in df.columns:
        return
        
    df_temp = df.copy()
    df_temp["Start Date"] = pd.to_datetime(df_temp["Start Date"], errors='coerce')
    
    # Group by month
    monthly_counts = df_temp.groupby(df_temp["Start Date"].dt.month).size()
    
    month_names = {
        1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
        7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
    }
    
    st.markdown(f"""
    <div class="monthly-breakdown">
        <div class="monthly-breakdown-title">📅 Monthly Breakdown for {year}</div>
        <div class="monthly-breakdown-grid">
    """, unsafe_allow_html=True)
    
    for month_num in range(1, 13):
        month_name = month_names[month_num]
        count = monthly_counts.get(month_num, 0)
        st.markdown(f"""
            <div class="month-item">
                <div class="month-name">{month_name}</div>
                <div class="month-count">{count}</div>
            </div>
        """, unsafe_allow_html=True)
    
    st.markdown("</div></div>", unsafe_allow_html=True)


def display_owner_breakdown(df: pd.DataFrame, category_name: str):
    """Display top owners breakdown for filtered results"""
    if df.empty or "Ticket Owner" not in df.columns:
        return
        
    owner_counts = df["Ticket Owner"].value_counts().head(5)
    
    if len(owner_counts) > 0:
        st.markdown("### 👥 Top Owners in Filtered Results")
        for owner, count in owner_counts.items():
            percentage = (count / len(df)) * 100
            st.markdown(f"**{owner}**: {count} tickets ({percentage:.1f}%)")


def extract_date_filters(df: pd.DataFrame):
    """Extract unique years and months from Start Date column for filtering"""
    if df.empty or "Start Date" not in df.columns:
        return [], []
    
    # Convert Start Date to datetime
    df_temp = df.copy()
    df_temp["Start Date"] = pd.to_datetime(df_temp["Start Date"], errors='coerce')
    
    # Extract years and months
    years = sorted(df_temp["Start Date"].dt.year.dropna().unique())
    months = [
        "All Months", "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]
    
    return years, months


def filter_by_date(df: pd.DataFrame, selected_year, selected_month):
    """Filter dataframe by year and month"""
    if df.empty or "Start Date" not in df.columns:
        return df
    
    df_filtered = df.copy()
    df_filtered["Start Date"] = pd.to_datetime(df_filtered["Start Date"], errors='coerce')
    
    # Filter by year
    if selected_year != "All Years":
        df_filtered = df_filtered[df_filtered["Start Date"].dt.year == selected_year]
    
    # Filter by month
    if selected_month != "All Months":
        month_num = {
            "January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
            "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12
        }[selected_month]
        df_filtered = df_filtered[df_filtered["Start Date"].dt.month == month_num]
    
    return df_filtered


def get_combined_date_filters(master_df, wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df):
    """Get combined years from all dataframes for consistent filtering"""
    all_years = set()
    
    for df in [master_df, wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df]:
        if df is not None and not df.empty and "Start Date" in df.columns:
            df_temp = df.copy()
            df_temp["Start Date"] = pd.to_datetime(df_temp["Start Date"], errors='coerce')
            years = df_temp["Start Date"].dt.year.dropna().unique()
            all_years.update(years)
    
    sorted_years = ["All Years"] + sorted(list(all_years))
    months = [
        "All Months", "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]
    
    return sorted_years, months


def display_user_statistics(wip_df: pd.DataFrame, dev_df: pd.DataFrame, wait_df: pd.DataFrame, 
                           hold_df: pd.DataFrame, open_df: pd.DataFrame, pending_df: pd.DataFrame, 
                           closed_df: pd.DataFrame, selected_owner: Any, year_filter=None, month_filter=None):
    """Display statistics for the selected owner with animated counters"""
    if selected_owner == "All owners":
        return
    
    def _count(df):
        if df is None or df.empty or "Ticket Owner" not in df.columns:
            return 0
        sub_df = df[df["Ticket Owner"] == selected_owner]
        if year_filter and month_filter:
            return len(filter_by_date(sub_df, year_filter, month_filter))
        return len(sub_df)

    wip_count = _count(wip_df)
    dev_count = _count(dev_df)
    wait_count = _count(wait_df)
    hold_count = _count(hold_df)
    open_count = _count(open_df)
    pending_count = _count(pending_df)
    closed_count = _count(closed_df)
    
    total_active = wip_count + dev_count + wait_count + hold_count + open_count + pending_count
    
    # Display date range info
    date_info = ""
    if year_filter and month_filter:
        date_info = f" ({year_filter if year_filter != 'All Years' else 'All Years'} - {month_filter})"
    
    # Display user statistics with animated counters
    st.markdown(f"""
    <div class="filter-section">
        <div class="filter-title">📊 {selected_owner}'s Ticket Statistics{date_info}</div>
        <div class="user-stats-container">
            <div class="user-stat-card" style="--border-color: #007bff; --number-color: #007bff;">
                <div class="stat-number animate-count">{wip_count}</div>
                <div class="stat-label">Work in Progress</div>
            </div>
            <div class="user-stat-card" style="--border-color: #fd7e14; --number-color: #fd7e14;">
                <div class="stat-number animate-count">{dev_count}</div>
                <div class="stat-label">Under Development</div>
            </div>
            <div class="user-stat-card" style="--border-color: #dc3545; --number-color: #dc3545;">
                <div class="stat-number animate-count">{wait_count}</div>
                <div class="stat-label">Awaiting User Info</div>
            </div>
            <div class="user-stat-card" style="--border-color: #6f42c1; --number-color: #6f42c1;">
                <div class="stat-number animate-count">{hold_count}</div>
                <div class="stat-label">Hold</div>
            </div>
        </div>
        <div class="user-stats-container">
            <div class="user-stat-card" style="--border-color: #28a745; --number-color: #28a745;">
                <div class="stat-number animate-count">{open_count}</div>
                <div class="stat-label">Open</div>
            </div>
            <div class="user-stat-card" style="--border-color: #ffc107; --number-color: #856404;">
                <div class="stat-number animate-count">{pending_count}</div>
                <div class="stat-label">Pending</div>
            </div>
            <div class="user-stat-card" style="--border-color: #6c757d; --number-color: #6c757d;">
                <div class="stat-number animate-count">{closed_count}</div>
                <div class="stat-label">Closed</div>
            </div>
            <div class="user-stat-card" style="--border-color: #17a2b8; --number-color: #17a2b8;">
                <div class="stat-number animate-count">{total_active}</div>
                <div class="stat-label">Total Active</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def get_ticket_workflow_status(ticket_no: str, wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, master_df):
    """Determine ticket's current status in the workflow and return detailed information"""
    
    # Check if ticket exists in master list
    master_match = master_df[master_df["Ticket Number"].str.upper() == ticket_no.upper()]
    if master_match.empty:
        return None, "not_in_master", {
            "title": "🚫 Ticket Not Found",
            "description": "Ticket not found in Master list",
            "color": "#dc3545",
            "next_step": "Please check the ticket number and try again"
        }
    
    # Check current status across all files
    status_checks = {
        "closed": closed_df[closed_df["Ticket Number"].str.upper() == ticket_no.upper()],
        "wait": wait_df[wait_df["Ticket Number"].str.upper() == ticket_no.upper()],
        "dev": dev_df[dev_df["Ticket Number"].str.upper() == ticket_no.upper()],
        "wip": wip_df[wip_df["Ticket Number"].str.upper() == ticket_no.upper()],
        "hold": hold_df[hold_df["Ticket Number"].str.upper() == ticket_no.upper()],
        "pending": pending_df[pending_df["Ticket Number"].str.upper() == ticket_no.upper()],
        "open": open_df[open_df["Ticket Number"].str.upper() == ticket_no.upper()]
    }
    
    # Find current status with priority order
    current_status = None
    current_data = None
    
    for status, df_match in status_checks.items():
        if not df_match.empty:
            current_status = status
            current_data = df_match.iloc[0]
            break
    
    # Workflow descriptions
    workflow_info = {
        "closed": {
            "title": "🎉 Closed - Resolved",
            "description": "Ticket has been successfully resolved and closed",
            "color": "#6c757d",
            "next_step": "No further action needed"
        },
        "wait": {
            "title": "⚠️ Awaiting User Info",
            "description": "Solution provided, waiting for your response/confirmation",
            "color": "#dc3545", 
            "next_step": "Customer response required"
        },
        "dev": {
            "title": "🔧 Under Development",
            "description": "New feature/solution is being developed",
            "color": "#fd7e14",
            "next_step": "Will move to 'Awaiting User Info' when ready"
        },
        "wip": {
            "title": "📊 Work in Progress", 
            "description": "Support team is actively working on the ticket",
            "color": "#007bff",
            "next_step": "May move to Development, Hold, or Awaiting User Info"
        },
        "hold": {
            "title": "⏸️ On Hold",
            "description": "Ticket paused - team is researching solution",
            "color": "#6f42c1",
            "next_step": "Will resume when solution approach is found"
        },
        "pending": {
            "title": "🟠 Pending",
            "description": "Change request raised but not yet started by backend team", 
            "color": "#ffc107",
            "next_step": "Waiting for backend team to begin work"
        },
        "open": {
            "title": "🟢 Open",
            "description": "Ticket available for assignment to support team",
            "color": "#28a745",
            "next_step": "Will be assigned and move to Work in Progress"
        }
    }
    
    if current_status:
        return current_data, current_status, workflow_info[current_status]
    else:
        # Ticket in master but not in any status file
        return master_match.iloc[0], "unknown", {
            "title": "❓ Status Unknown",
            "description": "Ticket exists in master list but current status unclear",
            "color": "#6c757d",
            "next_step": "Contact support for status clarification"
        }


# Email directory parsing, matching, and SMTP dispatch functions are located in mail_service.py


def display_workflow_diagram():
    """Display the ticket workflow process"""
    st.markdown("""
    <div class="workflow-container">
        <div class="workflow-title">🔄 Ticket Workflow Process</div>
        <div class="workflow-steps">
            <div class="workflow-step">
                <div class="step-number">1</div>
                <div class="step-content">
                    <div class="step-title">Query Created</div>
                    <div class="step-desc">Customer creates ticket → Added to Master List</div>
                </div>
            </div>
            <div class="workflow-arrow">→</div>
            <div class="workflow-step">
                <div class="step-number">2</div>
                <div class="step-content">
                    <div class="step-title">🟢 Open</div>
                    <div class="step-desc">Available for assignment</div>
                </div>
            </div>
            <div class="workflow-arrow">→</div>
            <div class="workflow-step">
                <div class="step-number">3</div>
                <div class="step-content">
                    <div class="step-title">📊 Work in Progress</div>
                    <div class="step-desc">Assigned to support team</div>
                </div>
            </div>
        </div>
        
        <div class="workflow-branches">
            <div class="branch">
                <div class="branch-arrow">↓</div>
                <div class="branch-step">
                    <div class="step-title">⚠️ Awaiting User Info</div>
                    <div class="step-desc">Solution provided, waiting for response</div>
                </div>
            </div>
            <div class="branch">
                <div class="branch-arrow">↓</div>
                <div class="branch-step">
                    <div class="step-title">🔧 Under Development</div>
                    <div class="step-desc">New feature being built</div>
                </div>
            </div>
            <div class="branch">
                <div class="branch-arrow">↓</div>
                <div class="branch-step">
                    <div class="step-title">⏸️ Hold</div>
                    <div class="step-desc">Researching solution</div>
                </div>
            </div>
            <div class="branch">
                <div class="branch-arrow">↓</div>
                <div class="branch-step">
                    <div class="step-title">🟠 Pending</div>
                    <div class="step-desc">Backend team not started</div>
                </div>
            </div>
        </div>
        
        <div class="workflow-final">
            <div class="final-arrow">↓</div>
            <div class="final-step">
                <div class="step-title">🎉 Closed</div>
                <div class="step-desc">Issue resolved successfully</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def get_filtered_data(df: pd.DataFrame, owner_filter: Any, year_filter=None, month_filter=None, creator_filter: Any = "All creators") -> pd.DataFrame:
    """Filter dataframe by owner, creator, year, and month"""
    filtered_df = df.copy()
    
    # Filter by owner
    if owner_filter != "All owners" and "Ticket Owner" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["Ticket Owner"] == owner_filter]
    
    # Filter by creator
    if creator_filter != "All creators" and "Ticket Creator" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["Ticket Creator"] == creator_filter]
    
    # Filter by date if filters are provided
    if year_filter is not None and month_filter is not None:
        filtered_df = filter_by_date(filtered_df, year_filter, month_filter)
    
    return pd.DataFrame(filtered_df)



REQUIRED_COLS = [
    "Ticket Number", "Ticket Creator", "Subject", "Customer Name",
    "Ticket Department", "Status", "Ticket Owner", "Ticket Priority",
    "Ticket Aging", "Start Date", "Required Date",
    "Estimate Resolution Date", "Description",
]


@st.cache_data(show_spinner=False)
def load_excel(uploaded_file) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    df["Ticket Number"] = df["Ticket Number"].astype(str).str.strip()
    if "Ticket Aging" in df.columns:
        df["Ticket Aging"] = pd.to_numeric(df["Ticket Aging"], errors="coerce").fillna(0).astype(int)
    return df


def validate_columns(df: pd.DataFrame, label: str) -> bool:
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(f"{label} file is missing expected column(s): {', '.join(missing)}")
        return False
    return True


@st.dialog("🎫 Ticket Description & Conversation Thread", width="large")
def show_ticket_description_dialog(
    t_num: str,
    creator: str,
    matched_addr: str,
    t_subj: str,
    aging: int,
    ticket_owner: str,
    customer_name: str,
    start_date_val: str,
    req_date_val: str,
    est_date_val: str,
    priority_val: str,
    dept_val: str,
    ticket_desc: str,
    email_subject: str = "",
    email_body: str = "",
    mailto_url: str = "",
    is_sent: bool = False,
    cfg_sender_email: str = ""
):
    """Opens a centered modal dialog box floating above the main UI showing ticket description & conversation."""
    st.markdown(f"### 🎫 Ticket #{t_num}: {t_subj}")
    
    # Metadata Overview Callout Banner
    st.markdown(f"""
    <div style="background: #f8fafc; border-left: 4px solid #0284c7; border-radius: 8px; padding: 0.85rem 1.1rem; margin: 0.4rem 0 0.9rem 0; border: 1px solid #e2e8f0; border-left-width: 4px;">
        <div style="font-size: 0.92rem; color: #0f172a; margin-bottom: 0.35rem;">
            👨‍💻 <strong>Ticket Owner / Respondent:</strong> <span style="color: #1d4ed8; font-weight: 700;">{ticket_owner}</span> <span style="color: #64748b;">({dept_val})</span>
            &nbsp;&nbsp;•&nbsp;&nbsp;
            👤 <strong>Created By:</strong> <span style="color: #0369a1; font-weight: 600;">{creator}</span> <span style="color: #64748b;">({matched_addr or 'No email'})</span>
        </div>
        <div style="font-size: 0.82rem; color: #475569; display: flex; flex-wrap: wrap; gap: 0.85rem; margin-top: 0.4rem; padding-top: 0.4rem; border-top: 1px dashed #cbd5e1;">
            <span>🏢 Customer: <strong>{customer_name}</strong></span>
            <span>⚡ Priority: <strong>{priority_val}</strong></span>
            <span>📅 Opened: <strong>{start_date_val}</strong></span>
            <span>🎯 Est. Resolution: <strong>{est_date_val}</strong></span>
            <span>⏱️ Aging: <strong>{aging} days</strong></span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Ticket Description & Conversation Thread
    st.markdown("#### 📝 Ticket Description & Conversation Thread")
    if ticket_desc and ticket_desc not in ("nan", "None", ""):
        desc_paragraphs = [p.strip() for p in ticket_desc.split("\n") if p.strip()]
        initial_msg = desc_paragraphs[0] if desc_paragraphs else ticket_desc
        
        st.markdown(f"""
        <div class="conversation-box">
            <div class="conversation-bubble conversation-bubble-user">
                <div class="conversation-sender conversation-sender-user">
                    <span>👤 Initial Inquiry by {creator}</span>
                    <span class="conversation-time">📅 Opened: {start_date_val}</span>
                </div>
                <div class="conversation-text">{initial_msg}</div>
            </div>
            <div class="conversation-bubble conversation-bubble-support">
                <div class="conversation-sender conversation-sender-support">
                    <span>👨‍💻 Last Reply / Investigation by Ticket Owner: <strong>{ticket_owner}</strong></span>
                    <span class="conversation-time">⏱️ Status: Awaiting User Info ({aging} days)</span>
                </div>
                <div class="conversation-text"><strong>Latest Status Update:</strong> Ticket investigated by {ticket_owner}. Currently waiting on user clarification or confirmation from {creator} to proceed with resolution.</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if len(desc_paragraphs) > 1:
            st.markdown("##### 📋 Complete Raw Description")
            st.markdown(f"""
            <div class="dialog-raw-desc-box">{ticket_desc}</div>
            """, unsafe_allow_html=True)
    else:
        st.info(f"No specific description text was attached to ticket #{t_num} in the uploaded Excel file.")

    # Outgoing Follow-Up Reminder Draft
    if email_subject or email_body:
        st.markdown("#### ✉️ Outgoing Follow-Up Reminder Draft")
        st.text_input("Subject", email_subject, key=f"dlg_subj_{t_num}", disabled=True)
        st.text_area("Body", email_body, key=f"dlg_body_{t_num}", height=140, disabled=True)
        
    st.markdown("<hr style='margin: 1rem 0 0.8rem 0; border: none; border-top: 1px solid #e2e8f0;'>", unsafe_allow_html=True)
    
    # Bottom Action Bar
    col_act1, col_act2, col_close = st.columns([1.6, 1.6, 1])
    with col_act1:
        if mailto_url:
            st.link_button("✉️ Open in Outlook", mailto_url, use_container_width=True)
    with col_act2:
        if email_subject and cfg_sender_email:
            if not is_sent:
                if st.button("✅ Mark as Sent", key=f"dlg_mark_{t_num}", type="primary", use_container_width=True):
                    record_sent_reminder(
                        ticket_number=t_num,
                        creator_name=creator,
                        recipient_email=matched_addr or "",
                        subject=email_subject,
                        sent_by=cfg_sender_email,
                        delivery_mode="Outlook"
                    )
                    st.toast(f"✅ Marked Ticket #{t_num} as Sent in SQLite!", icon="🔴")
                    st.rerun()
            else:
                if st.button("↩️ Undo Sent", key=f"dlg_undo_{t_num}", type="secondary", use_container_width=True):
                    remove_sent_reminder(t_num)
                    st.toast(f"Reset Ticket #{t_num} to Not Sent.", icon="🔵")
                    st.rerun()
    with col_close:
        if st.button("✖️ Close", key=f"dlg_close_{t_num}", use_container_width=True):
            st.rerun()


def ticket_card(row: pd.Series):
    st.markdown(f"""
        <div class="ticket-card">
            <div class="ticket-header">
                🎫 {row['Ticket Number']} — {row['Subject']}
            </div>
            <div class="ticket-info-grid">
                <div class="ticket-info-item">
                    <div class="ticket-info-label">👤 Owner</div>
                    <div class="ticket-info-value">{row.get('Ticket Owner', '—')}</div>
                </div>
                <div class="ticket-info-item">
                    <div class="ticket-info-label">🏢 Customer</div>
                    <div class="ticket-info-value">{row.get('Customer Name', '—')}</div>
                </div>
                <div class="ticket-info-item">
                    <div class="ticket-info-label">⚡ Priority</div>
                    <div class="ticket-info-value">{row.get('Ticket Priority', '—')}</div>
                </div>
                <div class="ticket-info-item">
                    <div class="ticket-info-label">⏱️ Aging</div>
                    <div class="ticket-info-value">{row.get('Ticket Aging', '—')} days</div>
                </div>
                <div class="ticket-info-item">
                    <div class="ticket-info-label">📅 Start Date</div>
                    <div class="ticket-info-value">{row.get('Start Date', '—')}</div>
                </div>
                <div class="ticket-info-item">
                    <div class="ticket-info-label">📋 Required Date</div>
                    <div class="ticket-info-value">{row.get('Required Date', '—')}</div>
                </div>
                <div class="ticket-info-item">
                    <div class="ticket-info-label">🎯 Est. Resolution</div>
                    <div class="ticket-info-value">{row.get('Estimate Resolution Date', '—')}</div>
                </div>
                <div class="ticket-info-item">
                    <div class="ticket-info-label">✍️ Creator</div>
                    <div class="ticket-info-value">{row.get('Ticket Creator', '—')}</div>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    if str(row.get("Description", "")).strip() not in ("", "nan", "None"):
        with st.expander("📄 Description", expanded=True):
            st.write(row["Description"])
            t_num_card = str(row.get('Ticket Number', '')).strip()
            if st.button("🔍 Open Full Description in Center Modal Popup", key=f"btn_desc_lookup_{t_num_card}", use_container_width=True):
                show_ticket_description_dialog(
                    t_num=t_num_card,
                    creator=str(row.get('Ticket Creator', '—')).strip(),
                    matched_addr="",
                    t_subj=str(row.get('Subject', '')).strip(),
                    aging=int(row.get('Ticket Aging', 0) if pd.notna(row.get('Ticket Aging')) else 0),
                    ticket_owner=str(row.get('Ticket Owner', 'Unassigned')).strip(),
                    customer_name=str(row.get('Customer Name', '—')).strip(),
                    start_date_val=str(row.get('Start Date', '—')).strip(),
                    req_date_val=str(row.get('Required Date', '—')).strip(),
                    est_date_val=str(row.get('Estimate Resolution Date', '—')).strip(),
                    priority_val=str(row.get('Ticket Priority', 'Normal')).strip(),
                    dept_val=str(row.get('Ticket Department', 'Support')).strip(),
                    ticket_desc=str(row.get("Description", "")).strip()
                )


def related_list(df: pd.DataFrame, owner: str, exclude_ticket: str, empty_msg: str, offer_draft: bool, section_title: str):
    related = df[(df["Ticket Owner"] == owner) & (df["Ticket Number"] != exclude_ticket)]
    related = related.sort_values("Ticket Aging", ascending=False)
    
    st.markdown('<div class="related-section">', unsafe_allow_html=True)
    st.markdown(f'<div class="related-header">{section_title}</div>', unsafe_allow_html=True)
    
    if related.empty:
        st.markdown(f'<p style="color: #7f8c8d; font-style: italic;">{empty_msg}</p>', unsafe_allow_html=True)
    else:
        for _, r in related.iterrows():
            priority_color = {"High": "#e74c3c", "Medium": "#f39c12", "Low": "#27ae60"}.get(str(r.get('Ticket Priority', '')), "#34495e")
            st.markdown(f"""
                <div style="
                    background: #f8f9fa; 
                    padding: 1rem; 
                    border-radius: 8px; 
                    margin: 0.5rem 0;
                    border-left: 4px solid {priority_color};
                ">
                    <strong>🎫 {r['Ticket Number']}</strong> — {r['Subject']}<br>
                    <small style="color: #7f8c8d;">⏱️ {int(r['Ticket Aging'])} days • Priority: {r.get('Ticket Priority', 'Unknown')}</small>
                </div>
            """, unsafe_allow_html=True)
            
            if offer_draft and has_key_configured():
                if st.button("✍️ Draft reply", key=f"draft_{r['Ticket Number']}"):
                    with st.spinner("Drafting reply..."):
                        try:
                            text = draft_reply(r)
                            st.text_area(
                                f"Draft reply — {r['Ticket Number']}", text, height=150,
                                key=f"draft_text_{r['Ticket Number']}",
                            )
                        except Exception as e:
                            st.error(f"Couldn't generate a draft: {e}")
    
    st.markdown('</div>', unsafe_allow_html=True)


def render_chatbot_assistant(master_df, wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, mail_to_df=None):
    """Render an AI ticket assistant chat box that answers questions about uploaded Excel files."""
    st.divider()
    
    # Stay collapsed when fresh data is loaded so the user remains on their top dashboard!
    has_active_conversation = len(st.session_state.get("chatbot_messages", [])) > 1
    is_expanded = st.session_state.get("chatbot_open", False) or has_active_conversation

    with st.expander("💬 AI Ticket Assistant — Ask Questions About Your Data", expanded=is_expanded):
        st.markdown('<div class="chat-assistant-wrapper">', unsafe_allow_html=True)
        
        col_h1, col_h2 = st.columns([3, 1])
        with col_h1:
            st.markdown("""
            <div class="chat-header-title">
                <span>💬 AI Ticket Assistant</span>
                <span class="chat-header-badge">Smart AI</span>
            </div>
            <p style="color: #64748b; font-size: 0.88rem; margin: 0.2rem 0 0.6rem 0;">
                Ask any question about your uploaded Excel ticket data, owner workloads, priorities, aging, or specific ticket numbers.
            </p>
            """, unsafe_allow_html=True)
        with col_h2:
            if st.button("🗑️ Clear Chat", key="clear_chat_history", use_container_width=True):
                st.session_state["chatbot_messages"] = []
                st.session_state["chatbot_open"] = False
                st.rerun()

        data_ctx = generate_data_context(
            master_df=master_df,
            wip_df=wip_df,
            dev_df=dev_df,
            wait_df=wait_df,
            hold_df=hold_df,
            open_df=open_df,
            pending_df=pending_df,
            closed_df=closed_df,
            mail_to_df=mail_to_df
        )

        if "chatbot_messages" not in st.session_state:
            st.session_state["chatbot_messages"] = [
                {
                    "role": "assistant",
                    "content": "👋 Hi! I'm your AI Ticket Assistant. Ask me anything about your uploaded Excel files! For example: *'What is the ticket count breakdown?'*, *'Who has the most open tickets?'*, or search for a specific ticket number."
                }
            ]

        # Quick suggestion prompt chips
        st.markdown("**💡 Quick Questions:**")
        scol1, scol2, scol3, scol4 = st.columns(4)
        suggestion_prompt = None
        with scol1:
            if st.button("📊 Excel Data Summary", key="chip_summary", use_container_width=True):
                suggestion_prompt = "Give me a complete summary of all uploaded Excel ticket files, counts, and status."
        with scol2:
            if st.button("👥 Top Owner Workloads", key="chip_owners", use_container_width=True):
                suggestion_prompt = "Who are the top ticket owners and how many tickets does each owner have?"
        with scol3:
            if st.button("⏱️ Highest Aging Tickets", key="chip_aging", use_container_width=True):
                suggestion_prompt = "Which ticket categories or owners have the highest aging days?"
        with scol4:
            if st.button("⚠️ Pending & Hold Info", key="chip_pending", use_container_width=True):
                suggestion_prompt = "How many tickets are in Pending or Hold status and what are their details?"

        # Display chat message history
        for message in st.session_state["chatbot_messages"]:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        # Chat Input Box
        user_input = st.chat_input("Ask any question about your uploaded Excel files...")
        
        prompt = suggestion_prompt or user_input
        
        if prompt:
            st.session_state["chatbot_open"] = True
            st.session_state["chatbot_messages"].append({"role": "user", "content": prompt})
            
            with st.spinner("Analyzing Excel data with AI..."):
                response_text = get_chatbot_response(
                    user_question=prompt,
                    data_context=data_ctx,
                    history=st.session_state["chatbot_messages"]
                )
                st.session_state["chatbot_messages"].append({"role": "assistant", "content": response_text})
            st.rerun()

        st.markdown('</div>', unsafe_allow_html=True)



# ---------------------------------------------------------------- Sidebar
with st.sidebar:
    # ============ SQLITE DATABASE ============
    st.markdown('<div class="sidebar-header">🗄️ SQLite Database</div>', unsafe_allow_html=True)
    st.caption("Run the project directly using data stored in SQLite database (no file upload required):")
    
    if st.button("⚡ Load SQLite Database", key="btn_load_sqlite_demo", use_container_width=True, type="primary"):
        data = load_all_from_sqlite()
        if data:
            load_dataset_into_session(data, source_label="SQLite Database")
            st.toast("✅ SQLite Database loaded successfully!", icon="🚀")
            st.rerun()
        else:
            st.error("Could not load data from SQLite database.")

    if st.session_state.get("folder_data_loaded", False):
        if st.button("❌ Clear Loaded Data", key="btn_clear_loaded", use_container_width=True):
            clear_loaded_session_data()
            st.rerun()

    st.divider()

    st.markdown('<div class="sidebar-header">📂 Upload Ticket Reports</div>', unsafe_allow_html=True)
    
    # ============ FAST FOLDER UPLOAD (NEW!) ============
    st.markdown("### ⚡ Fast Folder Upload (New!)")
    st.caption("Upload a folder containing all 9 Excel files. Files are auto-detected by name and mapped to their categories.")
    
    folder_files = st.file_uploader(
        "📁 Select all Excel files from your folder",
        type="xlsx",
        accept_multiple_files=True,
        key="folder_upload",
        help="Upload all 9 files at once. Files will be auto-detected by filename (e.g., master_file.xlsx, mail_to.xlsx, etc.)"
    )
    
    if folder_files:
        folder_result = process_folder_upload(folder_files)
        is_valid, validation_msg = validate_folder_upload(folder_result)
        
        if is_valid:
            has_mailto = folder_result.get("mail_to") is not None
            if has_mailto:
                st.success("✅ All 8 ticket workflow files + Mail Directory detected and validated!")
            else:
                st.success("✅ All 8 ticket workflow files detected and validated!")
            summary = get_folder_upload_summary(folder_result)
            st.caption(summary)
            
            if st.button("📤 Use Folder Data", key="use_folder_data", use_container_width=True):
                st.session_state["folder_data_loaded"] = True
                st.session_state["master_df"] = folder_result.get("master")
                st.session_state["wip_df"] = folder_result.get("wip")
                st.session_state["dev_df"] = folder_result.get("dev")
                st.session_state["wait_df"] = folder_result.get("wait")
                st.session_state["hold_df"] = folder_result.get("hold")
                st.session_state["open_df"] = folder_result.get("open")
                st.session_state["pending_df"] = folder_result.get("pending")
                st.session_state["closed_df"] = folder_result.get("closed")
                st.session_state["mail_to_df"] = folder_result.get("mail_to", None)
                st.success("✅ Folder data loaded! Reloading app...")
                st.rerun()
        else:
            st.warning(f"⚠️ {validation_msg}")
    
    st.divider()
    st.markdown("### 📋 Manual Upload (Optional)")
    st.caption("Or upload files individually below:")
    
    # Compact upload sections with specific placeholders
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**📋 MASTER**")
        master_file = st.file_uploader("Master Ticket List", type="xlsx", key="master", label_visibility="collapsed", help="Master ticket list - all tickets")
        
        st.markdown("**📊 WIP**")
        wip_file = st.file_uploader("Work in Progress", type="xlsx", key="wip", label_visibility="collapsed", help="Work in Progress tickets")
        
        st.markdown("**🔧 DEV**")
        dev_file = st.file_uploader("Under Development", type="xlsx", key="dev", label_visibility="collapsed", help="Under Development tickets")
        
        st.markdown("**⚠️ WAIT**")
        wait_file = st.file_uploader("Awaiting User Info", type="xlsx", key="wait", label_visibility="collapsed", help="Awaiting User Info tickets")
    
    with col2:
        st.markdown("**⏸️ HOLD**")
        hold_file = st.file_uploader("Hold Tickets", type="xlsx", key="hold", label_visibility="collapsed", help="Hold tickets")
        
        st.markdown("**🟢 OPEN**")
        open_file = st.file_uploader("Open Tickets", type="xlsx", key="open", label_visibility="collapsed", help="Open tickets")
        
        st.markdown("**🟠 PENDING**")
        pending_file = st.file_uploader("Pending Tickets", type="xlsx", key="pending", label_visibility="collapsed", help="Pending tickets")
        
        st.markdown("**⚫ CLOSED**")
        closed_file = st.file_uploader("Closed Tickets", type="xlsx", key="closed", label_visibility="collapsed", help="Closed tickets")

    st.markdown("**✉️ MAIL TO (Email Directory)**")
    st.caption("Upload mapping file (Ticket Creator → Email ID, e.g. Sanket Tambe → sanket.tambe@unitread.co.in)")
    mail_to_file = st.file_uploader("Mail Directory Mapping", type=["xlsx", "xls", "csv"], key="mail_to", label_visibility="collapsed", help="Excel/CSV mapping Ticket Creators to their Email IDs")



# ---------------------------------------------------------------- Main
# Header section
st.markdown("""
<div class="landing-header-container">
    <div class="landing-header-left">
        <div class="landing-app-icon">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                <rect x="3" y="5" width="18" height="6" rx="3"></rect>
                <line x1="7" y1="8" x2="11" y2="8"></line>
                <rect x="3" y="13" width="18" height="6" rx="3"></rect>
                <line x1="7" y1="16" x2="11" y2="16"></line>
            </svg>
        </div>
        <div class="landing-header-titles">
            <h1 class="landing-title">Ticket Flow Tracker</h1>
            <p class="landing-subtitle">IT Operations • Ticket Analytics & Workflow Management</p>
        </div>
    </div>
    <div class="system-status-pill">
        <span class="status-pulse-dot"></span>
        <span class="status-text">System Online</span>
        <svg class="status-pulse-icon" width="16" height="12" viewBox="0 0 24 16" fill="none" stroke="#10b981" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round">
            <path d="M1 8h5l2.5-6 4 12 3-8 2.5 4h5" />
        </svg>
    </div>
</div>
""", unsafe_allow_html=True)


# Check if using folder data or manual uploads
use_folder_data = st.session_state.get("folder_data_loaded", False)

if use_folder_data:
    # Load from session state (folder upload)
    master_df = st.session_state.get("master_df")
    wip_df = st.session_state.get("wip_df")
    dev_df = st.session_state.get("dev_df")
    wait_df = st.session_state.get("wait_df")
    hold_df = st.session_state.get("hold_df")
    open_df = st.session_state.get("open_df")
    pending_df = st.session_state.get("pending_df")
    closed_df = st.session_state.get("closed_df")
    mail_to_df = st.session_state.get("mail_to_df", None)

    if any(df is None for df in [master_df, wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df]):
        st.markdown("""
        <div class="welcome-section">
            <div class="welcome-title">🚀 Folder Data Incomplete</div>
            <div class="welcome-text">Please re-upload your folder files or select all 8 Excel files in the sidebar.</div>
        </div>
        """, unsafe_allow_html=True)
        render_footer()
        st.stop()

    assert isinstance(master_df, pd.DataFrame)
    assert isinstance(wip_df, pd.DataFrame)
    assert isinstance(dev_df, pd.DataFrame)
    assert isinstance(wait_df, pd.DataFrame)
    assert isinstance(hold_df, pd.DataFrame)
    assert isinstance(open_df, pd.DataFrame)
    assert isinstance(pending_df, pd.DataFrame)
    assert isinstance(closed_df, pd.DataFrame)
else:
    # Load from manual uploads
    if not all([master_file, wip_file, dev_file, wait_file, hold_file, open_file, pending_file, closed_file]):
        col_left, col_center, col_right = st.columns([1, 1.6, 1])

        with col_center:
            with st.container(border=True):
                st.markdown("""
                <div class="qs-header-row" style="display: flex; align-items: flex-start; gap: 0.75rem; margin-bottom: 1.5rem;">
                    <span class="qs-bolt-icon" style="font-size: 1.4rem; color: #4f46e5; line-height: 1.1;">⚡</span>
                    <div>
                        <div class="qs-header-title" style="font-size: 1.22rem; font-weight: 700; color: #0f172a; line-height: 1.2;">Quick Start</div>
                        <div class="qs-header-subtitle" style="font-size: 0.85rem; color: #64748b; margin-top: 0.2rem;">Choose how you want to load your ticket data.</div>
                    </div>
                </div>
                <div class="qs-subcard-body" style="display: flex; flex-direction: column; justify-content: flex-start;">
                    <div class="qs-circle-icon qs-circle-blue" style="width: 44px; height: 44px; border-radius: 50%; background: #2563eb; color: #ffffff; display: flex; align-items: center; justify-content: center; margin-bottom: 0.9rem; box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);">
                        <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                            <ellipse cx="12" cy="6" rx="8" ry="3"></ellipse>
                            <path d="M4 6v6c0 1.66 3.58 3 8 3s8-1.34 8-3V6"></path>
                            <path d="M4 12v6c0 1.66 3.58 3 8 3s8-1.34 8-3v-6"></path>
                        </svg>
                    </div>
                    <div class="qs-card-title" style="font-size: 1.08rem; font-weight: 700; color: #0f172a; margin-bottom: 0.45rem;">Load SQLite Database</div>
                    <div class="qs-card-desc" style="font-size: 0.84rem; color: #64748b; line-height: 1.5; margin-bottom: 1.35rem;">
                        Launch instantly using our built-in SQLite database containing all 8 ticket workflow stages and mail directory.
                    </div>
                </div>
                """, unsafe_allow_html=True)

                if st.button("⚡ Load SQLite Database", key="btn_welcome_sqlite", use_container_width=True, type="primary"):
                    data = load_all_from_sqlite()
                    if data:
                        load_dataset_into_session(data, source_label="SQLite Database")
                        st.rerun()
                    else:
                        st.error("Failed to load SQLite data.")

        render_footer()
        st.stop()



    # Load 8 core workflow files
    master_df = load_excel(master_file)
    wip_df = load_excel(wip_file)
    dev_df = load_excel(dev_file) 
    wait_df = load_excel(wait_file)
    hold_df = load_excel(hold_file)
    open_df = load_excel(open_file)
    pending_df = load_excel(pending_file)
    closed_df = load_excel(closed_file)
    mail_to_df = load_excel(mail_to_file) if mail_to_file is not None else st.session_state.get("mail_to_df", None)

    # Store in session state so SQLite database tools can save it
    st.session_state["master_df"] = master_df
    st.session_state["wip_df"] = wip_df
    st.session_state["dev_df"] = dev_df
    st.session_state["wait_df"] = wait_df
    st.session_state["hold_df"] = hold_df
    st.session_state["open_df"] = open_df
    st.session_state["pending_df"] = pending_df
    st.session_state["closed_df"] = closed_df
    st.session_state["mail_to_df"] = mail_to_df

# Validate the 8 core ticket workflow files have required columns
files_to_validate = [
    (master_df, "Master Ticket List"),
    (wip_df, "Work in Progress"),
    (dev_df, "Under Development"), 
    (wait_df, "Awaiting User Info"),
    (hold_df, "Hold"),
    (open_df, "Open"),
    (pending_df, "Pending"),
    (closed_df, "Closed")
]

for df, label in files_to_validate:
    if not validate_columns(df, label):
        st.stop()

# Demo mode top banner if loaded from SQLite or demo folder
source_mode = st.session_state.get("data_source_mode")
if source_mode:
    badge_cls = "demo-badge-sqlite" if "SQLite" in source_mode else "demo-badge-folder"
    col_banner, col_banner_btn = st.columns([5, 1])
    with col_banner:
        contacts_count = len(mail_to_df) if mail_to_df is not None else 0
        st.markdown(f"""
        <div class="demo-banner">
            <div class="demo-banner-content">
                <span class="{badge_cls}">{source_mode}</span>
                <span>Active Demo Dataset: <strong>{len(master_df)} tickets</strong> across 8 workflow stages • <strong>{contacts_count} mail contacts</strong>. Ready for presentation & testing!</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col_banner_btn:
        if st.button("🔄 Clear Data", key="btn_banner_clear", use_container_width=True, help="Clear loaded demo data and return to upload screen"):
            clear_loaded_session_data()
            st.rerun()

# Enhanced metrics display for all 8 categories with fast count-up animation
st.html(f"""
<div class="metric-container">
    <div class="metric-card" style="border-left: 3px solid #17a2b8;">
        <div class="metric-value counter-value" data-target="{len(master_df)}" style="color: #17a2b8;">{len(master_df)}</div>
        <div class="metric-label">Master List</div>
    </div>
    <div class="metric-card wip-metric">
        <div class="metric-value counter-value" data-target="{len(wip_df)}" style="color: #007bff;">{len(wip_df)}</div>
        <div class="metric-label">Work in Progress</div>
    </div>
    <div class="metric-card dev-metric">
        <div class="metric-value counter-value" data-target="{len(dev_df)}" style="color: #fd7e14;">{len(dev_df)}</div>
        <div class="metric-label">Under Development</div>
    </div>
    <div class="metric-card wait-metric">
        <div class="metric-value counter-value" data-target="{len(wait_df)}" style="color: #dc3545;">{len(wait_df)}</div>
        <div class="metric-label">Awaiting User Info</div>
    </div>
</div>
<div class="metric-container">
    <div class="metric-card hold-metric">
        <div class="metric-value counter-value" data-target="{len(hold_df)}" style="color: #6f42c1;">{len(hold_df)}</div>
        <div class="metric-label">Hold</div>
    </div>
    <div class="metric-card open-metric">
        <div class="metric-value counter-value" data-target="{len(open_df)}" style="color: #28a745;">{len(open_df)}</div>
        <div class="metric-label">Open</div>
    </div>
    <div class="metric-card pending-metric">
        <div class="metric-value counter-value" data-target="{len(pending_df)}" style="color: #ffc107; text-shadow: 1px 1px 1px rgba(0,0,0,0.3);">{len(pending_df)}</div>
        <div class="metric-label">Pending</div>
    </div>
    <div class="metric-card closed-metric">
        <div class="metric-value counter-value" data-target="{len(closed_df)}" style="color: #6c757d;">{len(closed_df)}</div>
        <div class="metric-label">Closed</div>
    </div>
</div>

<script>
(function() {{
    let userHasScrolled = false;
    const releaseScroll = () => {{ 
        userHasScrolled = true;
        if (window._lockTopInterval) {{
            clearInterval(window._lockTopInterval);
            window._lockTopInterval = null;
        }}
    }};

    // Intercept programmatic scrollIntoView so lower elements don't yank the viewport down
    if (!window._origScrollIntoView) {{
        window._origScrollIntoView = Element.prototype.scrollIntoView;
        Element.prototype.scrollIntoView = function() {{
            if (!userHasScrolled) return;
            return window._origScrollIntoView.apply(this, arguments);
        }};
    }}

    // Intercept native focus so chat inputs or other lower elements cannot autofocus-scroll on load
    if (!window._origFocus) {{
        window._origFocus = HTMLElement.prototype.focus;
        HTMLElement.prototype.focus = function(options) {{
            if (this.tagName === 'TEXTAREA' && this.closest && this.closest('[data-testid="stChatInput"]')) {{
                if (!window._userEngagedWithChat) {{
                    return;
                }}
            }}
            return window._origFocus.apply(this, arguments);
        }};
    }}

    document.addEventListener('click', function(e) {{
        if (e.target && e.target.closest && e.target.closest('[data-testid="stChatInput"]')) {{
            window._userEngagedWithChat = true;
        }}
    }}, true);

    function lockTop() {{
        if (userHasScrolled) return;
        const targets = [
            document.querySelector('section.main'),
            document.querySelector('.stAppViewContainer'),
            document.querySelector('[data-testid="stAppViewContainer"]'),
            document.querySelector('[data-testid="stMain"]'),
            document.documentElement,
            document.body
        ];
        targets.forEach(t => {{
            if (t && t.scrollTop > 0) t.scrollTop = 0;
        }});
        if (window.scrollY > 0) window.scrollTo(0, 0);
    }}

    // Keep viewport anchored to top continuously across full page load
    lockTop();
    if (window._lockTopInterval) clearInterval(window._lockTopInterval);
    window._lockTopInterval = setInterval(lockTop, 30);
    setTimeout(() => {{
        if (window._lockTopInterval) {{
            clearInterval(window._lockTopInterval);
            window._lockTopInterval = null;
        }}
        lockTop();
    }}, 4500);

    // Only allow manual scroll after initial settle
    setTimeout(() => {{
        window.addEventListener('wheel', (e) => {{
            if (Math.abs(e.deltaY) > 8) releaseScroll();
        }}, {{ passive: true }});
        window.addEventListener('touchmove', releaseScroll, {{ passive: true }});
        window.addEventListener('keydown', (e) => {{
            if (['ArrowDown', 'ArrowUp', 'PageDown', 'PageUp', 'Space', 'Home', 'End'].includes(e.code)) {{
                releaseScroll();
            }}
        }}, {{ passive: true }});
    }}, 1000);


    function animateCounters() {{
        const counters = document.querySelectorAll('.counter-value[data-target]');
        if (!counters || !counters.length) return;

        lockTop();

        counters.forEach(counter => {{
            const target = parseInt(counter.getAttribute('data-target'), 10) || 0;
            if (counter._animating) return;
            counter._animating = true;

            if (target <= 1) {{
                counter.textContent = target;
                delete counter._animating;
                return;
            }}

            const startVal = 1;
            const duration = 850; // fast, energetic ~0.85s count-up
            const startTime = performance.now();

            function step(currentTime) {{
                const elapsed = currentTime - startTime;
                const progress = Math.min(elapsed / duration, 1);
                // Ease-out cubic: fast acceleration, smooth settle
                const ease = 1 - Math.pow(1 - progress, 3);
                const current = Math.floor(startVal + ease * (target - startVal));

                counter.textContent = current;

                if (progress < 1) {{
                    requestAnimationFrame(step);
                }} else {{
                    counter.textContent = target;
                    delete counter._animating;
                    lockTop();
                }}
            }}

            counter.textContent = startVal;
            requestAnimationFrame(step);
        }});
    }}

    animateCounters();
    requestAnimationFrame(animateCounters);
    setTimeout(animateCounters, 60);
}})();
</script>
""", unsafe_allow_javascript=True)

# Get combined date filters for global use
available_years, available_months = get_combined_date_filters(master_df, wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df)

st.divider()

# Search section - no container, just clean styling
st.markdown('<div class="search-section">', unsafe_allow_html=True)
st.markdown('<div class="search-title">🔍 Search a Ticket</div>', unsafe_allow_html=True)
ticket_no = st.text_input("Enter ticket number", placeholder="e.g. CT007948", key="ticket_search", label_visibility="collapsed").strip().upper()
st.markdown('</div>', unsafe_allow_html=True)

if ticket_no:
    # Use master ticket system to track workflow
    ticket_data, current_status, workflow_info = get_ticket_workflow_status(
        ticket_no, wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, master_df
    )
    
    if current_status == "not_in_master":
        st.error(f'🚫 **Ticket Not Found**: "{ticket_no}" does not exist in the Master Ticket List.')
        st.info("💡 **Note**: All tickets must first be created in the Master List before appearing in workflow categories.")
    else:
        # Display workflow status
        if isinstance(workflow_info, dict):
            st.markdown(f"""
            <div class="status-workflow" style="--status-color: {workflow_info.get('color', '#6c757d')};">
                <div class="status-workflow-title">{workflow_info.get('title', '')}</div>
                <div class="status-workflow-desc">{workflow_info.get('description', '')}</div>
                <div class="status-workflow-next">Next Step: {workflow_info.get('next_step', '')}</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Show status across all files for transparency
        wip_match = wip_df[wip_df["Ticket Number"].str.upper() == ticket_no]
        dev_match = dev_df[dev_df["Ticket Number"].str.upper() == ticket_no]
        wait_match = wait_df[wait_df["Ticket Number"].str.upper() == ticket_no]
        hold_match = hold_df[hold_df["Ticket Number"].str.upper() == ticket_no]
        open_match = open_df[open_df["Ticket Number"].str.upper() == ticket_no]
        pending_match = pending_df[pending_df["Ticket Number"].str.upper() == ticket_no]
        closed_match = closed_df[closed_df["Ticket Number"].str.upper() == ticket_no]

        st.markdown("### 📊 Status Tracking Across All Categories")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"""
            <div class="status-grid">
                <div class="status-card {'status-found' if not wip_match.empty else 'status-not-found'}">
                    <h4 style="margin-bottom: 0.5rem;">{'✅' if not wip_match.empty else '❌'} Work in Progress</h4>
                    <p style="margin: 0;">{'Found' if not wip_match.empty else 'Not Found'}</p>
                </div>
                <div class="status-card {'status-found' if not dev_match.empty else 'status-not-found'}">
                    <h4 style="margin-bottom: 0.5rem;">{'✅' if not dev_match.empty else '❌'} Under Development</h4>
                    <p style="margin: 0;">{'Found' if not dev_match.empty else 'Not Found'}</p>
                </div>
                <div class="status-card {'status-found' if not wait_match.empty else 'status-not-found'}">
                    <h4 style="margin-bottom: 0.5rem;">{'✅' if not wait_match.empty else '❌'} Awaiting User Info</h4>
                    <p style="margin: 0;">{'Found' if not wait_match.empty else 'Not Found'}</p>
                </div>
                <div class="status-card {'status-found' if not hold_match.empty else 'status-not-found'}">
                    <h4 style="margin-bottom: 0.5rem;">{'✅' if not hold_match.empty else '❌'} Hold</h4>
                    <p style="margin: 0;">{'Found' if not hold_match.empty else 'Not Found'}</p>
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="status-grid">
                <div class="status-card {'status-found' if not open_match.empty else 'status-not-found'}">
                    <h4 style="margin-bottom: 0.5rem;">{'✅' if not open_match.empty else '❌'} Open</h4>
                    <p style="margin: 0;">{'Found' if not open_match.empty else 'Not Found'}</p>
                </div>
                <div class="status-card {'status-found' if not pending_match.empty else 'status-not-found'}">
                    <h4 style="margin-bottom: 0.5rem;">{'✅' if not pending_match.empty else '❌'} Pending</h4>
                    <p style="margin: 0;">{'Found' if not pending_match.empty else 'Not Found'}</p>
                </div>
                <div class="status-card {'status-found' if not closed_match.empty else 'status-not-found'}">
                    <h4 style="margin-bottom: 0.5rem;">{'✅' if not closed_match.empty else '❌'} Closed</h4>
                    <p style="margin: 0;">{'Found' if not closed_match.empty else 'Not Found'}</p>
                </div>
                <div class="status-card status-found">
                    <h4 style="margin-bottom: 0.5rem;">✅ Master List</h4>
                    <p style="margin: 0;">Found</p>
                </div>
            </div>
            """, unsafe_allow_html=True)

        if ticket_data is not None:
            st.divider()
            st.markdown(f"### 🎫 Ticket Details - {workflow_info['title'] if workflow_info else 'Details'}")
            ticket_card(ticket_data)

            owner = ticket_data["Ticket Owner"]
            st.divider()
            st.markdown(f"### 👤 {owner}'s Other Active Tickets")
            
            # Show related tickets from all active categories (excluding closed)
            col1, col2, col3 = st.columns(3)
            with col1:
                related_list(wip_df, owner, ticket_data["Ticket Number"], 
                            "No other WIP tickets for this owner.", 
                            False, "📊 Work in Progress")
                related_list(hold_df, owner, ticket_data["Ticket Number"],
                            "No other hold tickets for this owner.",
                            False, "⏸️ Hold")
            with col2:
                related_list(dev_df, owner, ticket_data["Ticket Number"], 
                            "No other development tickets for this owner.", 
                            False, "🔧 Under Development")
                related_list(open_df, owner, ticket_data["Ticket Number"],
                            "No other open tickets for this owner.",
                            False, "🟢 Open")
            with col3:
                related_list(wait_df, owner, ticket_data["Ticket Number"], 
                            "No other tickets awaiting info for this owner.", 
                            True, "⚠️ Awaiting Reply")
                related_list(pending_df, owner, ticket_data["Ticket Number"],
                            "No other pending tickets for this owner.",
                            False, "🟠 Pending")

st.divider()

# Enhanced filtering section
st.markdown("### 🔍 Filter by Owner")
all_owners: Set[str] = set()
for df_item in [wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df]:
    if df_item is not None and not df_item.empty and "Ticket Owner" in df_item.columns:
        all_owners.update(str(o) for o in df_item["Ticket Owner"].dropna().unique())
all_owners_list = ["All owners"] + sorted(list(all_owners))

selected_owner = st.selectbox(
    "Select an owner to view their complete ticket statistics across all categories",
    all_owners_list,
    key="global_owner_filter"
)

# Display comprehensive user statistics
if selected_owner != "All owners":
    wip_user_count = len(wip_df[wip_df["Ticket Owner"] == selected_owner])
    dev_user_count = len(dev_df[dev_df["Ticket Owner"] == selected_owner])
    wait_user_count = len(wait_df[wait_df["Ticket Owner"] == selected_owner])
    hold_user_count = len(hold_df[hold_df["Ticket Owner"] == selected_owner])
    open_user_count = len(open_df[open_df["Ticket Owner"] == selected_owner])
    pending_user_count = len(pending_df[pending_df["Ticket Owner"] == selected_owner])
    closed_user_count = len(closed_df[closed_df["Ticket Owner"] == selected_owner])
    
    total_active_tickets = wip_user_count + dev_user_count + wait_user_count + hold_user_count + open_user_count + pending_user_count
    total_all_tickets = total_active_tickets + closed_user_count
    
    st.markdown(f"""
    <div class="filter-section">
        <div class="filter-title">📊 Complete Statistics for {selected_owner}</div>
        <div class="user-stats-container">
            <div class="user-stat-card" style="--border-color: #007bff; --number-color: #007bff;">
                <div class="stat-number animate-count">{wip_user_count}</div>
                <div class="stat-label">Work in Progress</div>
            </div>
            <div class="user-stat-card" style="--border-color: #fd7e14; --number-color: #fd7e14;">
                <div class="stat-number animate-count">{dev_user_count}</div>
                <div class="stat-label">Under Development</div>
            </div>
            <div class="user-stat-card" style="--border-color: #dc3545; --number-color: #dc3545;">
                <div class="stat-number animate-count">{wait_user_count}</div>
                <div class="stat-label">Awaiting User Info</div>
            </div>
            <div class="user-stat-card" style="--border-color: #6f42c1; --number-color: #6f42c1;">
                <div class="stat-number animate-count">{hold_user_count}</div>
                <div class="stat-label">Hold</div>
            </div>
        </div>
        <div class="user-stats-container">
            <div class="user-stat-card" style="--border-color: #28a745; --number-color: #28a745;">
                <div class="stat-number animate-count">{open_user_count}</div>
                <div class="stat-label">Open</div>
            </div>
            <div class="user-stat-card" style="--border-color: #ffc107; --number-color: #856404;">
                <div class="stat-number animate-count">{pending_user_count}</div>
                <div class="stat-label">Pending</div>
            </div>
            <div class="user-stat-card" style="--border-color: #6c757d; --number-color: #6c757d;">
                <div class="stat-number animate-count">{closed_user_count}</div>
                <div class="stat-label">Closed</div>
            </div>
            <div class="user-stat-card" style="--border-color: #17a2b8; --number-color: #17a2b8;">
                <div class="stat-number animate-count">{total_active_tickets}</div>
                <div class="stat-label">Total Active</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    if total_active_tickets > 0:
        # Enhanced workload analysis
        if wait_user_count > 0:
            st.warning(f"⚠️ **Action Needed**: {selected_owner} has {wait_user_count} ticket(s) awaiting user information that need attention.")
        if hold_user_count > 0:
            st.info(f"⏸️ **On Hold**: {selected_owner} has {hold_user_count} ticket(s) on hold - review for potential release.")
        if dev_user_count >= 5:
            st.info(f"💼 **High Workload**: {selected_owner} is actively working on {dev_user_count} tickets in development.")
        if total_active_tickets >= 15:
            st.error(f"🚨 **Overloaded**: {selected_owner} has {total_active_tickets} total active tickets - consider redistributing workload.")
        if closed_user_count > 0:
            st.success(f"✅ **Productivity**: {selected_owner} has completed {closed_user_count} ticket(s).")

st.divider()

tab0, tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📋 Master List", "📊 Work in Progress", "🔧 Under Development", "⚠️ Awaiting User Info", 
    "⏸️ Hold", "🟢 Open", "🟠 Pending", "✅ Closed"
])

with tab0:
    # Filters for Master List
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
        owners = ["All owners"] + sorted(master_df["Ticket Owner"].dropna().unique().tolist())
        master_owner_filter = st.selectbox("Master Owner", owners, key="master_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Creator</div>', unsafe_allow_html=True)
        creators = ["All creators"] + (sorted(master_df["Ticket Creator"].dropna().unique().tolist()) if "Ticket Creator" in master_df.columns else [])
        master_creator_filter = st.selectbox("Master Creator", creators, key="master_creator_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        master_year_filter = st.selectbox("Master Year", available_years, key="master_year_filter", label_visibility="collapsed")
    with col4:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        master_month_filter = st.selectbox("Master Month", available_months, key="master_month_filter", label_visibility="collapsed")
    
    # Display prominent filter results
    display_filter_results(master_df, "Master List", master_owner_filter, master_year_filter, master_month_filter, "#17a2b8", creator_filter=master_creator_filter)
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, master_owner_filter, master_year_filter, master_month_filter)
    
    # Filter and display data
    master_view = get_filtered_data(master_df, master_owner_filter, master_year_filter, master_month_filter, creator_filter=master_creator_filter)
    st.dataframe(master_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)


with tab1:
    # Filters for Work in Progress
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
        owners = ["All owners"] + sorted(wip_df["Ticket Owner"].dropna().unique().tolist())
        wip_owner_filter = st.selectbox("WIP Owner", owners, key="wip_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        wip_year_filter = st.selectbox("WIP Year", available_years, key="wip_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        wip_month_filter = st.selectbox("WIP Month", available_months, key="wip_month_filter", label_visibility="collapsed")
    
    # Display prominent filter results
    display_filter_results(wip_df, "Work in Progress", wip_owner_filter, wip_year_filter, wip_month_filter, "#007bff")
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, wip_owner_filter, wip_year_filter, wip_month_filter)
    
    # Filter and display data
    wip_view = get_filtered_data(wip_df, wip_owner_filter, wip_year_filter, wip_month_filter)
    st.dataframe(wip_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)

with tab2:
    # Filters for Under Development
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
        owners = ["All owners"] + sorted(dev_df["Ticket Owner"].dropna().unique().tolist())
        dev_owner_filter = st.selectbox("Dev Owner", owners, key="dev_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        dev_year_filter = st.selectbox("Dev Year", available_years, key="dev_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        dev_month_filter = st.selectbox("Dev Month", available_months, key="dev_month_filter", label_visibility="collapsed")
    
    # Display prominent filter results
    display_filter_results(dev_df, "Under Development", dev_owner_filter, dev_year_filter, dev_month_filter, "#fd7e14")
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, dev_owner_filter, dev_year_filter, dev_month_filter)
    
    # Filter and display data
    dev_view = get_filtered_data(dev_df, dev_owner_filter, dev_year_filter, dev_month_filter)
    st.dataframe(dev_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)

with tab3:
    # Filters for Awaiting User Info
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
        owners = ["All owners"] + sorted(wait_df["Ticket Owner"].dropna().unique().tolist())
        wait_owner_filter = st.selectbox("Wait Owner", owners, key="wait_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Creator</div>', unsafe_allow_html=True)
        wait_creators = ["All creators"] + (sorted(wait_df["Ticket Creator"].dropna().unique().tolist()) if "Ticket Creator" in wait_df.columns else [])
        wait_creator_filter = st.selectbox("Wait Creator", wait_creators, key="wait_creator_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        wait_year_filter = st.selectbox("Wait Year", available_years, key="wait_year_filter", label_visibility="collapsed")
    with col4:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        wait_month_filter = st.selectbox("Wait Month", available_months, key="wait_month_filter", label_visibility="collapsed")
    
    # Display prominent filter results
    display_filter_results(wait_df, "Awaiting User Info", wait_owner_filter, wait_year_filter, wait_month_filter, "#dc3545", creator_filter=wait_creator_filter)
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, wait_owner_filter, wait_year_filter, wait_month_filter)
    
    wait_view = get_filtered_data(wait_df, wait_owner_filter, wait_year_filter, wait_month_filter, creator_filter=wait_creator_filter)


    # =========================================================================
    # 📢 TICKET CREATOR REMINDER DISPATCHER
    # =========================================================================
    email_directory = build_email_directory(mail_to_df)
    
    st.markdown("""
    <div class="reminder-dispatcher-container">
        <div class="reminder-header-row">
            <div class="reminder-title">
                <span>📢 Ticket Creator Reminder Dispatcher</span>
            </div>
        </div>
        <div class="reminder-subtitle">
            Send follow-up reminders to creators so ticket owners can get replies and move tickets forward or close them.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Show Email Directory Connection Status
    if email_directory:
        st.success(f"✅ **Mail Directory Active**: {len(email_directory)} creator emails mapped from Mail To file.")
        with st.expander(f"📋 View Email Directory Mappings ({len(email_directory)} entries)", expanded=False):
            mapping_rows = [{"Creator / Name": k, "Email Address": v} for k, v in list(email_directory.items()) if k == k.lower()]
            if mapping_rows:
                st.dataframe(pd.DataFrame(mapping_rows), use_container_width=True, hide_index=True)
    else:
        st.warning("⚠️ **Mail Directory Not Loaded**: Please upload your **Mail To** file in the sidebar to auto-match each ticket creator's email address.")

    # Sender and Email Dispatch Settings
    with st.expander("⚙️ Sender & Email Settings (Defaults: Sandeep Yadav | United Tyrekrafts)", expanded=False):
        c_set1, c_set2 = st.columns(2)
        with c_set1:
            cfg_sender_name = st.text_input("Sender Name", value="Sandeep Yadav", key="rem_sender_name")
            cfg_sender_company = st.text_input("Company Name", value="United Tyrekrafts Pvt. Ltd.", key="rem_sender_company")
            cfg_sender_email = st.text_input("Sender Email Address", value=os.environ.get("SENDER_EMAIL", "utkerp@outlook.com"), key="rem_sender_email")
            cfg_sender_phone = st.text_input("Contact Phone Number", value=os.environ.get("SENDER_PHONE", "9682548514"), key="rem_sender_phone")
        with c_set2:
            cfg_smtp_host = st.text_input("SMTP Server (for direct 1-click background sending)", value=os.environ.get("SMTP_HOST", "smtp.office365.com"), key="rem_smtp_host")
            cfg_smtp_port = st.number_input("SMTP Port", value=int(os.environ.get("SMTP_PORT", "587")), step=1, key="rem_smtp_port")
            cfg_smtp_password = st.text_input("SMTP / App Password (Optional)", type="password", value=os.environ.get("SMTP_PASSWORD", ""), key="rem_smtp_password", help="Enter app password to send directly in background without opening Outlook")

    # Match tickets with emails
    reminder_list: List[Dict[str, Any]] = []
    for _, r in wait_view.iterrows():
        creator_val = str(r.get("Ticket Creator", "")).strip()
        matched_email = find_creator_email(creator_val, email_directory)
        sub_text, body_text = generate_reminder_email(r, cfg_sender_name, cfg_sender_company, cfg_sender_email, cfg_sender_phone)
        
        mailto_url = generate_mailto_url(matched_email, sub_text, body_text) if matched_email else ""
            
        reminder_list.append({
            "ticket_no": r.get("Ticket Number", ""),
            "creator": creator_val,
            "owner": str(r.get("Ticket Owner", "Unassigned")).strip(),
            "matched_email": matched_email,
            "subject": r.get("Subject", ""),
            "aging": r.get("Ticket Aging", 0),
            "department": r.get("Ticket Department", "Support"),
            "email_subject": sub_text,
            "email_body": body_text,
            "mailto_url": mailto_url,
            "raw_row": r
        })

    matched_count = sum(1 for item in reminder_list if item["matched_email"])
    unmatched_count = len(reminder_list) - matched_count

    # Load persistent sent reminder records from SQLite
    sent_reminders_map = get_sent_reminders_map()
    total_sent_in_db = len(sent_reminders_map)

    # Bulk Action Bar
    rcol1, rcol2, rcol3 = st.columns([2, 1, 1])
    with rcol1:
        st.markdown(f"**Tickets in view:** `{len(reminder_list)}` | **Emails Matched:** `{matched_count}` | **Already Sent:** `{sum(1 for it in reminder_list if str(it['ticket_no']).strip() in sent_reminders_map)}`")
    with rcol2:
        send_all_btn = st.button("🚀 Send All Reminders (Single Click)", type="primary", use_container_width=True, key="btn_send_all_reminders")
    with rcol3:
        show_preview = st.checkbox("Show Email Cards", value=True, key="chk_show_reminder_preview")

    # Handle Send All Click (via SMTP background)
    if send_all_btn:
        if not reminder_list:
            st.info("No tickets in current view to send reminders for.")
        elif not cfg_smtp_password:
            st.warning(
                f"🔑 **SMTP Password Needed for Bulk Background Sending**\n\n"
                f"To send emails automatically in the background from `{cfg_sender_email}`, please expand **'⚙️ Sender & Email Settings'** above and enter your email/app password.\n\n"
                f"💡 **Immediate Alternative**: Click the **'✉️ Send via Outlook'** button next to each ticket below to open it pre-filled and log it with 1 click!"
            )
        else:
            with st.spinner(f"Sending reminders from {cfg_sender_email}..."):
                sent_success = 0
                sent_failed = 0
                error_msgs = []
                progress_bar = st.progress(0)
                
                for idx, item in enumerate(reminder_list):
                    recipient = item["matched_email"]
                    if recipient:
                        ok, msg = send_smtp_email(
                            smtp_host=cfg_smtp_host,
                            smtp_port=int(cfg_smtp_port),
                            sender_email=cfg_sender_email,
                            sender_password=cfg_smtp_password,
                            recipient_email=recipient,
                            subject=item["email_subject"],
                            body=item["email_body"]
                        )
                        if ok:
                            sent_success += 1
                            record_sent_reminder(
                                ticket_number=item["ticket_no"],
                                creator_name=item["creator"],
                                recipient_email=recipient,
                                subject=item["email_subject"],
                                sent_by=cfg_sender_email,
                                delivery_mode="SMTP"
                            )
                        else:
                            sent_failed += 1
                            error_msgs.append(f"{item['ticket_no']} ({recipient}): {msg}")
                    else:
                        sent_failed += 1
                        error_msgs.append(f"{item['ticket_no']} ({item['creator']}): No email address found")
                    
                    progress_bar.progress((idx + 1) / len(reminder_list))
                
                # Refresh map after bulk send
                sent_reminders_map = get_sent_reminders_map()
                
                if sent_success > 0:
                    st.success(f"🎉 **Completed**: Successfully sent & logged {sent_success} reminder emails to SQLite!")
                if sent_failed > 0:
                    st.error(f"⚠️ {sent_failed} email(s) could not be sent:")
                    for err in error_msgs[:5]:
                        st.caption(f"- {err}")

    # =========================================================================
    # 📬 SENT MAILS HISTORY & LOGS (Prevents Overriding & Provides Full Visibility)
    # =========================================================================
    with st.expander(f"📬 Sent Reminders History ({total_sent_in_db} logged in SQLite)", expanded=False):
        sent_df = get_sent_reminders_dataframe()
        if not sent_df.empty:
            st.markdown("Below is the persistent audit trail of all reminders sent. Each ticket is permanently logged so you won't duplicate or override sent notices.")
            st.dataframe(sent_df, use_container_width=True, hide_index=True)
            col_dl1, col_dl2 = st.columns([1, 3])
            with col_dl1:
                csv_sent = sent_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Export Sent Log (CSV)",
                    data=csv_sent,
                    file_name="sent_reminders_log.csv",
                    mime="text/csv",
                    key="btn_download_sent_log"
                )
        else:
            st.info("No reminders sent yet. Once you send via Outlook or SMTP, the details will appear here automatically.")

    # Reminder Cards / Direct Outlook buttons
    if show_preview and reminder_list:
        st.markdown("#### 📬 Ticket Follow-Up Queue")
        st.caption("🔴 **Red Border / Badge**: Mail has already been sent (persisted in SQLite) | 🔵 **Light Blue Border / Badge**: Reminder not yet sent")
        
        for item in reminder_list:
            t_num = str(item["ticket_no"]).strip()
            creator = item["creator"]
            matched_addr = item["matched_email"]
            t_subj = item["subject"]
            aging = item["aging"]
            mailto_link = item["mailto_url"]
            
            # Check sent status in SQLite
            is_sent = t_num in sent_reminders_map
            sent_info = sent_reminders_map.get(t_num)
            
            # Card style: Red if sent, Light Blue if not sent
            card_class = "reminder-card-sent" if is_sent else "reminder-card-not-sent"
            status_badge = (
                f'<span class="reminder-badge-sent">🔴 Sent ({sent_info["sent_at"][:16] if sent_info else "Yes"})</span>' 
                if is_sent 
                else '<span class="reminder-badge-not-sent">🔵 Not Sent</span>'
            )
            email_badge = f'<span class="reminder-email-tag">✉️ {matched_addr}</span>' if matched_addr else '<span class="reminder-email-missing">⚠️ Email not found in directory</span>'
            
            c_left, c_right = st.columns([3.6, 1.6])
            with c_left:
                st.markdown(f"""
                <div class="{card_class}">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.3rem;">
                        <strong>🎫 {t_num}</strong> — {t_subj}
                        {status_badge}
                    </div>
                    <div style="font-size: 0.85rem; color: #475569; margin-top: 0.2rem;">
                        👤 Creator: <strong>{creator}</strong> {email_badge} • ⏱️ {aging} days
                    </div>
                </div>
                """, unsafe_allow_html=True)
            with c_right:
                if mailto_link:
                    # 1. Direct browser mailto link - guaranteed to launch Outlook immediately on click
                    outlook_btn_cls = "reminder-outlook-link-resend" if is_sent else "reminder-outlook-link"
                    outlook_btn_txt = "✉️ Re-open in Outlook" if is_sent else "✉️ Open in Outlook"
                    
                    st.markdown(f"""
                    <div style="margin-top: 0.2rem; margin-bottom: 0.4rem;">
                        <a href="{mailto_link}" class="{outlook_btn_cls}" style="width: 100%; justify-content: center; box-sizing: border-box;">
                            {outlook_btn_txt}
                        </a>
                    </div>
                    """, unsafe_allow_html=True)

                    # 2. Manual status toggle button: saves to SQLite and turns card Red (or toggles back)
                    if not is_sent:
                        if st.button("✅ Mark as Sent", key=f"btn_mark_sent_{t_num}", type="primary", use_container_width=True, help="Click once you send the email in Outlook to turn this card red and log to SQLite"):
                            record_sent_reminder(
                                ticket_number=t_num,
                                creator_name=creator,
                                recipient_email=matched_addr or "",
                                subject=item["email_subject"],
                                sent_by=cfg_sender_email,
                                delivery_mode="Outlook"
                            )
                            st.toast(f"✅ Marked Ticket #{t_num} as Sent in SQLite!", icon="🔴")
                            st.rerun()
                    else:
                        if st.button("↩️ Undo Sent", key=f"btn_undo_sent_{t_num}", type="secondary", use_container_width=True, help="Reset status back to light blue (Not Sent)"):
                            remove_sent_reminder(t_num)
                            st.toast(f"Reset Ticket #{t_num} to Not Sent.", icon="🔵")
                            st.rerun()
                else:
                    st.caption("No email mapped")

            raw_ticket = item.get("raw_row")
            ticket_desc = str(raw_ticket.get("Description", "")).strip() if raw_ticket is not None else ""
            ticket_owner = str(item.get("owner") or (raw_ticket.get("Ticket Owner", "Unassigned") if raw_ticket is not None else "Unassigned")).strip()
            customer_name = str(raw_ticket.get("Customer Name", "—")).strip() if raw_ticket is not None else "—"
            start_date_val = str(raw_ticket.get("Start Date", "—")).strip() if raw_ticket is not None else "—"
            req_date_val = str(raw_ticket.get("Required Date", "—")).strip() if raw_ticket is not None else "—"
            est_date_val = str(raw_ticket.get("Estimate Resolution Date", "—")).strip() if raw_ticket is not None else "—"
            priority_val = str(raw_ticket.get("Ticket Priority", "Normal")).strip() if raw_ticket is not None else "Normal"
            dept_val = str(raw_ticket.get("Ticket Department", "Support")).strip() if raw_ticket is not None else "Support"
            # Button to trigger the Modal Dialog Popup (Floating above UI in center)
            if st.button(
                f"💬 View Conversation & Description for {t_num} ({creator})",
                key=f"btn_desc_modal_{t_num}",
                use_container_width=True,
                help="Click to open full ticket description and conversation in a centered popup modal"
            ):
                show_ticket_description_dialog(
                    t_num=t_num,
                    creator=creator,
                    matched_addr=matched_addr,
                    t_subj=t_subj,
                    aging=aging,
                    ticket_owner=ticket_owner,
                    customer_name=customer_name,
                    start_date_val=start_date_val,
                    req_date_val=req_date_val,
                    est_date_val=est_date_val,
                    priority_val=priority_val,
                    dept_val=dept_val,
                    ticket_desc=ticket_desc,
                    email_subject=item.get("email_subject", ""),
                    email_body=item.get("email_body", ""),
                    mailto_url=mailto_link,
                    is_sent=is_sent,
                    cfg_sender_email=cfg_sender_email
                )



    st.divider()
    st.markdown("#### 📋 Awaiting User Info Tickets Table")
    st.dataframe(wait_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)


with tab4:
    # Filters for Hold
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
        owners = ["All owners"] + sorted(hold_df["Ticket Owner"].dropna().unique().tolist())
        hold_owner_filter = st.selectbox("Hold Owner", owners, key="hold_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        hold_year_filter = st.selectbox("Hold Year", available_years, key="hold_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        hold_month_filter = st.selectbox("Hold Month", available_months, key="hold_month_filter", label_visibility="collapsed")
    
    # Display prominent filter results
    display_filter_results(hold_df, "Hold", hold_owner_filter, hold_year_filter, hold_month_filter, "#6f42c1")
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, hold_owner_filter, hold_year_filter, hold_month_filter)
    
    # Filter and display data
    hold_view = get_filtered_data(hold_df, hold_owner_filter, hold_year_filter, hold_month_filter)
    st.dataframe(hold_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)

with tab5:
    # Filters for Open
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
        owners = ["All owners"] + sorted(open_df["Ticket Owner"].dropna().unique().tolist())
        open_owner_filter = st.selectbox("Open Owner", owners, key="open_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        open_year_filter = st.selectbox("Open Year", available_years, key="open_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        open_month_filter = st.selectbox("Open Month", available_months, key="open_month_filter", label_visibility="collapsed")
    
    # Display prominent filter results
    display_filter_results(open_df, "Open", open_owner_filter, open_year_filter, open_month_filter, "#28a745")
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, open_owner_filter, open_year_filter, open_month_filter)
    
    # Filter and display data
    open_view = get_filtered_data(open_df, open_owner_filter, open_year_filter, open_month_filter)
    st.dataframe(open_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)

with tab6:
    # Filters for Pending
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
        owners = ["All owners"] + sorted(pending_df["Ticket Owner"].dropna().unique().tolist())
        pending_owner_filter = st.selectbox("Pending Owner", owners, key="pending_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        pending_year_filter = st.selectbox("Pending Year", available_years, key="pending_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        pending_month_filter = st.selectbox("Pending Month", available_months, key="pending_month_filter", label_visibility="collapsed")
    
    # Display prominent filter results
    display_filter_results(pending_df, "Pending", pending_owner_filter, pending_year_filter, pending_month_filter, "#ffc107")
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, pending_owner_filter, pending_year_filter, pending_month_filter)
    
    # Filter and display data
    pending_view = get_filtered_data(pending_df, pending_owner_filter, pending_year_filter, pending_month_filter)
    st.dataframe(pending_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)

with tab7:
    # Filters for Closed
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
        owners = ["All owners"] + sorted(closed_df["Ticket Owner"].dropna().unique().tolist())
        closed_owner_filter = st.selectbox("Closed Owner", owners, key="closed_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        closed_year_filter = st.selectbox("Closed Year", available_years, key="closed_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        closed_month_filter = st.selectbox("Closed Month", available_months, key="closed_month_filter", label_visibility="collapsed")
    
    # Display prominent filter results
    display_filter_results(closed_df, "Closed", closed_owner_filter, closed_year_filter, closed_month_filter, "#6c757d")
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, closed_owner_filter, closed_year_filter, closed_month_filter)
    
    # Filter and display data
    closed_view = get_filtered_data(closed_df, closed_owner_filter, closed_year_filter, closed_month_filter)
    st.dataframe(closed_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)

# ---------------------------------------------------------------- AI Assistant & Footer
render_chatbot_assistant(master_df, wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, mail_to_df)
render_footer()

