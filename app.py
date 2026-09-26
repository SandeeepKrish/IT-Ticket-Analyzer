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
from typing import Any, Set

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

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

load_dotenv()

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


def display_filter_results(df: pd.DataFrame, category_name: str, owner_filter: Any, year_filter: Any, month_filter: Any, category_color: str):
    """Display prominent filter results with ticket counts and breakdowns"""
    
    # Get filtered data
    filtered_df = get_filtered_data(df, owner_filter, year_filter, month_filter)
    total_tickets = len(filtered_df)
    
    # Build filter description
    filter_parts = []
    if owner_filter != "All owners":
        filter_parts.append(f"Owner: {owner_filter}")
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
        if not df.empty and "Start Date" in df.columns:
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


def get_filtered_data(df: pd.DataFrame, owner_filter: Any, year_filter=None, month_filter=None) -> pd.DataFrame:
    """Filter dataframe by owner, year, and month"""
    filtered_df = df.copy()
    
    # Filter by owner
    if owner_filter != "All owners":
        filtered_df = filtered_df[filtered_df["Ticket Owner"] == owner_filter]
    
    # Filter by date if filters are provided
    if year_filter is not None and month_filter is not None:
        filtered_df = filter_by_date(filtered_df, year_filter, month_filter)
    
    return filtered_df


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


def render_chatbot_assistant(master_df, wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df):
    """Render an AI ticket assistant chat box that answers questions about uploaded Excel files."""
    st.divider()
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
            st.rerun()

    data_ctx = generate_data_context(
        master_df=master_df,
        wip_df=wip_df,
        dev_df=dev_df,
        wait_df=wait_df,
        hold_df=hold_df,
        open_df=open_df,
        pending_df=pending_df,
        closed_df=closed_df
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

    # Chat Input Box at the bottom
    user_input = st.chat_input("Ask any question about your uploaded Excel files...")
    
    prompt = suggestion_prompt or user_input
    
    if prompt:
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
    st.markdown('<div class="sidebar-header">📂 Upload Ticket Reports</div>', unsafe_allow_html=True)
    
    # ============ FAST FOLDER UPLOAD (NEW!) ============
    st.markdown("### ⚡ Fast Folder Upload (New!)")
    st.caption("Upload a folder containing all 8 Excel files. Files are auto-detected by name and mapped to their categories.")
    
    folder_files = st.file_uploader(
        "📁 Select all Excel files from your folder",
        type="xlsx",
        accept_multiple_files=True,
        key="folder_upload",
        help="Upload all 8 files at once. Files will be auto-detected by filename (e.g., master_file.xlsx, hold_file.xlsx, etc.)"
    )
    
    if folder_files:
        folder_result = process_folder_upload(folder_files)
        is_valid, validation_msg = validate_folder_upload(folder_result)
        
        if is_valid:
            st.success("✅ All 8 files detected and validated!")
            summary = get_folder_upload_summary(folder_result)
            st.caption(summary)
            
            if st.button("📤 Use Folder Data", key="use_folder_data", use_container_width=True):
                st.session_state["folder_data_loaded"] = True
                st.session_state["master_df"] = folder_result["master"]
                st.session_state["wip_df"] = folder_result["wip"]
                st.session_state["dev_df"] = folder_result["dev"]
                st.session_state["wait_df"] = folder_result["wait"]
                st.session_state["hold_df"] = folder_result["hold"]
                st.session_state["open_df"] = folder_result["open"]
                st.session_state["pending_df"] = folder_result["pending"]
                st.session_state["closed_df"] = folder_result["closed"]
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
        master_file = st.file_uploader("", type="xlsx", key="master", help="Master ticket list - all tickets")
        
        st.markdown("**📊 WIP**")
        wip_file = st.file_uploader("", type="xlsx", key="wip", help="Work in Progress tickets")
        
        st.markdown("**🔧 DEV**")
        dev_file = st.file_uploader("", type="xlsx", key="dev", help="Under Development tickets")
        
        st.markdown("**⚠️ WAIT**")
        wait_file = st.file_uploader("", type="xlsx", key="wait", help="Awaiting User Info tickets")
    
    with col2:
        st.markdown("**⏸️ HOLD**")
        hold_file = st.file_uploader("", type="xlsx", key="hold", help="Hold tickets")
        
        st.markdown("**🟢 OPEN**")
        open_file = st.file_uploader("", type="xlsx", key="open", help="Open tickets")
        
        st.markdown("**🟠 PENDING**")
        pending_file = st.file_uploader("", type="xlsx", key="pending", help="Pending tickets")
        
        st.markdown("**⚫ CLOSED**")
        closed_file = st.file_uploader("", type="xlsx", key="closed", help="Closed tickets")

    st.divider()
    st.markdown('<div class="sidebar-header">🤖 AI Assistant</div>', unsafe_allow_html=True)
    st.caption("Add an OpenAI API key to enable one-click draft replies for tickets awaiting user info.")
    key_input = st.text_input("OpenAI API key", type="password", placeholder="sk-...", key="api_key")
    
    # Use environment variable if no key is entered in UI
    if key_input:
        st.session_state["OPENAI_API_KEY"] = key_input
        os.environ["OPENAI_API_KEY"] = key_input
    elif not key_input and os.environ.get("OPENAI_API_KEY"):
        st.success("✅ API key loaded from .env file")

# ---------------------------------------------------------------- Main
# Header section
st.markdown("""
<div class="main-header">
    <div class="main-title">🎫 Ticket Flow Tracker</div>
    <div class="main-subtitle">Work in Progress → Under Development → Awaiting User Info, mapped by owner so nothing sits stuck</div>
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
        st.markdown("""
        <div class="welcome-section">
            <div class="welcome-title">🚀 Upload Required Files</div>
            <div class="welcome-text">Please upload all 8 ticket export files in the sidebar to begin comprehensive ticket analysis.</div>
        </div>
        """, unsafe_allow_html=True)
        render_footer()
        st.stop()

    # Load all 8 files including master
    master_df = load_excel(master_file)
    wip_df = load_excel(wip_file)
    dev_df = load_excel(dev_file) 
    wait_df = load_excel(wait_file)
    hold_df = load_excel(hold_file)
    open_df = load_excel(open_file)
    pending_df = load_excel(pending_file)
    closed_df = load_excel(closed_file)

# Validate all files have required columns
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

# Enhanced metrics display for all categories + master
st.markdown(f"""
<div class="metric-container">
    <div class="metric-card" style="border-left: 3px solid #17a2b8;">
        <div class="metric-value" style="color: #17a2b8;">{len(master_df)}</div>
        <div class="metric-label">Master List</div>
    </div>
    <div class="metric-card wip-metric">
        <div class="metric-value" style="color: #007bff;">{len(wip_df)}</div>
        <div class="metric-label">Work in Progress</div>
    </div>
    <div class="metric-card dev-metric">
        <div class="metric-value" style="color: #fd7e14;">{len(dev_df)}</div>
        <div class="metric-label">Under Development</div>
    </div>
    <div class="metric-card wait-metric">
        <div class="metric-value" style="color: #dc3545;">{len(wait_df)}</div>
        <div class="metric-label">Awaiting User Info</div>
    </div>
</div>
<div class="metric-container">
    <div class="metric-card hold-metric">
        <div class="metric-value" style="color: #6f42c1;">{len(hold_df)}</div>
        <div class="metric-label">Hold</div>
    </div>
    <div class="metric-card open-metric">
        <div class="metric-value" style="color: #28a745;">{len(open_df)}</div>
        <div class="metric-label">Open</div>
    </div>
    <div class="metric-card pending-metric">
        <div class="metric-value" style="color: #ffc107; text-shadow: 1px 1px 1px rgba(0,0,0,0.3);">{len(pending_df)}</div>
        <div class="metric-label">Pending</div>
    </div>
    <div class="metric-card closed-metric">
        <div class="metric-value" style="color: #6c757d;">{len(closed_df)}</div>
        <div class="metric-label">Closed</div>
    </div>
</div>
""", unsafe_allow_html=True)

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
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
        owners = ["All owners"] + sorted(master_df["Ticket Owner"].dropna().unique().tolist())
        master_owner_filter = st.selectbox("", owners, key="master_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        master_year_filter = st.selectbox("", available_years, key="master_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        master_month_filter = st.selectbox("", available_months, key="master_month_filter", label_visibility="collapsed")
    
    # Display prominent filter results
    display_filter_results(master_df, "Master List", master_owner_filter, master_year_filter, master_month_filter, "#17a2b8")
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, master_owner_filter, master_year_filter, master_month_filter)
    
    # Filter and display data
    master_view = get_filtered_data(master_df, master_owner_filter, master_year_filter, master_month_filter)
    st.dataframe(master_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)

with tab1:
    # Filters for Work in Progress
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
        owners = ["All owners"] + sorted(wip_df["Ticket Owner"].dropna().unique().tolist())
        wip_owner_filter = st.selectbox("", owners, key="wip_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        wip_year_filter = st.selectbox("", available_years, key="wip_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        wip_month_filter = st.selectbox("", available_months, key="wip_month_filter", label_visibility="collapsed")
    
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
        dev_owner_filter = st.selectbox("", owners, key="dev_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        dev_year_filter = st.selectbox("", available_years, key="dev_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        dev_month_filter = st.selectbox("", available_months, key="dev_month_filter", label_visibility="collapsed")
    
    # Display prominent filter results
    display_filter_results(dev_df, "Under Development", dev_owner_filter, dev_year_filter, dev_month_filter, "#fd7e14")
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, dev_owner_filter, dev_year_filter, dev_month_filter)
    
    # Filter and display data
    dev_view = get_filtered_data(dev_df, dev_owner_filter, dev_year_filter, dev_month_filter)
    st.dataframe(dev_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)

with tab3:
    # Filters for Awaiting User Info
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
        owners = ["All owners"] + sorted(wait_df["Ticket Owner"].dropna().unique().tolist())
        wait_owner_filter = st.selectbox("", owners, key="wait_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        wait_year_filter = st.selectbox("", available_years, key="wait_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        wait_month_filter = st.selectbox("", available_months, key="wait_month_filter", label_visibility="collapsed")
    
    # Display prominent filter results
    display_filter_results(wait_df, "Awaiting User Info", wait_owner_filter, wait_year_filter, wait_month_filter, "#dc3545")
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, wait_owner_filter, wait_year_filter, wait_month_filter)
    
    # Filter and display data
    wait_view = get_filtered_data(wait_df, wait_owner_filter, wait_year_filter, wait_month_filter)
    st.dataframe(wait_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)

with tab4:
    # Filters for Hold
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
        owners = ["All owners"] + sorted(hold_df["Ticket Owner"].dropna().unique().tolist())
        hold_owner_filter = st.selectbox("", owners, key="hold_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        hold_year_filter = st.selectbox("", available_years, key="hold_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        hold_month_filter = st.selectbox("", available_months, key="hold_month_filter", label_visibility="collapsed")
    
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
        open_owner_filter = st.selectbox("", owners, key="open_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        open_year_filter = st.selectbox("", available_years, key="open_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        open_month_filter = st.selectbox("", available_months, key="open_month_filter", label_visibility="collapsed")
    
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
        pending_owner_filter = st.selectbox("", owners, key="pending_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        pending_year_filter = st.selectbox("", available_years, key="pending_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        pending_month_filter = st.selectbox("", available_months, key="pending_month_filter", label_visibility="collapsed")
    
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
        closed_owner_filter = st.selectbox("", owners, key="closed_owner_filter", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="filter-label">Filter by Year</div>', unsafe_allow_html=True)
        closed_year_filter = st.selectbox("", available_years, key="closed_year_filter", label_visibility="collapsed")
    with col3:
        st.markdown('<div class="filter-label">Filter by Month</div>', unsafe_allow_html=True)
        closed_month_filter = st.selectbox("", available_months, key="closed_month_filter", label_visibility="collapsed")
    
    # Display prominent filter results
    display_filter_results(closed_df, "Closed", closed_owner_filter, closed_year_filter, closed_month_filter, "#6c757d")
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df, closed_owner_filter, closed_year_filter, closed_month_filter)
    
    # Filter and display data
    closed_view = get_filtered_data(closed_df, closed_owner_filter, closed_year_filter, closed_month_filter)
    st.dataframe(closed_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)

# ---------------------------------------------------------------- AI Assistant & Footer
render_chatbot_assistant(master_df, wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df)
render_footer()

