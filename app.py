"""
Ticket Flow Tracker
--------------------
Upload your "Under Development" and "Awaiting User Info" ticket exports
(.xlsx) and search a ticket number to see its status and every other open
ticket for the same owner — so nothing sits stuck waiting on a reply.

Run with:  streamlit run app.py
"""

import os

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from ai import draft_reply, has_key_configured

load_dotenv()

st.set_page_config(
    page_title="Ticket Flow Tracker", 
    page_icon="🎫", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for professional styling
st.markdown("""
<style>
    /* Main page styling */
    .stApp {
        background: #ffffff;
        min-height: 100vh;
    }
    
    /* Header styling */
    .main-header {
        padding: 1.5rem 0;
        margin-bottom: 2rem;
        border-bottom: 1px solid #e9ecef;
        background: transparent;
    }
    
    .main-title {
        color: #2c3e50;
        font-size: 1.8rem;
        font-weight: 600;
        margin-bottom: 0.3rem;
        text-align: left;
    }
    
    .main-subtitle {
        color: #6c757d;
        font-size: 0.95rem;
        text-align: left;
        margin-bottom: 0;
    }
    
    /* Metric cards */
    .metric-container {
        display: flex;
        gap: 1rem;
        margin: 2rem 0;
        justify-content: center;
    }
    
    .metric-card {
        background: #f8f9fa;
        padding: 1.2rem;
        border-radius: 8px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border: 1px solid #dee2e6;
        min-width: 160px;
    }
    
    .metric-value {
        font-size: 2rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    
    .metric-label {
        color: #6c757d;
        font-size: 0.85rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-weight: 500;
    }
    
    .wip-metric { border-left: 3px solid #007bff; }
    .dev-metric { border-left: 3px solid #fd7e14; }
    .wait-metric { border-left: 3px solid #dc3545; }
    
    /* Status indicators */
    .status-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 1rem;
        margin: 1.5rem 0;
    }
    
    .status-card {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border: 1px solid #dee2e6;
    }
    
    .status-found {
        border-left: 3px solid #28a745;
        background: #d4edda;
        color: #155724;
    }
    
    .status-not-found {
        border-left: 3px solid #dc3545;
        background: #f8d7da;
        color: #721c24;
    }
    
    /* Ticket card styling */
    .ticket-card {
        background: #ffffff;
        padding: 1.5rem;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        border: 1px solid #dee2e6;
        margin: 1rem 0;
    }
    
    .ticket-header {
        color: #2c3e50;
        font-size: 1.3rem;
        font-weight: 600;
        margin-bottom: 1rem;
        padding-bottom: 0.8rem;
        border-bottom: 2px solid #e9ecef;
    }
    
    .ticket-info-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 0.8rem;
        margin-bottom: 1rem;
    }
    
    .ticket-info-item {
        text-align: center;
        padding: 0.8rem;
        background: #f8f9fa;
        border-radius: 6px;
        border: 1px solid #e9ecef;
    }
    
    .ticket-info-label {
        font-weight: 500;
        color: #495057;
        font-size: 0.85rem;
        margin-bottom: 0.4rem;
    }
    
    .ticket-info-value {
        color: #212529;
        font-size: 0.95rem;
    }
    
    /* Search styling - Remove white container */
    .search-section {
        margin: 1.5rem 0;
    }
    
    .search-title {
        color: #2c3e50;
        font-size: 1.3rem;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    
    /* Sidebar styling */
    .sidebar-header {
        color: #2c3e50;
        font-size: 1.1rem;
        font-weight: 600;
        margin-bottom: 1rem;
        text-align: left;
    }
    
    /* Related tickets styling */
    .related-section {
        background: #f8f9fa;
        padding: 1.2rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 1rem;
        border: 1px solid #dee2e6;
    }
    
    .related-header {
        color: #2c3e50;
        font-weight: 600;
        font-size: 1rem;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #dee2e6;
    }
    
    /* Hide streamlit branding and margins */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stApp > header {display: none;}
    
    /* Custom button styling */
    .stButton > button {
        background: #007bff;
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.4rem 0.8rem;
        font-weight: 500;
        font-size: 0.9rem;
        transition: all 0.2s ease;
    }
    
    .stButton > button:hover {
        background: #0056b3;
        transform: translateY(-1px);
        box-shadow: 0 2px 8px rgba(0, 123, 255, 0.3);
    }
    
    /* Info boxes styling */
    .stAlert {
        border-radius: 6px;
        border: 1px solid #dee2e6;
    }
    
    /* Enhanced filter section styling */
    .filter-section {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 8px;
        margin: 1.5rem 0;
        border: 1px solid #dee2e6;
    }
    
    .filter-title {
        color: #2c3e50;
        font-size: 1.2rem;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    
    /* User statistics cards */
    .user-stats-container {
        display: flex;
        gap: 1rem;
        margin: 1rem 0;
        justify-content: center;
        flex-wrap: wrap;
    }
    
    .user-stat-card {
        background: #ffffff;
        padding: 1rem;
        border-radius: 8px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border: 1px solid #dee2e6;
        min-width: 140px;
        position: relative;
        overflow: hidden;
    }
    
    .user-stat-card::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 3px;
        background: var(--border-color);
    }
    
    .stat-number {
        font-size: 1.8rem;
        font-weight: 600;
        margin-bottom: 0.3rem;
        color: var(--number-color);
    }
    
    .stat-label {
        color: #6c757d;
        font-size: 0.8rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-weight: 500;
    }
    
    /* Spinner animation for numbers */
    @keyframes countUp {
        from { transform: translateY(20px); opacity: 0; }
        to { transform: translateY(0); opacity: 1; }
    }
    
    .animate-count {
        animation: countUp 0.5s ease-out;
    }
    
    /* Tab styling improvements */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 6px;
        padding: 0.5rem 1rem;
        font-weight: 500;
    }
    
    .stTabs [aria-selected="true"] {
        background: #007bff;
        color: white !important;
        border-color: #007bff;
    }
    
    /* Owner filter styling */
    .owner-filter-container {
        background: #ffffff;
        padding: 1rem;
        border-radius: 6px;
        border: 1px solid #dee2e6;
        margin-bottom: 1rem;
    }
    
    .filter-label {
        color: #495057;
        font-weight: 500;
        font-size: 0.9rem;
        margin-bottom: 0.5rem;
    }
    /* Welcome section styling */
    .welcome-section {
        background: #f8f9fa;
        padding: 2rem;
        border-radius: 8px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin: 2rem 0;
        border: 1px solid #dee2e6;
    }
    
    .welcome-title {
        color: #007bff;
        font-size: 1.5rem;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    
    .welcome-text {
        color: #6c757d;
        font-size: 1rem;
        line-height: 1.4;
    }
</style>
""", unsafe_allow_html=True)

def display_user_statistics(wip_df: pd.DataFrame, dev_df: pd.DataFrame, wait_df: pd.DataFrame, selected_owner: str):
    """Display statistics for the selected owner with animated counters"""
    if selected_owner == "All owners":
        return
    
    # Calculate statistics for the selected owner
    wip_count = len(wip_df[wip_df["Ticket Owner"] == selected_owner])
    dev_count = len(dev_df[dev_df["Ticket Owner"] == selected_owner])
    wait_count = len(wait_df[wait_df["Ticket Owner"] == selected_owner])
    total_count = wip_count + dev_count + wait_count
    
    # Display user statistics with animated counters
    st.markdown(f"""
    <div class="filter-section">
        <div class="filter-title">📊 {selected_owner}'s Ticket Statistics</div>
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
            <div class="user-stat-card" style="--border-color: #28a745; --number-color: #28a745;">
                <div class="stat-number animate-count">{total_count}</div>
                <div class="stat-label">Total Tickets</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def get_filtered_data(df: pd.DataFrame, owner_filter: str):
    """Filter dataframe by owner and return the filtered data"""
    if owner_filter == "All owners":
        return df
    return df[df["Ticket Owner"] == owner_filter]


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
    
    st.markdown(f'<div class="related-section">', unsafe_allow_html=True)
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


# ---------------------------------------------------------------- Sidebar
with st.sidebar:
    st.markdown('<div class="sidebar-header">📂 Upload Reports</div>', unsafe_allow_html=True)
    
    st.markdown("### 📊 Work in Progress")
    wip_file = st.file_uploader("Work in Progress export (.xlsx)", type="xlsx", key="wip")
    
    st.markdown("### 🔧 Under Development")
    dev_file = st.file_uploader("Under Development export (.xlsx)", type="xlsx", key="dev")
    
    st.markdown("### ⏳ Awaiting User Info")
    wait_file = st.file_uploader("Awaiting User Info export (.xlsx)", type="xlsx", key="wait")

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

if not wip_file or not dev_file or not wait_file:
    st.markdown("""
    <div class="welcome-section">
        <div class="welcome-title">🚀 Get Started</div>
        <div class="welcome-text">Upload all three Excel exports in the sidebar to begin tracking your tickets.</div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

wip_df = load_excel(wip_file)
dev_df = load_excel(dev_file)
wait_df = load_excel(wait_file)

if not (validate_columns(wip_df, "Work in Progress") and 
        validate_columns(dev_df, "Under Development") and 
        validate_columns(wait_df, "Awaiting User Info")):
    st.stop()

# Metrics display
st.markdown("""
<div class="metric-container">
    <div class="metric-card wip-metric">
        <div class="metric-value" style="color: #007bff;">{}</div>
        <div class="metric-label">Work in Progress</div>
    </div>
    <div class="metric-card dev-metric">
        <div class="metric-value" style="color: #fd7e14;">{}</div>
        <div class="metric-label">Under Development</div>
    </div>
    <div class="metric-card wait-metric">
        <div class="metric-value" style="color: #dc3545;">{}</div>
        <div class="metric-label">Awaiting User Info</div>
    </div>
</div>
""".format(len(wip_df), len(dev_df), len(wait_df)), unsafe_allow_html=True)

st.divider()

# Search section - no container, just clean styling
st.markdown('<div class="search-section">', unsafe_allow_html=True)
st.markdown('<div class="search-title">🔍 Search a Ticket</div>', unsafe_allow_html=True)
ticket_no = st.text_input("Enter ticket number", placeholder="e.g. CT007948", key="ticket_search", label_visibility="collapsed").strip().upper()
st.markdown('</div>', unsafe_allow_html=True)

if ticket_no:
    wip_match = wip_df[wip_df["Ticket Number"].str.upper() == ticket_no]
    dev_match = dev_df[dev_df["Ticket Number"].str.upper() == ticket_no]
    wait_match = wait_df[wait_df["Ticket Number"].str.upper() == ticket_no]

    # Status indicators
    st.markdown("### 📊 Ticket Status Overview")
    
    wip_found = not wip_match.empty
    dev_found = not dev_match.empty
    wait_found = not wait_match.empty
    
    st.markdown(f"""
    <div class="status-grid">
        <div class="status-card {'status-found' if wip_found else 'status-not-found'}">
            <h4 style="margin-bottom: 0.5rem;">{'✅' if wip_found else '❌'} Work in Progress</h4>
            <p style="margin: 0;">{'Ticket Found' if wip_found else 'Not Found'}</p>
        </div>
        <div class="status-card {'status-found' if dev_found else 'status-not-found'}">
            <h4 style="margin-bottom: 0.5rem;">{'✅' if dev_found else '❌'} Under Development</h4>
            <p style="margin: 0;">{'Ticket Found' if dev_found else 'Not Found'}</p>
        </div>
        <div class="status-card {'status-found' if wait_found else 'status-not-found'}">
            <h4 style="margin-bottom: 0.5rem;">{'✅' if wait_found else '❌'} Awaiting User Info</h4>
            <p style="margin: 0;">{'Ticket Found' if wait_found else 'Not Found'}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Primary status determination and display
    row, bucket = None, None
    if wait_found:
        st.info(f"🎯 **Primary Status**: **{ticket_no}** is currently in **Awaiting User Info** - This needs your reply!")
        row, bucket = wait_match.iloc[0], "wait"
    elif dev_found:
        st.info(f"🎯 **Primary Status**: **{ticket_no}** is currently in **Under Development** - Active work ongoing.")
        row, bucket = dev_match.iloc[0], "dev"
    elif wip_found:
        st.info(f"🎯 **Primary Status**: **{ticket_no}** is in **Work in Progress** - Initial phase.")
        row, bucket = wip_match.iloc[0], "wip"
    else:
        st.error(f'🚫 **No Results**: Ticket "{ticket_no}" was not found in any of the three sheets.')

    if row is not None:
        st.divider()
        st.markdown("### 🎫 Ticket Details")
        ticket_card(row)

        owner = row["Ticket Owner"]
        st.divider()
        st.markdown(f"### 👤 {owner}'s Other Open Tickets")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            related_list(wip_df, owner, row["Ticket Number"], 
                        "No other tickets in work in progress for this owner.", 
                        False, "📊 Work in Progress")
        with col2:
            related_list(dev_df, owner, row["Ticket Number"], 
                        "No other tickets in development for this owner.", 
                        False, "🔧 Under Development")
        with col3:
            related_list(wait_df, owner, row["Ticket Number"], 
                        "No other tickets awaiting info for this owner.", 
                        True, "⏳ Waiting on Reply")

st.divider()

# Enhanced filtering section
st.markdown("### 🔍 Filter by Owner")
all_owners = set()
all_owners.update(wip_df["Ticket Owner"].dropna().unique())
all_owners.update(dev_df["Ticket Owner"].dropna().unique())
all_owners.update(wait_df["Ticket Owner"].dropna().unique())
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
    total_user_tickets = wip_user_count + dev_user_count + wait_user_count
    
    # Calculate percentages
    wip_percentage = (wip_user_count / total_user_tickets * 100) if total_user_tickets > 0 else 0
    dev_percentage = (dev_user_count / total_user_tickets * 100) if total_user_tickets > 0 else 0
    wait_percentage = (wait_user_count / total_user_tickets * 100) if total_user_tickets > 0 else 0
    
    st.markdown(f"""
    <div class="filter-section">
        <div class="filter-title">📊 Complete Statistics for {selected_owner}</div>
        <div class="user-stats-container">
            <div class="user-stat-card" style="--border-color: #007bff; --number-color: #007bff;">
                <div class="stat-number animate-count">{wip_user_count}</div>
                <div class="stat-label">Work in Progress</div>
                <div style="font-size: 0.7rem; color: #6c757d; margin-top: 0.2rem;">{wip_percentage:.1f}% of total</div>
            </div>
            <div class="user-stat-card" style="--border-color: #fd7e14; --number-color: #fd7e14;">
                <div class="stat-number animate-count">{dev_user_count}</div>
                <div class="stat-label">Under Development</div>
                <div style="font-size: 0.7rem; color: #6c757d; margin-top: 0.2rem;">{dev_percentage:.1f}% of total</div>
            </div>
            <div class="user-stat-card" style="--border-color: #dc3545; --number-color: #dc3545;">
                <div class="stat-number animate-count">{wait_user_count}</div>
                <div class="stat-label">Awaiting User Info</div>
                <div style="font-size: 0.7rem; color: #6c757d; margin-top: 0.2rem;">{wait_percentage:.1f}% of total</div>
            </div>
            <div class="user-stat-card" style="--border-color: #28a745; --number-color: #28a745;">
                <div class="stat-number animate-count">{total_user_tickets}</div>
                <div class="stat-label">Total Active Tickets</div>
                <div style="font-size: 0.7rem; color: #6c757d; margin-top: 0.2rem;">100% workload</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    if total_user_tickets > 0:
        # Show workload analysis
        if wait_user_count > 0:
            st.warning(f"⚠️ **Action Needed**: {selected_owner} has {wait_user_count} ticket(s) awaiting user information that need attention.")
        if dev_user_count >= 5:
            st.info(f"💼 **High Workload**: {selected_owner} is actively working on {dev_user_count} tickets in development.")
        if total_user_tickets >= 10:
            st.error(f"🚨 **Overloaded**: {selected_owner} has {total_user_tickets} total active tickets - consider redistributing workload.")

st.divider()

tab1, tab2, tab3 = st.tabs(["📋 Work in Progress", "📋 Under Development", "📋 Awaiting User Info"])

with tab1:
    st.markdown('<div class="owner-filter-container">', unsafe_allow_html=True)
    st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
    owners = ["All owners"] + sorted(wip_df["Ticket Owner"].dropna().unique().tolist())
    wip_owner_filter = st.selectbox("", owners, key="wip_owner_filter", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, wip_owner_filter)
    
    # Filter and display data
    wip_view = get_filtered_data(wip_df, wip_owner_filter)
    st.dataframe(wip_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)

with tab2:
    st.markdown('<div class="owner-filter-container">', unsafe_allow_html=True)
    st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
    owners = ["All owners"] + sorted(dev_df["Ticket Owner"].dropna().unique().tolist())
    dev_owner_filter = st.selectbox("", owners, key="dev_owner_filter", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, dev_owner_filter)
    
    # Filter and display data
    dev_view = get_filtered_data(dev_df, dev_owner_filter)
    st.dataframe(dev_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)

with tab3:
    st.markdown('<div class="owner-filter-container">', unsafe_allow_html=True)
    st.markdown('<div class="filter-label">Filter by Owner</div>', unsafe_allow_html=True)
    owners = ["All owners"] + sorted(wait_df["Ticket Owner"].dropna().unique().tolist())
    wait_owner_filter = st.selectbox("", owners, key="wait_owner_filter", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Display user statistics if specific owner is selected
    display_user_statistics(wip_df, dev_df, wait_df, wait_owner_filter)
    
    # Filter and display data
    wait_view = get_filtered_data(wait_df, wait_owner_filter)
    st.dataframe(wait_view.sort_values("Ticket Aging", ascending=False), use_container_width=True, hide_index=True)
