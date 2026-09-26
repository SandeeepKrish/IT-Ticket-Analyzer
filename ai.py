"""
AI-powered chatbot and reply drafting for ticket management system.

Includes both the original reply drafting functionality and a new
data-aware chatbot that can answer questions about uploaded ticket data.
"""

import os
import re
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd

try:
    from openai import OpenAI
    HAS_OPENAI = True
except ImportError:
    OpenAI = None  # type: ignore[assignment,misc]
    HAS_OPENAI = False


def has_key_configured() -> bool:
    return bool(os.environ.get("OPENAI_API_KEY"))


def detect_file_type(filename: str) -> Optional[str]:
    """
    Detect file type based on filename using fuzzy matching.
    Ignores case, spaces, underscores, and hyphens for flexible matching.
    
    Mapping logic:
    - "master", "masterticket" → master
    - "wip", "work", "workinprogress" → wip
    - "dev", "development", "underdevelopment" → dev
    - "await", "awaiting", "awaiting_user_info", "user_info" → wait
    - "hold" → hold
    - "open" → open
    - "pending" → pending
    - "closed", "close" → closed
    """
    # Normalize filename: lowercase, remove spaces/underscores/hyphens, keep only alphanumeric
    normalized = re.sub(r'[^a-z0-9]', '', filename.lower())
    
    # Define mapping with keywords to match
    # Order matters: check longer/more specific patterns first
    mapping = [
        (['masterticket', 'master'], 'master'),
        (['workinprogress', 'workprogress', 'wip', 'work'], 'wip'),  # Matches various WIP patterns
        (['underdevelopment', 'development', 'dev'], 'dev'),  # dev comes after longer patterns
        (['awaiting', 'await', 'userinfo'], 'wait'),  # Matches "await", "awaiting", "awaiting_user_info", etc.
        (['hold'], 'hold'),
        (['open'], 'open'),
        (['pending'], 'pending'),
        (['closed', 'close'], 'closed'),  # Matches both "closed" and "close"
    ]
    
    # Check each category
    for keywords, category in mapping:
        for keyword in keywords:
            if keyword in normalized:
                return category
    
    return None


def process_folder_upload(uploaded_files: List[Any]) -> Dict[str, Any]:
    """
    Process multiple uploaded files and map them to their categories.
    
    Expected files:
    - Master or Master Ticket
    - Work in Progress or WIP
    - Under Development or Development
    - Awaiting User Info or Awaiting
    - Hold
    - Open
    - Pending
    - Closed
    
    Returns:
        dict: Mapped dataframes with keys: master, wip, dev, wait, hold, open, pending, closed
    """
    result: Dict[str, Any] = {
        'master': None,
        'wip': None,
        'dev': None,
        'wait': None,
        'hold': None,
        'open': None,
        'pending': None,
        'closed': None,
        'detected_files': [],
        'unrecognized_files': []
    }
    
    if not uploaded_files:
        return result
    
    for uploaded_file in uploaded_files:
        filename = ""
        try:
            # Get the file name
            filename = uploaded_file.name if hasattr(uploaded_file, 'name') else str(uploaded_file)
            
            # Detect file type
            file_type = detect_file_type(filename)
            
            if file_type:
                # Read the Excel file
                df = pd.read_excel(uploaded_file, engine="openpyxl")
                df.columns = [str(c).strip() for c in df.columns]
                if "Ticket Number" in df.columns:
                    df["Ticket Number"] = df["Ticket Number"].astype(str).str.strip()
                if "Ticket Aging" in df.columns:
                    df["Ticket Aging"] = pd.to_numeric(df["Ticket Aging"], errors="coerce").fillna(0).astype(int)
                
                # Store the dataframe
                result[file_type] = df
                result['detected_files'].append({
                    'original_name': filename,
                    'detected_type': file_type,
                    'rows': len(df)
                })
            else:
                result['unrecognized_files'].append(filename)
                
        except Exception as e:
            result['unrecognized_files'].append(f"{filename} (Error: {str(e)})")
    
    return result


def validate_folder_upload(folder_result: Dict[str, Any]) -> Tuple[bool, str]:
    """
    Validate that all required files were detected and loaded.
    
    Returns:
        tuple: (is_valid, message)
    """
    required_keys = ['master', 'wip', 'dev', 'wait', 'hold', 'open', 'pending', 'closed']
    missing_files = [key for key in required_keys if folder_result.get(key) is None]
    
    if missing_files:
        missing_display = ', '.join([key.upper() for key in missing_files])
        return False, f"Missing files for: {missing_display}. Please ensure all 8 files are included."
    
    return True, "All files loaded successfully!"


def get_folder_upload_summary(folder_result: Dict[str, Any]) -> str:
    """Generate a summary of folder upload results"""
    summary = "📂 **Folder Upload Summary:**\n\n"
    
    if folder_result.get('detected_files'):
        summary += "**✅ Detected Files:**\n"
        for detected in folder_result['detected_files']:
            summary += f"- {detected['original_name']} → {detected['detected_type'].upper()} ({detected['rows']} tickets)\n"
    
    if folder_result.get('unrecognized_files'):
        summary += "\n**⚠️ Unrecognized Files:**\n"
        for unrecognized in folder_result['unrecognized_files']:
            summary += f"- {unrecognized}\n"
    
    return summary


def draft_reply(row: pd.Series) -> str:
    """Ask OpenAI to draft a short follow-up asking the customer for the
    missing information needed to move this ticket forward."""
    if OpenAI is None:
        raise RuntimeError("OpenAI package is not installed. Please install it using `pip install openai`.")

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("No OpenAI API key configured.")

    client: Any = OpenAI(api_key=api_key)

    description = str(row.get("Description", "") or "").strip()

    prompt = f"""Write a short, polite follow-up message to a customer asking them to
provide the information needed to move their support ticket forward. It's
currently marked "Awaiting User Info" in our system.

Ticket number: {row.get('Ticket Number', '')}
Subject: {row.get('Subject', '')}
Customer: {row.get('Customer Name', '')}
Ticket description: {description if description else '(no description provided)'}

Keep it under 120 words, professional and specific about what's still
needed from them. Do not invent details that aren't in the description —
if it's unclear exactly what's missing, ask them to confirm the details
needed to proceed. Sign off simply as "Support Team"."""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        max_tokens=300,
        messages=[{"role": "user", "content": prompt}],
    )
    content = response.choices[0].message.content
    return (content or "").strip()


def search_data_for_question(user_question: str, data_context: Dict[str, Any]) -> str:
    """Search uploaded Excel data for specific keywords, ticket numbers, or owner names."""
    raw_dfs = data_context.get('raw_dfs', {})
    if not raw_dfs:
        return ""
    
    found_info: List[str] = []
    
    # Split question into clean search tokens
    tokens = [t.strip() for t in re.split(r'[\s,;:?!\'"()]+', user_question) if len(t.strip()) > 1]
    
    # 1. Search for matching ticket number
    for category, df in raw_dfs.items():
        if df is None or df.empty or "Ticket Number" not in df.columns:
            continue
        for token in tokens:
            matches = df[df["Ticket Number"].astype(str).str.upper() == token.upper()]
            if not matches.empty:
                for _, row in matches.head(3).iterrows():
                    found_info.append(
                        f"Found matching ticket '{row['Ticket Number']}' in {category}:\n"
                        f"  Subject: {row.get('Subject', 'N/A')}\n"
                        f"  Owner: {row.get('Ticket Owner', 'N/A')}\n"
                        f"  Customer: {row.get('Customer Name', 'N/A')}\n"
                        f"  Status: {row.get('Status', category)}\n"
                        f"  Priority: {row.get('Ticket Priority', 'N/A')}\n"
                        f"  Aging: {row.get('Ticket Aging', 'N/A')} days\n"
                        f"  Start Date: {row.get('Start Date', 'N/A')}\n"
                        f"  Description: {str(row.get('Description', 'N/A'))[:200]}"
                    )
    
    # 2. Search for matching owner name
    for category, df in raw_dfs.items():
        if df is None or df.empty or "Ticket Owner" not in df.columns:
            continue
        for token in tokens:
            if len(token) >= 3 and token.lower() not in ["show", "list", "tickets", "what", "how", "many", "tell", "about", "with", "from"]:
                owner_matches = df[df["Ticket Owner"].astype(str).str.lower().str.contains(token.lower(), na=False)]
                if not owner_matches.empty:
                    count = len(owner_matches)
                    sample_tickets = owner_matches["Ticket Number"].astype(str).head(5).tolist()
                    found_info.append(
                        f"Owner match for '{token}' in {category}: {count} ticket(s). Sample Ticket Numbers: {', '.join(sample_tickets)}"
                    )
    
    if found_info:
        unique_info = list(dict.fromkeys(found_info))
        return "\n\nDETAILED MATCHES FOUND IN UPLOADED EXCEL DATA:\n" + "\n".join(unique_info[:10])
    return ""


def generate_data_context(
    master_df: Optional[pd.DataFrame] = None,
    wip_df: Optional[pd.DataFrame] = None,
    dev_df: Optional[pd.DataFrame] = None,
    wait_df: Optional[pd.DataFrame] = None,
    hold_df: Optional[pd.DataFrame] = None,
    open_df: Optional[pd.DataFrame] = None,
    pending_df: Optional[pd.DataFrame] = None,
    closed_df: Optional[pd.DataFrame] = None,
) -> Dict[str, Any]:
    """Generate comprehensive context about the current data for the chatbot"""
    df_map = {
        'Master': master_df,
        'Work in Progress (WIP)': wip_df,
        'Under Development (DEV)': dev_df,
        'Awaiting User Info (WAIT)': wait_df,
        'Hold': hold_df,
        'Open': open_df,
        'Pending': pending_df,
        'Closed': closed_df,
    }
    
    counts = {name: (len(df) if df is not None and not df.empty else 0) for name, df in df_map.items()}
    
    all_owners: List[str] = []
    priorities: List[str] = []
    aging_stats: List[str] = []
    
    for name, df in df_map.items():
        if df is not None and not df.empty:
            if "Ticket Owner" in df.columns:
                all_owners.extend(df["Ticket Owner"].dropna().astype(str).tolist())
            if "Ticket Priority" in df.columns:
                priorities.extend(df["Ticket Priority"].dropna().astype(str).tolist())
            if "Ticket Aging" in df.columns and len(df) > 0:
                mean_a = df["Ticket Aging"].mean()
                max_a = df["Ticket Aging"].max()
                aging_stats.append(f"- {name}: Avg Aging {mean_a:.1f} days (Max: {max_a} days)")

    owner_summary = "No owner data"
    if all_owners:
        top_owners = pd.Series(all_owners).value_counts().head(10)
        owner_summary = "\n".join([f"- {owner}: {count} tickets" for owner, count in top_owners.items()])
        
    prio_summary = "No priority data"
    if priorities:
        prio_counts = pd.Series(priorities).value_counts()
        prio_summary = "\n".join([f"- {prio}: {count} tickets" for prio, count in prio_counts.items()])

    aging_summary = "\n".join(aging_stats) if aging_stats else "No aging data available"
    
    return {
        'counts': counts,
        'master_count': counts['Master'],
        'wip_count': counts['Work in Progress (WIP)'],
        'dev_count': counts['Under Development (DEV)'],
        'wait_count': counts['Awaiting User Info (WAIT)'],
        'hold_count': counts['Hold'],
        'open_count': counts['Open'],
        'pending_count': counts['Pending'],
        'closed_count': counts['Closed'],
        'top_owners': owner_summary,
        'priority_summary': prio_summary,
        'aging_summary': aging_summary,
        'raw_dfs': df_map
    }


def get_chatbot_response(
    user_question: str, 
    data_context: Dict[str, Any],
    history: Optional[List[Dict[str, str]]] = None
) -> str:
    """Get AI chatbot response for user questions about uploaded ticket Excel data."""
    if OpenAI is None:
        return "The OpenAI package is not installed. Please install `openai` to use the chatbot."

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        return "No OpenAI API key configured. Please enter your OpenAI API key in the sidebar or `.env` file."

    client: Any = OpenAI(api_key=api_key)

    specific_data_matches = search_data_for_question(user_question, data_context)

    counts_str = "\n".join([f"- {k}: {v} tickets" for k, v in data_context.get('counts', {}).items()])
    
    system_prompt = (
        "You are an expert AI Data Assistant for the Ticket Flow Tracker system. "
        "Your primary job is to answer user questions about the uploaded Excel ticket data with extreme precision and helpfulness.\n\n"
        "UPLOADED EXCEL FILE OVERVIEW:\n"
        f"{counts_str}\n\n"
        "TOP OWNER WORKLOAD:\n"
        f"{data_context.get('top_owners', 'N/A')}\n\n"
        "PRIORITY DISTRIBUTION:\n"
        f"{data_context.get('priority_summary', 'N/A')}\n\n"
        "AGING METRICS:\n"
        f"{data_context.get('aging_summary', 'N/A')}\n"
        f"{specific_data_matches}\n\n"
        "GUIDELINES:\n"
        "- Direct, concise, accurate answer based on the Excel file summary and matches above.\n"
        "- Use markdown formatting (bolding, bullet points, headers) for clean visual layout.\n"
        "- If ticket numbers, owners, or specific statuses are mentioned, cite exact numbers and details."
    )

    messages: List[Dict[str, Any]] = [{"role": "system", "content": system_prompt}]
    
    if history:
        for msg in history[-6:]:
            if msg.get("role") in ("user", "assistant") and msg.get("content"):
                messages.append({"role": msg["role"], "content": msg["content"]})
    
    messages.append({"role": "user", "content": user_question})

    try:
        response: Any = client.chat.completions.create(
            model="gpt-4o-mini",
            max_tokens=700,
            temperature=0.3,
            messages=messages,
        )
        content = response.choices[0].message.content
        return (content or "").strip()
    except Exception as e:
        return f"I encountered an error analyzing your data: {str(e)}. Please check your OpenAI API key configuration."


