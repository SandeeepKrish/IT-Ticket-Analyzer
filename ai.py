"""
AI-powered chatbot and reply drafting for ticket management system.

Includes both the original reply drafting functionality and a new
data-aware chatbot that can answer questions about uploaded ticket data.
"""

import os
import pandas as pd
import zipfile
import io


def has_key_configured() -> bool:
    return bool(os.environ.get("OPENAI_API_KEY"))


def detect_file_type(filename: str) -> str:
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
    import re
    
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


def process_folder_upload(uploaded_files: list) -> dict:
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
    from openpyxl import load_workbook
    
    result = {
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
        try:
            # Get the file name
            filename = uploaded_file.name if hasattr(uploaded_file, 'name') else str(uploaded_file)
            
            # Detect file type
            file_type = detect_file_type(filename)
            
            if file_type:
                # Read the Excel file
                df = pd.read_excel(uploaded_file, engine="openpyxl")
                df.columns = [str(c).strip() for c in df.columns]
                
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


def validate_folder_upload(folder_result: dict) -> tuple[bool, str]:
    """
    Validate that all required files were detected and loaded.
    
    Returns:
        tuple: (is_valid, message)
    """
    required_keys = ['master', 'wip', 'dev', 'wait', 'hold', 'open', 'pending', 'closed']
    missing_files = [key for key in required_keys if folder_result[key] is None]
    
    if missing_files:
        missing_display = ', '.join([key.upper() for key in missing_files])
        return False, f"Missing files for: {missing_display}. Please ensure all 8 files are included."
    
    return True, "All files loaded successfully!"


def get_folder_upload_summary(folder_result: dict) -> str:
    """Generate a summary of folder upload results"""
    summary = "📂 **Folder Upload Summary:**\n\n"
    
    if folder_result['detected_files']:
        summary += "**✅ Detected Files:**\n"
        for detected in folder_result['detected_files']:
            summary += f"- {detected['original_name']} → {detected['detected_type'].upper()} ({detected['rows']} tickets)\n"
    
    if folder_result['unrecognized_files']:
        summary += "\n**⚠️ Unrecognized Files:**\n"
        for unrecognized in folder_result['unrecognized_files']:
            summary += f"- {unrecognized}\n"
    
    return summary


def draft_reply(row: pd.Series) -> str:
    """Ask OpenAI to draft a short follow-up asking the customer for the
    missing information needed to move this ticket forward."""
    from openai import OpenAI

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("No OpenAI API key configured.")

    client = OpenAI(api_key=api_key)

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
    return response.choices[0].message.content.strip()


def get_chatbot_response(user_question: str, data_context: dict) -> str:
    """Get AI response to user questions about the ticket data"""
    from openai import OpenAI

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("No OpenAI API key configured.")

    client = OpenAI(api_key=api_key)

    # Build context about the data
    context_prompt = f"""You are a helpful assistant for a ticket management system. Answer questions about the uploaded ticket data.

CURRENT DATA SUMMARY:
- Master List: {data_context.get('master_count', 0)} tickets
- Work in Progress: {data_context.get('wip_count', 0)} tickets  
- Under Development: {data_context.get('dev_count', 0)} tickets
- Awaiting User Info: {data_context.get('wait_count', 0)} tickets
- Hold: {data_context.get('hold_count', 0)} tickets
- Open: {data_context.get('open_count', 0)} tickets
- Pending: {data_context.get('pending_count', 0)} tickets
- Closed: {data_context.get('closed_count', 0)} tickets

TOP OWNERS BY TICKET COUNT:
{data_context.get('top_owners', 'No data available')}

WORKFLOW INFORMATION:
- Open → Work in Progress → (Awaiting User Info / Under Development / Hold / Pending) → Closed
- Master List contains all tickets across the system
- Tickets can move between different statuses based on progress

COMMON QUESTIONS I CAN HELP WITH:
- Ticket counts and statistics
- Owner workloads and assignments
- Status explanations and workflow
- Finding specific tickets or information
- Data analysis and trends

Please provide helpful, accurate responses based on this data. If asked about specific ticket details that aren't in the summary, suggest they use the search function or specific tabs.

USER QUESTION: {user_question}"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            max_tokens=500,
            temperature=0.7,
            messages=[
                {
                    "role": "system", 
                    "content": "You are a helpful ticket management assistant. Provide clear, concise answers about ticket data and workflow. Be friendly but professional."
                },
                {"role": "user", "content": context_prompt}
            ],
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"I'm sorry, I encountered an error: {str(e)}. Please make sure your OpenAI API key is configured correctly."


def generate_data_context(master_df, wip_df, dev_df, wait_df, hold_df, open_df, pending_df, closed_df):
    """Generate context about the current data for the chatbot"""
    
    # Get top owners
    all_owners = []
    for df in [wip_df, dev_df, wait_df, hold_df, open_df, pending_df]:
        if df is not None:
            all_owners.extend(df["Ticket Owner"].dropna().tolist())
    
    if all_owners:
        owner_counts = pd.Series(all_owners).value_counts().head(5)
        top_owners = "\n".join([f"- {owner}: {count} active tickets" for owner, count in owner_counts.items()])
    else:
        top_owners = "No active tickets found"
    
    return {
        'master_count': len(master_df) if master_df is not None else 0,
        'wip_count': len(wip_df) if wip_df is not None else 0,
        'dev_count': len(dev_df) if dev_df is not None else 0,
        'wait_count': len(wait_df) if wait_df is not None else 0,
        'hold_count': len(hold_df) if hold_df is not None else 0,
        'open_count': len(open_df) if open_df is not None else 0,
        'pending_count': len(pending_df) if pending_df is not None else 0,
        'closed_count': len(closed_df) if closed_df is not None else 0,
        'top_owners': top_owners
    }
