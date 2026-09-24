"""
Optional AI-assisted reply drafting for tickets awaiting user info.

Only used if an OpenAI API key is configured (via the sidebar or a
.env file). Nothing in the rest of the app depends on this — it's purely
a convenience button.
"""

import os

import pandas as pd


def has_key_configured() -> bool:
    return bool(os.environ.get("OPENAI_API_KEY"))


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
