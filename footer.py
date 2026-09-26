"""
Footer Component
----------------
Renders a stylish, responsive footer at the bottom of the Ticket Flow Tracker UI.
"""

import streamlit as st


def render_footer() -> None:
    """Render the application footer."""
    st.markdown(
        """
        <div class="footer-container">
            <div class="footer-content">
                <p class="footer-text">
                    Copyright @2026 Created by <strong>Sandeep</strong> | <span>Ticket Flow Tracker</span>
                </p>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
