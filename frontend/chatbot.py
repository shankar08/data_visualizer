# src/frontend/chatbot.py
import streamlit as st
import re
from frontend.templates.template import DEFAULT_STYLE

# =====================================================
# APPLY CHATBOT / GPT COMMANDS TO CHART STYLE
# =====================================================
def apply_chatbot_command(prompt: str):
    """
    Updates st.session_state.chart_style_updates and chart titles based on user command.
    GPT/NLP-ready: you can extend this to call LLM for advanced parsing.
    """
    updates = st.session_state.get("chart_style_updates", {}).copy()
    prompt_lower = prompt.lower()

    # ------------------
    # Colors
    # ------------------
    color_map = {
        "red": "#FF0000",
        "blue": "#0000FF",
        "green": "#00FF00",
        "orange": "#FF7F0E",
    }
    for key, val in color_map.items():
        if key in prompt_lower:
            updates["primary_color"] = val
    if "secondary red" in prompt_lower:
        updates["secondary_color"] = "#FF0000"
    if "secondary blue" in prompt_lower:
        updates["secondary_color"] = "#0000FF"

    # ------------------
    # Grid
    # ------------------
    if "remove grid" in prompt_lower or "no grid" in prompt_lower:
        st.session_state.grid = False
    if "show grid" in prompt_lower:
        st.session_state.grid = True

    # ------------------
    # Titles
    # ------------------
    if "chart title" in prompt_lower:
        title_text = prompt.split("chart title")[-1].strip()
        st.session_state.chart_title = title_text
        updates["chart_title"] = title_text

    if "x axis" in prompt_lower:
        st.session_state.x_axis_title = prompt.split("x axis")[-1].strip()

    if "y axis" in prompt_lower:
        st.session_state.y_axis_title = prompt.split("y axis")[-1].strip()

    # ------------------
    # Font size
    # ------------------
    size_match = re.search(r'font size (\d+)', prompt_lower)
    if size_match:
        updates["font_size"] = int(size_match.group(1))

    st.session_state.chart_style_updates = updates
    return updates