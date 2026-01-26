"""Streamlit UI component builders"""
import streamlit as st
import pandas as pd


def render_sidebar_template_selection(template_registry):
    """Render template selection in sidebar"""
    st.sidebar.divider()
    st.sidebar.header("📄 Report Template")
    selected_template = st.sidebar.selectbox(
        "Select Report Template",
        list(template_registry.keys())
    )
    template_config = template_registry[selected_template]
    st.sidebar.caption(template_config["description"])
    return selected_template, template_config


def render_sidebar_file_upload():
    """Render file upload in sidebar"""
    st.sidebar.divider()
    st.sidebar.header("📂 Import & Chart Settings")
    uploaded_file = st.sidebar.file_uploader("Upload Excel", type=["xlsx"])
    return uploaded_file


def render_sidebar_chart_settings(df, template_config):
    """Render chart configuration options"""
    numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
    all_cols = df.columns.tolist()

    col_x = st.sidebar.selectbox("X-axis (Group By)", all_cols)
    aggregation = st.sidebar.selectbox(
        "Aggregation", ["None", "Sum", "Mean", "Count"]
    )

    default_chart_type = template_config.get("default_chart_type", "Bar Chart")
    selected_chart_type = st.sidebar.selectbox(
        "Chart Type",
        ["Bar Chart", "Line Chart", "Pie Chart"],
        index=["Bar Chart", "Line Chart", "Pie Chart"].index(default_chart_type)
    )

    default_primary = [c for c in template_config.get("default_primary_y", []) if c in numeric_cols]
    default_secondary = [c for c in template_config.get("default_secondary_y", []) if c in numeric_cols]

    if selected_chart_type == "Pie Chart":
        if not numeric_cols:
            st.sidebar.error("❌ No numeric columns available for pie chart")
            primary_y = []
        else:
            # Safer default index calculation
            default_index = 0
            if default_primary and default_primary[0] in numeric_cols:
                default_index = numeric_cols.index(default_primary[0])
            
            primary_y = [st.sidebar.selectbox(
                "Y-axis",
                numeric_cols,
                index=default_index
            )]
        secondary_y = []
    else:
        primary_y = st.sidebar.multiselect(
            "Primary Y-axis",
            numeric_cols,
            default=default_primary or (numeric_cols[:1] if numeric_cols else [])
        )
        secondary_y = st.sidebar.multiselect(
            "Secondary Y-axis",
            [c for c in numeric_cols if c not in primary_y],
            default=default_secondary
        )

    return col_x, aggregation, selected_chart_type, primary_y, secondary_y


def render_style_settings(template_style):
    """Render current style settings expander"""
    with st.expander("🎨 Current Style Settings"):
        # Show titles if they exist
        if template_style.get("chart_title") or template_style.get("x_axis_title") or template_style.get("y_axis_title"):
            st.write("**Titles:**")
            if template_style.get("chart_title"):
                st.write(f"- Chart: {template_style['chart_title']}")
            if template_style.get("x_axis_title"):
                st.write(f"- X-axis: {template_style['x_axis_title']}")
            if template_style.get("y_axis_title"):
                st.write(f"- Y-axis: {template_style['y_axis_title']}")
            st.divider()
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.write(f"**Font:** {template_style['font_family']} {template_style['font_size']}pt")
            st.write(f"**Grid:** {'On' if template_style['grid'] else 'Off'}")
        with col2:
            st.color_picker("Primary Color", template_style["primary_color"], disabled=True)
            st.color_picker("Axis Line", template_style["axis_line_color"], disabled=True)
        with col3:
            st.color_picker("Secondary Color", template_style.get("secondary_color", "#ff7f0e"), disabled=True)
            st.color_picker("Axis Text", template_style["axis_color"], disabled=True)


def render_ai_chat_section(session_state, agent):
    """Render AI chat assistant section"""
    st.divider()
    st.subheader("🤖 AI Chart Assistant")
    st.caption("Modify the chart using plain English. Changes apply instantly.")

    if agent is None:
        st.warning("AI Assistant is not available. Check OPENAI_API_KEY.")
    else:
        with st.form("chat_form", clear_on_submit=True):
            col1, col2 = st.columns([6, 1])

            with col1:
                user_message = st.text_input(
                    "Type a command (e.g. 'chart title Hello', 'remove grid', 'make font Arial 10')"
                )

            with col2:
                submitted = st.form_submit_button("Send")

            if submitted and user_message.strip():
                return user_message
        
        # Chat history
        if session_state.chat_history:
            with st.expander("💬 Conversation", expanded=False):
                for msg in session_state.chat_history:
                    if msg["role"] == "user":
                        st.markdown(f"**You:** {msg['content']}")
                    else:
                        st.markdown(f"**AI:** {msg['content']}")

        # Controls
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 Reset Style"):
                return "reset_style"
        with c2:
            if st.button("🗑️ Clear Chat"):
                return "clear_chat"
    
    return None