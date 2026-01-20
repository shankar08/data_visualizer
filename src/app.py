"""
Main Streamlit application for Excel Chart Generator with AI Assistant
"""
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from pathlib import Path
import sys

# Add current directory to path
APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

# Import modules
from src.config import OPENAI_API_KEY, ROBOTO_PATH
from src.chatbot.chart_modifier import ChartModifierAgent
from src.chatbot.style_updater import StyleUpdater
from src.templates.template import TEMPLATE_REGISTRY
from src.chart_formatter import apply_vba_formatting
from src.excel_exporter import create_excel_report
from src.ui_components import (
    render_sidebar_template_selection,
    render_sidebar_file_upload,
    render_sidebar_chart_settings,
    render_style_settings,
    render_ai_chat_section
)


def setup_fonts():
    """Setup custom fonts for matplotlib"""
    if ROBOTO_PATH.exists():
        fm.fontManager.addfont(str(ROBOTO_PATH))
        plt.rcParams["font.family"] = "Roboto"


def process_dataframe(df, col_x, aggregation, numeric_cols):
    """Process dataframe with aggregation if needed"""
    if aggregation != "None":
        if aggregation == "Sum":
            plot_df = df.groupby(col_x)[numeric_cols].sum().reset_index()
        elif aggregation == "Mean":
            plot_df = df.groupby(col_x)[numeric_cols].mean().reset_index()
        elif aggregation == "Count":
            plot_df = df.groupby(col_x)[numeric_cols].count().reset_index()
    else:
        plot_df = df.copy()
    
    # Remove NaN values that might occur during aggregation
    plot_df = plot_df.dropna(subset=numeric_cols, how='all')
    return plot_df


def render_chart_preview(plot_df, col_x, primary_y, secondary_y, chart_type, template_style):
    """Render matplotlib chart preview"""
    st.subheader("👀 Chart Preview")

    plt.rcParams.update({"font.size": template_style["font_size"]})

    fig, ax1 = plt.subplots(figsize=(8, 4))
    
    # Axis titles
    if template_style.get("x_axis_title"):
        ax1.set_xlabel(
            template_style["x_axis_title"],
            fontsize=template_style["font_size"],
            color=template_style["axis_color"]
        )
    else:
        ax1.set_xlabel("")

    if template_style.get("y_axis_title"):
        ax1.set_ylabel(
            template_style["y_axis_title"],
            fontsize=template_style["font_size"],
            color=template_style["axis_color"]
        )
    else:
        ax1.set_ylabel("")
    
    # Spine styling
    for spine in ["top", "right"]:
        ax1.spines[spine].set_visible(False)
    for spine in ["bottom", "left"]:
        ax1.spines[spine].set_color(template_style["axis_line_color"])
        ax1.spines[spine].set_linewidth(0.5)
    ax1.tick_params(axis="both", colors=template_style["axis_color"], direction="out")

    if template_style.get("chart_title"):
        ax1.set_title(
            template_style["chart_title"],
            fontsize=template_style["font_size"] + 2,
            color=template_style["axis_color"]
        )

    ax2 = ax1.twinx() if secondary_y else None

    # Plot data
    if chart_type in ["Bar Chart", "Line Chart"]:
        for col in primary_y:
            if chart_type == "Bar Chart":
                ax1.bar(plot_df[col_x], plot_df[col],
                        label=col,
                        color=template_style["primary_color"],
                        alpha=0.85,
                        width=0.5)
            else:
                ax1.plot(plot_df[col_x], plot_df[col],
                         label=col,
                         linewidth=1.5,
                         color=template_style["primary_color"])

        if ax2:
            ax2.spines["top"].set_visible(False)
            ax2.spines["right"].set_color(template_style["axis_line_color"])
            ax2.spines["right"].set_linewidth(0.5)
            ax2.tick_params(colors=template_style["axis_color"], direction="out")
            for col in secondary_y:
                ax2.plot(plot_df[col_x], plot_df[col],
                         linestyle="--",
                         linewidth=1.2,
                         label=col,
                         color=template_style.get("secondary_color", "#ff7f0e"))

        ax1.legend(loc="upper left", frameon=False, fontsize=template_style["font_size"])
        if ax2:
            ax2.legend(loc="upper right", frameon=False, fontsize=template_style["font_size"])
    else:
        ax1.pie(plot_df[primary_y[0]],
                labels=plot_df[col_x],
                autopct="%1.1f%%",
                textprops={"color": template_style["axis_color"], 
                          "fontsize": template_style["font_size"]})
        ax1.axis("equal")

    if not template_style["grid"]:
        ax1.grid(False)
    else:
        ax1.grid(True, alpha=0.3)
    
    st.pyplot(fig)
    plt.close(fig)


def handle_ai_command(user_message, current_style, session_state, agent):
    """Handle AI chat command"""
    modifications = agent.process_command(user_message, current_style)

    if modifications:
        new_style = StyleUpdater.apply_modifications(current_style, modifications)
        session_state.custom_style = new_style

        session_state.chat_history.append(
            {"role": "user", "content": user_message}
        )
        session_state.chat_history.append(
            {
                "role": "assistant",
                "content": f"✅ Applied: {', '.join(modifications.keys())}"
            }
        )
        st.rerun()
    else:
        st.error("❌ I couldn't understand that command.")


def main():
    """Main application"""
    # =====================================================
    # PAGE CONFIG
    # =====================================================
    st.set_page_config(page_title="Excel Branded Charts with AI", layout="wide")
    st.title("📊 Excel Chart Generator with AI Assistant")

    # =====================================================
    # SESSION STATE INITIALIZATION
    # =====================================================
    if 'custom_style' not in st.session_state:
        st.session_state.custom_style = None
    if 'chat_history' not in st.session_state:
        st.session_state.chat_history = []
    if 'agent' not in st.session_state:
        st.session_state.agent = None

    # =====================================================
    # AI AGENT INITIALIZATION
    # =====================================================
    if st.session_state.agent is None:
        if not OPENAI_API_KEY:
            st.warning("❌ OPENAI_API_KEY not found in .env")
        else:
            try:
                st.session_state.agent = ChartModifierAgent()
            except Exception as e:
                st.session_state.agent = None
                st.error("❌ Failed to initialize AI Assistant")
                st.exception(e)

    # =====================================================
    # SETUP FONTS
    # =====================================================
    setup_fonts()

    # =====================================================
    # SIDEBAR - TEMPLATE SELECTION
    # =====================================================
    selected_template, template_config = render_sidebar_template_selection(TEMPLATE_REGISTRY)
    
    # Use custom style if available, otherwise use template style
    if st.session_state.custom_style is not None:
        TEMPLATE_STYLE = st.session_state.custom_style
        st.sidebar.info("🎨 Using AI-customized style")
    else:
        TEMPLATE_STYLE = template_config["style"]

    # =====================================================
    # SIDEBAR - FILE UPLOAD
    # =====================================================
    uploaded_file = render_sidebar_file_upload()

    # =====================================================
    # MAIN CONTENT
    # =====================================================
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            
            # Validate data
            if df.empty:
                st.error("❌ Uploaded file is empty. Please upload a file with data.")
                return
            
            numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()

            if not numeric_cols:
                st.error("❌ No numeric columns found in the uploaded file.")
                return

            # =====================================================
            # CHART CONFIGURATION
            # =====================================================
            col_x, aggregation, chart_type, primary_y, secondary_y = render_sidebar_chart_settings(
                df, template_config
            )

            if not primary_y:
                st.error("❌ Please select at least one primary Y-axis column.")
                return

            # =====================================================
            # DATA EDITING
            # =====================================================
            st.subheader("📝 Edit Data")
            edited_df = st.data_editor(df, width="stretch")

            # Process dataframe
            plot_df = process_dataframe(edited_df, col_x, aggregation, numeric_cols)

            # =====================================================
            # STYLE SETTINGS
            # =====================================================
            render_style_settings(TEMPLATE_STYLE)

            # =====================================================
            # CHART PREVIEW
            # =====================================================
            render_chart_preview(plot_df, col_x, primary_y, secondary_y, chart_type, TEMPLATE_STYLE)

            # =====================================================
            # EXPORT TO EXCEL
            # =====================================================
            st.subheader("⬇️ Export Branded Excel")

            if st.button("Generate Excel Report"):
                try:
                    output = create_excel_report(
                        plot_df, col_x, primary_y, secondary_y, chart_type,
                        template_config, TEMPLATE_STYLE
                    )
                    
                    st.success("✅ Excel file generated with AI-customized styling!")
                    
                    st.download_button(
                        "⬇️ Download Excel Report",
                        data=output,
                        file_name="ai_customized_excel_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"❌ Error generating Excel file: {str(e)}")
                    st.exception(e)

        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            st.exception(e)

    else:
        st.info("📤 Upload an Excel file from the sidebar to begin.")

    # =====================================================
    # AI CHAT ASSISTANT
    # =====================================================
    action = render_ai_chat_section(st.session_state, st.session_state.agent)
    
    if action == "reset_style":
        st.session_state.custom_style = None
        st.session_state.chat_history = []
        if st.session_state.agent:
            st.session_state.agent.reset_conversation()
        st.rerun()
    
    elif action == "clear_chat":
        st.session_state.chat_history = []
        if st.session_state.agent:
            st.session_state.agent.reset_conversation()
        st.rerun()
    
    elif action and isinstance(action, str) and action not in ["reset_style", "clear_chat"]:
        # It's a user message
        current_style = TEMPLATE_STYLE
        handle_ai_command(action, current_style, st.session_state, st.session_state.agent)


if __name__ == "__main__":
    main()