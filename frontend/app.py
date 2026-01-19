import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from io import BytesIO
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.drawing.fill import SolidColorFillProperties, ColorChoice
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RichTextProperties
from pathlib import Path
import sys

# Add project root to path for imports (MUST be before template import)
BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from templates.template import TEMPLATE_REGISTRY


def main():

    # =====================================================
    # HELPER FUNCTION: Apply VBA-matching Chart Formatting
    # =====================================================
    def apply_vba_formatting(chart, template_style):
        """
        Applies all VBA formatting to openpyxl chart object:
        - Removes title
        - Sets gap width to 50%
        - Formats axes (lines, tick marks, fonts)
        - Formats legend
        - Removes gridlines
        """
        from openpyxl.chart.shapes import GraphicalProperties
        from openpyxl.drawing.text import Font as DrawingFont
        
        # 1. REMOVE TITLE (VBA: ch.HasTitle = False)
        chart.title = None
        
        # 2. SET GAP WIDTH = 50% (VBA: cg.GapWidth = 50)
        if hasattr(chart, 'gapWidth'):
            chart.gapWidth = 50
        
        # 3. AXIS FORMATTING
        # Convert hex colors to RGB tuples
        def hex_to_rgb(hex_color):
            hex_color = hex_color.lstrip('#')
            return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        
        axis_line_rgb = hex_to_rgb(template_style["axis_line_color"])
        axis_text_rgb = hex_to_rgb(template_style["axis_color"])
        
        # Format both X and Y axes
        for axis in [chart.x_axis, chart.y_axis]:
            # Remove major gridlines (VBA: HasMajorGridlines = False)
            axis.majorGridlines = None
            
            # Axis line formatting (VBA: Format.Line properties)
            if axis.spPr is None:
                axis.spPr = GraphicalProperties()
            
            # Create line with proper color
            line_props = LineProperties(
                w=6350  # 0.5pt in EMUs (VBA: Weight = 0.5)
            )
            # Set color using ColorChoice with srgbClr
            line_props.solidFill = ColorChoice(
                srgbClr=f'{axis_line_rgb[0]:02X}{axis_line_rgb[1]:02X}{axis_line_rgb[2]:02X}'
            )
            axis.spPr.ln = line_props
            
            # Tick marks (VBA: MajorTickMark = xlTickMarkOutside)
            axis.majorTickMark = "out"
            axis.minorTickMark = None
            
            # Tick label font formatting (VBA: TickLabels.Font)
            if axis.txPr is None:
                axis.txPr = RichText()
            if axis.txPr.p is None:
                axis.txPr.p = [Paragraph()]
            
            para = axis.txPr.p[0]
            if para.pPr is None:
                para.pPr = ParagraphProperties()
            if para.pPr.defRPr is None:
                para.pPr.defRPr = CharacterProperties()
            
            # Set font: Roboto, size 7, RGB(64,64,64)
            if para.pPr.defRPr.latin is None:
                para.pPr.defRPr.latin = DrawingFont(typeface=template_style["font_family"])
            else:
                para.pPr.defRPr.latin.typeface = template_style["font_family"]
            
            para.pPr.defRPr.sz = template_style["font_size"] * 100  # Points to hundredths
            # Use ColorChoice for text color
            para.pPr.defRPr.solidFill = ColorChoice(
                srgbClr=f'{axis_text_rgb[0]:02X}{axis_text_rgb[1]:02X}{axis_text_rgb[2]:02X}'
            )
        
        # 4. LEGEND FORMATTING (VBA: Legend.Font)
        if chart.legend:
            if chart.legend.txPr is None:
                chart.legend.txPr = RichText()
            if chart.legend.txPr.p is None:
                chart.legend.txPr.p = [Paragraph()]
            
            para = chart.legend.txPr.p[0]
            if para.pPr is None:
                para.pPr = ParagraphProperties()
            if para.pPr.defRPr is None:
                para.pPr.defRPr = CharacterProperties()
            
            # Set font properly
            if para.pPr.defRPr.latin is None:
                para.pPr.defRPr.latin = DrawingFont(typeface=template_style["font_family"])
            else:
                para.pPr.defRPr.latin.typeface = template_style["font_family"]
            
            para.pPr.defRPr.sz = template_style["font_size"] * 100
            # Use ColorChoice for text color
            para.pPr.defRPr.solidFill = ColorChoice(
                srgbClr=f'{axis_text_rgb[0]:02X}{axis_text_rgb[1]:02X}{axis_text_rgb[2]:02X}'
            )

    # =====================================================
    # PATH SETUP FOR FONTS
    # =====================================================
    # Since app.py is in frontend/ and fonts are in frontend/fonts/
    FONTS_DIR = Path(__file__).resolve().parent / "fonts"
    ROBOTO_PATH = FONTS_DIR / "Roboto-Regular.ttf"

    if ROBOTO_PATH.exists():
        fm.fontManager.addfont(str(ROBOTO_PATH))
        plt.rcParams["font.family"] = "Roboto"
    else:
        st.warning(f"Roboto font not found at {ROBOTO_PATH}, using default font.")

    # =====================================================
    # PAGE CONFIG
    # =====================================================
    st.set_page_config(page_title="Excel Branded Charts", layout="wide")
    st.title("📊 Excel → Branded Dynamic Chart Generator")

    # =====================================================
    # TEMPLATE SELECTION
    # =====================================================
    st.sidebar.header("📄 Report Template")
    selected_template = st.sidebar.selectbox(
        "Select Report Template",
        list(TEMPLATE_REGISTRY.keys())
    )
    template_config = TEMPLATE_REGISTRY[selected_template]
    st.sidebar.caption(template_config["description"])

    TEMPLATE_STYLE = template_config["style"]

    st.sidebar.divider()
    st.sidebar.header("📂 Import & Chart Settings")
    uploaded_file = st.sidebar.file_uploader("Upload Excel", type=["xlsx"])

    # =====================================================
    # MAIN APP
    # =====================================================
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
        all_cols = df.columns.tolist()

        col_x = st.sidebar.selectbox("X-axis (Group By)", all_cols)
        aggregation = st.sidebar.selectbox(
            "Aggregation", ["None", "Sum", "Mean", "Count"]
        )

        # Chart Type Defaults
        default_chart_type = template_config.get("default_chart_type", "Bar Chart")
        chart_type = st.sidebar.selectbox(
            "Chart Type",
            ["Bar Chart", "Line Chart", "Pie Chart"],
            index=["Bar Chart", "Line Chart", "Pie Chart"].index(default_chart_type)
        )

        # Y-axis defaults
        default_primary = [c for c in template_config.get("default_primary_y", []) if c in numeric_cols]
        default_secondary = [c for c in template_config.get("default_secondary_y", []) if c in numeric_cols]

        if chart_type == "Pie Chart":
            primary_y = st.sidebar.selectbox(
                "Y-axis",
                numeric_cols,
                index=0 if not default_primary else numeric_cols.index(default_primary[0])
            )
            secondary_y = []
        else:
            primary_y = st.sidebar.multiselect(
                "Primary Y-axis",
                numeric_cols,
                default=default_primary or numeric_cols[:1]
            )
            secondary_y = st.sidebar.multiselect(
                "Secondary Y-axis",
                [c for c in numeric_cols if c not in primary_y],
                default=default_secondary
            )

        # =====================================================
        # DATA EDITOR
        # =====================================================
        st.subheader("📝 Edit Data")
        edited_df = st.data_editor(df, use_container_width=True)

        # =====================================================
        # AGGREGATION
        # =====================================================
        if aggregation != "None":
            if aggregation == "Sum":
                plot_df = edited_df.groupby(col_x)[numeric_cols].sum().reset_index()
            elif aggregation == "Mean":
                plot_df = edited_df.groupby(col_x)[numeric_cols].mean().reset_index()
            elif aggregation == "Count":
                plot_df = edited_df.groupby(col_x)[numeric_cols].count().reset_index()
        else:
            plot_df = edited_df.copy()

        # =====================================================
        # STREAMLIT PREVIEW — TEMPLATE STYLED
        # =====================================================
        st.subheader("👀 Chart Preview (Template Style)")

        plt.rcParams.update({
            "font.size": TEMPLATE_STYLE["font_size"]
        })

        fig, ax1 = plt.subplots(figsize=(8, 4))
        for spine in ["top", "right"]:
            ax1.spines[spine].set_visible(False)
        for spine in ["bottom", "left"]:
            ax1.spines[spine].set_color(TEMPLATE_STYLE["axis_line_color"])
            ax1.spines[spine].set_linewidth(0.5)
        ax1.tick_params(axis="both", colors=TEMPLATE_STYLE["axis_color"], direction="out")

        ax2 = ax1.twinx() if secondary_y else None

        if chart_type in ["Bar Chart", "Line Chart"]:
            for col in primary_y:
                if chart_type == "Bar Chart":
                    ax1.bar(plot_df[col_x], plot_df[col],
                            label=col,
                            color=TEMPLATE_STYLE["primary_color"],
                            alpha=0.85,
                            width=0.5)  # Gap width equivalent
                else:
                    ax1.plot(plot_df[col_x], plot_df[col],
                             label=col,
                             linewidth=1.5,
                             color=TEMPLATE_STYLE["primary_color"])

            if ax2:
                ax2.spines["top"].set_visible(False)
                ax2.spines["right"].set_color(TEMPLATE_STYLE["axis_line_color"])
                ax2.spines["right"].set_linewidth(0.5)
                ax2.tick_params(colors=TEMPLATE_STYLE["axis_color"], direction="out")
                for col in secondary_y:
                    ax2.plot(plot_df[col_x], plot_df[col],
                             linestyle="--",
                             linewidth=1.2,
                             label=col,
                             color=TEMPLATE_STYLE["secondary_color"])

            ax1.legend(loc="upper left", frameon=False, fontsize=TEMPLATE_STYLE["font_size"])
            if ax2:
                ax2.legend(loc="upper right", frameon=False, fontsize=TEMPLATE_STYLE["font_size"])
        else:
            ax1.pie(plot_df[primary_y],
                    labels=plot_df[col_x],
                    autopct="%1.1f%%",
                    textprops={"color": TEMPLATE_STYLE["axis_color"], 
                              "fontsize": TEMPLATE_STYLE["font_size"]})
            ax1.axis("equal")

        if not TEMPLATE_STYLE["grid"]:
            ax1.grid(False)
        else:
            ax1.grid(True, alpha=0.3)

        st.pyplot(fig)

        # =====================================================
        # EXCEL EXPORT — FULLY VBA-MATCHED FORMATTING
        # =====================================================
        st.subheader("⬇️ Export Branded Excel (VBA-Matched Formatting)")

        if st.button("Generate Excel Report"):

            wb = Workbook()
            ws = wb.active
            ws.title = template_config.get("sheet_name", "Data")

            number_formats = template_config.get("number_formats", {})

            # Write headers
            for col_idx, col_name in enumerate(plot_df.columns, start=1):
                cell = ws.cell(row=1, column=col_idx, value=col_name)
                cell.font = cell.font.copy(bold=True)

            # Write data with number formatting
            for row_idx, row in enumerate(plot_df.itertuples(index=False), start=2):
                for col_idx, (col_name, value) in enumerate(zip(plot_df.columns, row), start=1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=value)
                    if col_name in number_formats:
                        cell.number_format = number_formats[col_name]

            # Create Table
            end_row = ws.max_row
            end_col = ws.max_column
            table_ref = f"A1:{ws.cell(row=end_row, column=end_col).coordinate}"
            table = Table(displayName="DataTable", ref=table_ref)
            style = TableStyleInfo(
                name=template_config.get("table_style", "TableStyleMedium9"),
                showRowStripes=True,
                showColumnStripes=False
            )
            table.tableStyleInfo = style
            ws.add_table(table)

            # -----------------------------
            # Add chart with VBA formatting
            # -----------------------------
            x_idx = 1
            chart_pos = template_config.get("chart_position", "G2")
            chart_type_to_use = template_config.get("default_chart_type", chart_type)

            if chart_type_to_use in ["Bar Chart", "Line Chart"]:
                chart = BarChart() if chart_type_to_use == "Bar Chart" else LineChart()
                
                # No title (VBA: HasTitle = False) - will be removed in apply_vba_formatting
                chart.y_axis.title = None
                chart.x_axis.title = None

                # Primary series
                for idx, col_name in enumerate(primary_y):
                    y_idx = plot_df.columns.get_loc(col_name) + 1
                    data = Reference(ws, min_col=y_idx, min_row=1, max_row=end_row)
                    chart.add_data(data, titles_from_data=True)
                    
                    # Apply primary color
                    if chart.series and idx < len(chart.series):
                        chart.series[idx].graphicalProperties.solidFill = TEMPLATE_STYLE["primary_color"].lstrip("#")

                # Categories
                cats = Reference(ws, min_col=x_idx, min_row=2, max_row=end_row)
                chart.set_categories(cats)

                # Secondary series
                if secondary_y:
                    sec_chart = LineChart()
                    sec_chart.y_axis.axId = 200
                    
                    for idx, col_name in enumerate(secondary_y):
                        y_idx = plot_df.columns.get_loc(col_name) + 1
                        data = Reference(ws, min_col=y_idx, min_row=1, max_row=end_row)
                        sec_chart.add_data(data, titles_from_data=True)
                        
                        # Apply secondary color
                        if sec_chart.series and idx < len(sec_chart.series):
                            sec_chart.series[idx].graphicalProperties.solidFill = TEMPLATE_STYLE["secondary_color"].lstrip("#")

                    # Apply VBA formatting to secondary chart too
                    apply_vba_formatting(sec_chart, TEMPLATE_STYLE)
                    chart += sec_chart
                
                # Apply all VBA formatting
                apply_vba_formatting(chart, TEMPLATE_STYLE)
                
            else:
                chart = PieChart()
                y_idx = plot_df.columns.get_loc(primary_y) + 1
                data = Reference(ws, min_col=y_idx, min_row=1, max_row=end_row)
                labels = Reference(ws, min_col=x_idx, min_row=2, max_row=end_row)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(labels)
                
                # Pie charts don't have gap width, but apply other formatting
                chart.title = None
                if chart.legend:
                    apply_vba_formatting(chart, TEMPLATE_STYLE)

            ws.add_chart(chart, chart_pos)

            # Export Excel
            output = BytesIO()
            wb.save(output)
            output.seek(0)

            st.success("✅ VBA-matched branded Excel file generated with:")
            st.markdown("""
            - ✅ Gap Width = 50%
            - ✅ No chart title
            - ✅ No gridlines
            - ✅ Axis lines: 0.5pt weight, custom color
            - ✅ Tick marks: Outside
            - ✅ Fonts: Roboto, size 7, custom colors
            - ✅ Legend formatted
            """)
            
            st.download_button(
                "⬇️ Download Excel Report",
                data=output,
                file_name="vba_matched_excel_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.info("Upload an Excel file from the sidebar to begin.")

# =====================================================
# RUN MAIN IF EXECUTED DIRECTLY
# =====================================================
if __name__ == "__main__":
    main()