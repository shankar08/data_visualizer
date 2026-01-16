def main():
    import streamlit as st
    import pandas as pd
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm
    from io import BytesIO
    from openpyxl import Workbook
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference
    from pathlib import Path
    from templates.template import TEMPLATE_REGISTRY

    # =====================================================
    # PATH SETUP
    # =====================================================
    BASE_DIR = Path(__file__).resolve().parent
    FONTS_DIR = BASE_DIR / "fonts"
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
            "Chart Type (Preview Only)",
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
                "Secondary Y-axis (Preview Only)",
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
                            alpha=0.85)
                else:
                    ax1.plot(plot_df[col_x], plot_df[col],
                             label=col,
                             linewidth=1.5)

            if ax2:
                ax2.spines["right"].set_color(TEMPLATE_STYLE["axis_line_color"])
                ax2.tick_params(colors=TEMPLATE_STYLE["axis_color"])
                for col in secondary_y:
                    ax2.plot(plot_df[col_x], plot_df[col],
                             linestyle="--",
                             linewidth=1.2,
                             label=col,
                             color=TEMPLATE_STYLE["secondary_color"])

            ax1.legend(loc="upper left", frameon=False)
            if ax2:
                ax2.legend(loc="upper right", frameon=False)
        else:
            ax1.pie(plot_df[primary_y],
                    labels=plot_df[col_x],
                    autopct="%1.1f%%",
                    textprops={"color": TEMPLATE_STYLE["axis_color"]})
            ax1.axis("equal")

        if not TEMPLATE_STYLE["grid"]:
            ax1.grid(False)

        st.pyplot(fig)

        # =====================================================
        # EXCEL EXPORT — FULLY PYTHON-DRIVEN
        # =====================================================
        st.subheader("⬇️ Export Branded Excel")

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
            # Add chart
            # -----------------------------
            x_idx = 1
            chart_pos = template_config.get("chart_position", "G2")
            chart_type_to_use = template_config.get("default_chart_type", chart_type)

            if chart_type_to_use in ["Bar Chart", "Line Chart"]:
                chart = BarChart() if chart_type_to_use == "Bar Chart" else LineChart()
                chart.title = selected_template
                chart.x_axis.title = col_x
                chart.y_axis.title = "Primary Axis"

                # Primary series
                for idx, col_name in enumerate(primary_y):
                    y_idx = plot_df.columns.get_loc(col_name) + 1
                    data = Reference(ws, min_col=y_idx, min_row=1, max_row=end_row)
                    chart.add_data(data, titles_from_data=True)
                    chart.series[idx].graphicalProperties.solidFill = TEMPLATE_STYLE["primary_color"].lstrip("#")

                # Categories
                cats = Reference(ws, min_col=x_idx, min_row=2, max_row=end_row)
                chart.set_categories(cats)

                # Secondary series
                if secondary_y:
                    sec_chart = LineChart()
                    sec_chart.y_axis.axId = 200
                    sec_chart.y_axis.title = "Secondary Axis"
                    sec_chart.x_axis.title = col_x

                    for idx, col_name in enumerate(secondary_y):
                        y_idx = plot_df.columns.get_loc(col_name) + 1
                        data = Reference(ws, min_col=y_idx, min_row=1, max_row=end_row)
                        sec_chart.add_data(data, titles_from_data=True)
                        sec_chart.series[idx].graphicalProperties.solidFill = TEMPLATE_STYLE["secondary_color"].lstrip("#")

                    chart += sec_chart
            else:
                chart = PieChart()
                y_idx = plot_df.columns.get_loc(primary_y) + 1
                data = Reference(ws, min_col=y_idx, min_row=1, max_row=end_row)
                labels = Reference(ws, min_col=x_idx, min_row=2, max_row=end_row)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(labels)

            ws.add_chart(chart, chart_pos)

            # Export Excel
            output = BytesIO()
            wb.save(output)
            output.seek(0)

            st.success("✅ Branded Excel file generated")
            st.download_button(
                "⬇️ Download Excel Report",
                data=output,
                file_name="branded_excel_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.info("Upload an Excel file from the sidebar to begin.")

# =====================================================
# RUN MAIN IF EXECUTED DIRECTLY
# =====================================================
if __name__ == "__main__":
    main()