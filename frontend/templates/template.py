# templates/template.py

TEMPLATE_REGISTRY = {
    "Sales Dashboard": {
        "description": "Revenue, growth & product mix",
        "style": {
            "font_family": "Roboto",
            "font_size": 8,
            "axis_line_color": "#808080",
            "axis_color": "#404040",
            "primary_color": "#1f77b4",
            "secondary_color": "#ff7f0e",
            "grid": False
        },
        "default_chart_type": "Bar Chart",
        "default_primary_y": ["Revenue", "Profit"],
        "default_secondary_y": ["Growth %"],
        "sheet_name": "Sales Data",
        "chart_position": "G2",
        "table_style": "TableStyleMedium9",
        "number_formats": {
            "Revenue": "$#,##0",
            "Profit": "$#,##0",
            "Growth %": "0.0%"
        }
    },

    "Finance Report": {
        "description": "P&L, margins & cost analysis",
        "style": {
            "font_family": "Roboto",
            "font_size": 8,
            "axis_line_color": "#707070",
            "axis_color": "#303030",
            "primary_color": "#2ca02c",
            "secondary_color": "#d62728",
            "grid": True
        },
        "default_chart_type": "Line Chart",
        "default_primary_y": ["Revenue", "Expenses"],
        "default_secondary_y": ["Net Margin %"],
        "sheet_name": "Finance Data",
        "chart_position": "H2",
        "table_style": "TableStyleMedium2",
        "number_formats": {
            "Revenue": "$#,##0",
            "Expenses": "$#,##0",
            "Net Margin %": "0.0%"
        }
    },

    "Operations KPI": {
        "description": "Efficiency & throughput metrics",
        "style": {
            "font_family": "Roboto",
            "font_size": 8,
            "axis_line_color": "#909090",
            "axis_color": "#505050",
            "primary_color": "#9467bd",
            "secondary_color": "#8c564b",
            "grid": True
        },
        "default_chart_type": "Bar Chart",
        "default_primary_y": ["Throughput", "Cycle Time"],
        "default_secondary_y": ["Defect Rate %"],
        "sheet_name": "Ops Data",
        "chart_position": "G2",
        "table_style": "TableStyleMedium4",
        "number_formats": {
            "Throughput": "#,##0",
            "Cycle Time": "0.0",
            "Defect Rate %": "0.0%"
        }
    }
}