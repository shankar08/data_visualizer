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
        }
    }
}