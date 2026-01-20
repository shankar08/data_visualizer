# templates/template.py
# VBA-matched styling templates for Excel chart generation

TEMPLATE_REGISTRY = {
    "Sales Dashboard": {
        "description": "Revenue, growth & product mix",
        "sheet_name": "Sales Data",
        "default_chart_type": "Bar Chart",
        "chart_position": "G2",
        "table_style": "TableStyleMedium9",
        
        # VBA-matched styling
        "style": {
            "font_family": "Roboto",
            "font_size": 7,  # VBA uses size 7
            "axis_line_color": "#BFBFBF",  # RGB(191, 191, 191) - VBA default
            "axis_color": "#404040",  # RGB(64, 64, 64) - VBA default
            "primary_color": "#1f77b4",
            "secondary_color": "#ff7f0e",
            "grid": False  # VBA: HasMajorGridlines = False
        },
        
        # Optional: Default data columns
        "default_primary_y": ["Revenue", "Sales"],
        "default_secondary_y": ["Growth %"],
        
        # Optional: Number formatting
        "number_formats": {
            "Revenue": "$#,##0",
            "Sales": "#,##0",
            "Growth %": "0.0%"
        }
    },
    
    "Finance Report": {
        "description": "P&L, margins & cost analysis",
        "sheet_name": "Financial Data",
        "default_chart_type": "Bar Chart",
        "chart_position": "H2",
        "table_style": "TableStyleMedium2",
        
        # VBA-matched styling
        "style": {
            "font_family": "Roboto",
            "font_size": 7,
            "axis_line_color": "#BFBFBF",  # Same as VBA
            "axis_color": "#404040",  # Same as VBA
            "primary_color": "#2ca02c",
            "secondary_color": "#d62728",
            "grid": False  # VBA: HasMajorGridlines = False
        },
        
        "default_primary_y": ["Revenue", "Profit"],
        "default_secondary_y": ["Margin %"],
        
        "number_formats": {
            "Revenue": "$#,##0",
            "Profit": "$#,##0",
            "Cost": "$#,##0",
            "Margin %": "0.0%"
        }
    },
    
    "Operations KPI": {
        "description": "Efficiency & throughput metrics",
        "sheet_name": "Operations Data",
        "default_chart_type": "Line Chart",
        "chart_position": "G2",
        "table_style": "TableStyleMedium6",
        
        # VBA-matched styling
        "style": {
            "font_family": "Roboto",
            "font_size": 7,
            "axis_line_color": "#BFBFBF",  # Same as VBA
            "axis_color": "#404040",  # Same as VBA
            "primary_color": "#9467bd",
            "secondary_color": "#8c564b",
            "grid": False  # VBA: HasMajorGridlines = False
        },
        
        "default_primary_y": ["Throughput", "Efficiency"],
        "default_secondary_y": ["Target %"],
        
        "number_formats": {
            "Throughput": "#,##0",
            "Efficiency": "0.0%",
            "Target %": "0.0%"
        }
    },
    
    # VBA Exact Match Template (mirrors your VBA code exactly)
    "VBA Exact Match": {
        "description": "Exact replication of VBA formatting",
        "sheet_name": "Data",
        "default_chart_type": "Bar Chart",
        "chart_position": "G2",
        "table_style": "TableStyleMedium9",
        
        # Exact VBA specifications
        "style": {
            "font_family": "Roboto",
            "font_size": 7,  # VBA: .Size = 7
            "axis_line_color": "#BFBFBF",  # VBA: RGB(191, 191, 191)
            "axis_color": "#404040",  # VBA: RGB(64, 64, 64)
            "primary_color": "#4472C4",  # Excel default blue
            "secondary_color": "#ED7D31",  # Excel default orange
            "grid": False  # VBA: HasMajorGridlines = False
        },
        
        "default_primary_y": [],
        "default_secondary_y": [],
        
        "number_formats": {}
    }
}


# Helper function to get template
def get_template(template_name):
    """
    Retrieve a template configuration by name.
    
    Args:
        template_name (str): Name of the template
        
    Returns:
        dict: Template configuration or None if not found
    """
    return TEMPLATE_REGISTRY.get(template_name)


# Helper function to list all templates
def list_templates():
    """
    Get all available template names.
    
    Returns:
        list: List of template names
    """
    return list(TEMPLATE_REGISTRY.keys())


# Helper function to validate template
def validate_template(template_config):
    """
    Validate that a template has all required fields.
    
    Args:
        template_config (dict): Template configuration to validate
        
    Returns:
        tuple: (is_valid, error_message)
    """
    required_fields = ["description", "style"]
    required_style_fields = [
        "font_family", "font_size", "axis_line_color", 
        "axis_color", "primary_color", "grid"
    ]
    
    # Check top-level fields
    for field in required_fields:
        if field not in template_config:
            return False, f"Missing required field: {field}"
    
    # Check style fields
    style = template_config.get("style", {})
    for field in required_style_fields:
        if field not in style:
            return False, f"Missing required style field: {field}"
    
    return True, "Template is valid"