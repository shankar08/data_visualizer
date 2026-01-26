"""Chart formatting utilities"""
from openpyxl.chart import PieChart, BarChart, LineChart
from openpyxl.drawing.fill import ColorChoice, SolidColorFillProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Font as DrawingFont
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.label import DataLabelList


def hex_to_rgb(hex_color):
    """Convert hex color to RGB tuple"""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def apply_vba_formatting(chart, template_style):
    """
    Applies all VBA formatting to openpyxl chart object to match Streamlit preview
    Skips axis formatting for pie charts which don't have axes
    """
    
    # 1. CHART TITLE - Apply if provided
    if template_style.get("chart_title"):
        chart.title = template_style["chart_title"]
        # Format title text
        if chart.title:
            title_rgb = hex_to_rgb(template_style["axis_color"])
            if chart.title.tx is None:
                chart.title.tx = RichText()
            if chart.title.tx.p is None:
                chart.title.tx.p = [Paragraph()]
            
            para = chart.title.tx.p[0]
            if para.pPr is None:
                para.pPr = ParagraphProperties()
            if para.pPr.defRPr is None:
                para.pPr.defRPr = CharacterProperties()
            
            para.pPr.defRPr.sz = (template_style["font_size"] + 2) * 100
            para.pPr.defRPr.solidFill = ColorChoice(
                srgbClr=f'{title_rgb[0]:02X}{title_rgb[1]:02X}{title_rgb[2]:02X}'
            )
            if para.pPr.defRPr.latin is None:
                para.pPr.defRPr.latin = DrawingFont(typeface=template_style["font_family"])
            else:
                para.pPr.defRPr.latin.typeface = template_style["font_family"]
    else:
        chart.title = None
    
    # 2. SET GAP WIDTH = 50% (only for bar/line charts)
    if hasattr(chart, 'gapWidth'):
        chart.gapWidth = 50
    
    # 3. AXIS FORMATTING (skip for pie charts - they don't have axes)
    if not isinstance(chart, PieChart):
        axis_line_rgb = hex_to_rgb(template_style["axis_line_color"])
        axis_text_rgb = hex_to_rgb(template_style["axis_color"])
        
        for axis in [chart.x_axis, chart.y_axis]:
            # Grid formatting - simply remove gridlines if grid is False
            # Note: openpyxl doesn't have great gridline control, so we just remove them
            if not template_style.get("grid", False):
                axis.majorGridlines = None
                axis.minorGridlines = None
            # If grid is True, we leave the default gridlines (can't easily customize in openpyxl)
            
            # Format axis line (spine)
            if axis.spPr is None:
                axis.spPr = GraphicalProperties()
            
            line_props = LineProperties(w=12700)  # Thicker line to match matplotlib
            line_props.solidFill = ColorChoice(
                srgbClr=f'{axis_line_rgb[0]:02X}{axis_line_rgb[1]:02X}{axis_line_rgb[2]:02X}'
            )
            axis.spPr.ln = line_props
            
            # Tick marks
            axis.majorTickMark = "out"
            axis.minorTickMark = None
            axis.tickLblPos = "low"  # Position labels below/left
            
            # Format axis text (labels)
            if axis.txPr is None:
                axis.txPr = RichText()
            if axis.txPr.p is None:
                axis.txPr.p = [Paragraph()]
            
            para = axis.txPr.p[0]
            if para.pPr is None:
                para.pPr = ParagraphProperties()
            if para.pPr.defRPr is None:
                para.pPr.defRPr = CharacterProperties()
            
            if para.pPr.defRPr.latin is None:
                para.pPr.defRPr.latin = DrawingFont(typeface=template_style["font_family"])
            else:
                para.pPr.defRPr.latin.typeface = template_style["font_family"]
            
            para.pPr.defRPr.sz = template_style["font_size"] * 100
            para.pPr.defRPr.solidFill = ColorChoice(
                srgbClr=f'{axis_text_rgb[0]:02X}{axis_text_rgb[1]:02X}{axis_text_rgb[2]:02X}'
            )
        
        # Apply axis titles if provided
        if template_style.get("x_axis_title"):
            chart.x_axis.title = template_style["x_axis_title"]
            # Format x-axis title
            if chart.x_axis.title and chart.x_axis.title.tx:
                title_rgb = hex_to_rgb(template_style["axis_color"])
                if chart.x_axis.title.tx.rich is None:
                    chart.x_axis.title.tx.rich = RichText()
                if chart.x_axis.title.tx.rich.p is None:
                    chart.x_axis.title.tx.rich.p = [Paragraph()]
                
                para = chart.x_axis.title.tx.rich.p[0]
                if para.pPr is None:
                    para.pPr = ParagraphProperties()
                if para.pPr.defRPr is None:
                    para.pPr.defRPr = CharacterProperties()
                
                para.pPr.defRPr.sz = template_style["font_size"] * 100
                para.pPr.defRPr.solidFill = ColorChoice(
                    srgbClr=f'{title_rgb[0]:02X}{title_rgb[1]:02X}{title_rgb[2]:02X}'
                )
        
        if template_style.get("y_axis_title"):
            chart.y_axis.title = template_style["y_axis_title"]
            # Format y-axis title
            if chart.y_axis.title and chart.y_axis.title.tx:
                title_rgb = hex_to_rgb(template_style["axis_color"])
                if chart.y_axis.title.tx.rich is None:
                    chart.y_axis.title.tx.rich = RichText()
                if chart.y_axis.title.tx.rich.p is None:
                    chart.y_axis.title.tx.rich.p = [Paragraph()]
                
                para = chart.y_axis.title.tx.rich.p[0]
                if para.pPr is None:
                    para.pPr = ParagraphProperties()
                if para.pPr.defRPr is None:
                    para.pPr.defRPr = CharacterProperties()
                
                para.pPr.defRPr.sz = template_style["font_size"] * 100
                para.pPr.defRPr.solidFill = ColorChoice(
                    srgbClr=f'{title_rgb[0]:02X}{title_rgb[1]:02X}{title_rgb[2]:02X}'
                )
        
        # Hide top and right spines (like matplotlib does)
        chart.x_axis.spPr = GraphicalProperties() if chart.x_axis.spPr is None else chart.x_axis.spPr
        chart.y_axis.spPr = GraphicalProperties() if chart.y_axis.spPr is None else chart.y_axis.spPr
    
    # 4. LEGEND FORMATTING (works for all chart types)
    if chart.legend:
        legend_text_rgb = hex_to_rgb(template_style["axis_color"])
        chart.legend.position = 'r'  # Position legend on right
        
        if chart.legend.txPr is None:
            chart.legend.txPr = RichText()
        if chart.legend.txPr.p is None:
            chart.legend.txPr.p = [Paragraph()]
        
        para = chart.legend.txPr.p[0]
        if para.pPr is None:
            para.pPr = ParagraphProperties()
        if para.pPr.defRPr is None:
            para.pPr.defRPr = CharacterProperties()
        
        if para.pPr.defRPr.latin is None:
            para.pPr.defRPr.latin = DrawingFont(typeface=template_style["font_family"])
        else:
            para.pPr.defRPr.latin.typeface = template_style["font_family"]
        
        para.pPr.defRPr.sz = template_style["font_size"] * 100
        para.pPr.defRPr.solidFill = ColorChoice(
            srgbClr=f'{legend_text_rgb[0]:02X}{legend_text_rgb[1]:02X}{legend_text_rgb[2]:02X}'
        )
    
    # 5. PLOT AREA - Remove fill to match transparent background
    if hasattr(chart, 'plot_area') and chart.plot_area:
        chart.plot_area.graphicalProperties = GraphicalProperties()
        # Make plot area background transparent
        chart.plot_area.graphicalProperties.noFill = True