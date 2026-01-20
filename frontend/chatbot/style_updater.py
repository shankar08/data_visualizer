"""
Handles updating chart styles based on chatbot commands
"""

from typing import Dict, Any
import copy


class StyleUpdater:
    """Updates template styles based on chatbot modifications"""
    
    @staticmethod
    def validate_color(color: str) -> bool:
        """Validate hex color format"""
        if not isinstance(color, str):
            return False
        if not color.startswith('#'):
            return False
        if len(color) != 7:
            return False
        try:
            int(color[1:], 16)
            return True
        except ValueError:
            return False
    
    @staticmethod
    def validate_font_size(size: int) -> bool:
        """Validate font size"""
        return isinstance(size, int) and 6 <= size <= 72
    
    @staticmethod
    def validate_modifications(modifications: Dict) -> Dict:
        """
        Validate and sanitize modifications
        
        Args:
            modifications: Dict of style changes from chatbot
            
        Returns:
            Validated modifications dict
        """
        valid_mods = {}
        
        # Validate colors
        color_fields = ["axis_line_color", "axis_color", "primary_color", "secondary_color"]
        for field in color_fields:
            if field in modifications:
                if StyleUpdater.validate_color(modifications[field]):
                    valid_mods[field] = modifications[field]
        
        # Validate font family
        if "font_family" in modifications:
            if isinstance(modifications["font_family"], str):
                valid_mods["font_family"] = modifications["font_family"]
        
        # Validate font size
        if "font_size" in modifications:
            if StyleUpdater.validate_font_size(modifications["font_size"]):
                valid_mods["font_size"] = modifications["font_size"]
        
        # Validate grid
        if "grid" in modifications:
            if isinstance(modifications["grid"], bool):
                valid_mods["grid"] = modifications["grid"]
        
        # Validate chart title
        if "chart_title" in modifications:
            if isinstance(modifications["chart_title"], str):
                valid_mods["chart_title"] = modifications["chart_title"]
        
        return valid_mods
    
    @staticmethod
    def apply_modifications(current_style: Dict, modifications: Dict) -> Dict:
        """
        Apply modifications to current style
        
        Args:
            current_style: Current style configuration
            modifications: Validated modifications to apply
            
        Returns:
            New style configuration
        """
        # Deep copy to avoid mutating original
        new_style = copy.deepcopy(current_style)
        
        # Validate modifications
        valid_mods = StyleUpdater.validate_modifications(modifications)
        
        # Apply each modification
        for key, value in valid_mods.items():
            new_style[key] = value
        
        return new_style