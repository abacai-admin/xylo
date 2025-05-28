"""Chart color utilities for PowerPoint presentations"""

from pptx.dml.color import RGBColor
from typing import Tuple

# Custom brand colors (hex codes)
BRAND_COLORS = [
    '#005F6B',  # Teal
    '#F6A628',  # Orange
    '#FFC966',  # Light Orange
    '#FFD700',  # Gold
    '#000000'   # Black
]

def hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    """Convert hex color code to RGB tuple"""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def set_chart_colors(chart):
    """Apply brand colors to chart series"""
    try:
        for i, series in enumerate(chart.series):
            color_index = i % len(BRAND_COLORS)
            rgb = hex_to_rgb(BRAND_COLORS[color_index])
            
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = RGBColor(*rgb)
    except Exception as e:
        print(f"Warning: Could not apply chart colors - {str(e)}")
