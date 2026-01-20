"""Configuration and constants"""
from pathlib import Path
import os
from dotenv import load_dotenv

load_dotenv()

# API Keys
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

# Paths
FONTS_DIR = Path(__file__).resolve().parent / "fonts"
ROBOTO_PATH = FONTS_DIR / "Roboto-Regular.ttf"

# Chart defaults
DEFAULT_CHART_TYPES = ["Bar Chart", "Line Chart", "Pie Chart"]
DEFAULT_AGGREGATIONS = ["None", "Sum", "Mean", "Count"]