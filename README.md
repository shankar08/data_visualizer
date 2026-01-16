# Excel Branded Charts — Streamlit App

📊 **Generate branded Excel reports with dynamic named-range charts and template-based styling directly from your data.**

---

## Project Overview

This project is a **Python & Streamlit-based application** that allows users to:

- Upload Excel data
- Edit and aggregate data interactively
- Preview charts (Bar, Line, Pie) styled according to predefined templates
- Export branded Excel reports with dynamic named-range tables and charts
- Support multi-series charts, secondary Y-axis, and template-defined colors/fonts

All template configurations are **Python-driven**, so you don’t need to maintain multiple Excel template files.

---

## Folder Structure

Data_visualizer/
│
├─ src/
│ ├─ frontend/
│ │ ├─ app.py # Main Streamlit app
│ │ ├─ fonts/
│ │ │ └─ Roboto-Regular.ttf # Custom font for charts
│ │ └─ templates/
│ │ └─ template.py # Template registry & style definitions
│ │
│ └─ streamlit_app.py # Entry point to launch Streamlit app
│
├─ README.md
└─ requirements.txt # Python dependencies

---

## Features

- **Multi-template support**: Choose different reporting templates (e.g., Sales Dashboard, Finance Report, Operations KPI) with custom styling.
- **Dynamic charts in UI**: Preview charts styled according to your selected template before exporting.
- **Multi-series charts**: Supports primary and secondary Y-axis series.
- **Aggregation options**: Sum, Mean, Count, or raw data.
- **Branded Excel export**: Excel reports include:
  - Dynamic named-range tables
  - Template-based colors, fonts, and styles
  - Charts embedded in Excel (Bar, Line, Pie)
- **Custom fonts**: Uses Roboto font in charts via bundled TTF.

---
