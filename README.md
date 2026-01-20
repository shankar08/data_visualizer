# 📊 Excel Chart Generator with AI Assistant

A powerful Streamlit application that generates professionally styled Excel reports with AI-powered chart customization. Create bar charts, line charts, and pie charts with custom branding, all powered by an intelligent AI assistant.

## ✨ Features

- **Multi-Chart Support**: Bar charts, line charts, and pie charts
- **Data Editing**: Edit your data directly in the app before exporting
- **Multiple Aggregations**: Sum, Mean, Count, or no aggregation
- **Template System**: Pre-configured report templates with custom styling
- **AI Assistant**: Natural language commands to modify chart styles in real-time
- **Styled Excel Export**: Export professionally formatted Excel reports with styled charts
- **Dual Axis Support**: Create charts with primary and secondary Y-axes
- **Responsive UI**: Clean, modern interface with Streamlit

## 🚀 Quick Start

### Prerequisites

- Python 3.11+
- pip or conda

### Installation

1. **Clone or download the project**

```bash
cd data_visualizer
```

2. **Create a virtual environment**

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

3. **Install dependencies**

```bash
pip install -r requirements.txt
```

4. **Set up environment variables**
   Create a `.env` file in the root directory:

```
OPENAI_API_KEY=your_openai_api_key_here
```

5. **Run the application**

```bash
streamlit run streamlit.py
```

The app will open at `http://localhost:8501`

## 📁 Project Structure

```
data_visualizer/
├── streamlit.py                 # Entry point
├── requirements.txt             # Project dependencies
├── README.md                    # Project documentation
├── .env                        # Environment variables (create this)
├── .gitignore                  # Git ignore rules
└── src/
    └── frontend/
        ├── app.py              # Main Streamlit application
        ├── config.py           # Configuration & constants
        ├── chart_formatter.py  # Chart formatting utilities
        ├── excel_exporter.py   # Excel export functionality
        ├── ui_components.py    # Streamlit UI components
        ├── fonts/
        │   └── Roboto-Regular.ttf
        ├── chatbot/
        │   ├── chart_modifier.py
        │   └── style_updater.py
        └── templates/
            └── template.py
```

## 🎯 How to Use

### 1. Upload Your Data

- Click the file uploader in the sidebar
- Select an Excel file with numeric data

### 2. Configure Your Chart

- **X-axis**: Select the column to group by
- **Aggregation**: Choose Sum, Mean, Count, or None
- **Chart Type**: Select Bar Chart, Line Chart, or Pie Chart
- **Y-axis**: Select primary (and optional secondary) columns

### 3. Edit Data (Optional)

- Modify values directly in the data editor table
- Changes are reflected in real-time

### 4. Preview & Customize

- View the chart preview
- Check current style settings (fonts, colors, grid)
- Use the AI assistant to customize the chart

### 5. Export to Excel

- Click "Generate Excel Report"
- Download the styled Excel file with embedded charts

### 6. AI Assistant Commands

Talk to the AI to modify your chart with natural language:

- `"chart title Sales Report"` - Set chart title
- `"remove grid"` - Hide gridlines
- `"make font Arial 10"` - Change font
- `"primary color red"` - Change primary color
- `"axis title X = Months"` - Set axis titles

## 🎨 Templates

The app includes pre-configured templates with:

- Default chart types
- Color schemes
- Font families
- Number formats
- Table styles

Create custom templates by modifying `frontend/templates/template.py`

## 📚 Module Documentation

### app.py - Main Application

The main Streamlit app that orchestrates all features.

**Key Functions:**

- `setup_fonts()` - Initialize matplotlib fonts
- `process_dataframe()` - Handle data aggregation
- `render_chart_preview()` - Display matplotlib preview
- `handle_ai_command()` - Process AI assistant commands
- `main()` - Main application flow

### config.py - Configuration

Centralized configuration and constants.

**Exports:**

- `OPENAI_API_KEY` - API key from environment
- `ROBOTO_PATH` - Path to Roboto font
- `DEFAULT_CHART_TYPES` - Available chart types
- `DEFAULT_AGGREGATIONS` - Available aggregations

### chart_formatter.py - Chart Formatting

Handles all OpenPyXL chart styling to match Streamlit preview.

**Key Functions:**

- `hex_to_rgb()` - Convert hex colors to RGB
- `apply_vba_formatting()` - Apply VBA-style formatting to charts

**Features:**

- Axis line styling
- Axis label formatting
- Legend customization
- Grid removal
- Transparent plot area
- Pie chart data labels

### excel_exporter.py - Excel Export

Creates and exports Excel workbooks with styled charts.

**Key Functions:**

- `create_excel_report()` - Main export function
- `create_chart()` - Creates and styles chart objects

**Supports:**

- Bar charts
- Line charts
- Pie charts
- Dual-axis charts
- Data labels
- Custom table styling

### ui_components.py - UI Components

Reusable Streamlit UI component builders.

**Components:**

- `render_sidebar_template_selection()` - Template selector
- `render_sidebar_file_upload()` - File uploader
- `render_sidebar_chart_settings()` - Chart configuration
- `render_style_settings()` - Style preview
- `render_ai_chat_section()` - AI assistant interface

## 🤖 AI Assistant Features

The AI assistant uses OpenAI's API to understand natural language commands and modify chart styles.

**Available Modifications:**

- Chart title
- Axis titles
- Font family and size
- Colors (primary, secondary, axis)
- Grid visibility
- Aggregation method

### Example Commands

```
"Change the title to Sales Data"
"Make the bars red"
"Remove the gridlines"
"Use Arial font at 12 points"
"Add a secondary y-axis"
"Show percentage on pie slices"
```

## 📊 Supported Chart Types

### Bar Chart

- Primary and secondary Y-axes
- Customizable colors
- Grid options
- Axis titles and labels

### Line Chart

- Multiple series support
- Dashed line for secondary series
- Customizable markers
- Responsive axis formatting

### Pie Chart

- Automatic percentage labels
- Value display
- Legend support
- Color customization

## 🎛️ Data Aggregation

- **None**: Use raw data as-is
- **Sum**: Sum values by X-axis group
- **Mean**: Average values by X-axis group
- **Count**: Count records by X-axis group

## 🔧 Customization

### Change Fonts

1. Add your font file to `frontend/fonts/`
2. Update `FONTS_DIR` and `ROBOTO_PATH` in `config.py`
3. Restart the app

### Modify Templates

Edit `frontend/templates/template.py`:

```python
TEMPLATE_REGISTRY = {
    "Custom Template": {
        "description": "My custom template",
        "style": {
            "primary_color": "#FF5733",
            "font_family": "Arial",
            "font_size": 11,
            # ... more settings
        },
        # ... more config
    }
}
```

### Add New Chart Types

1. Extend `chart_formatter.py` with new chart type support
2. Add chart type to `DEFAULT_CHART_TYPES` in `config.py`
3. Update `excel_exporter.py` with export logic

## 🐛 Troubleshooting

### "ModuleNotFoundError: No module named 'templates'"

- Ensure you're running from the correct directory
- Check that `src/` is in `sys.path`
- Verify file structure matches project layout

### Chart not visible in Excel export

- Check that chart type matches template default
- Verify data contains numeric columns
- Ensure axis references are correct

### AI Assistant not responding

- Check `OPENAI_API_KEY` is set in `.env`
- Verify API key is valid and has sufficient quota
- Check internet connection

### Font not appearing

- Verify font file exists in `frontend/fonts/`
- Check font filename in `config.py`
- Ensure matplotlib can access the font

## 📦 Dependencies

- `streamlit>=1.28.0` - Web app framework
- `pandas>=2.0.0` - Data manipulation
- `matplotlib>=3.7.0` - Chart rendering
- `openpyxl>=3.10.0` - Excel file handling
- `python-dotenv>=1.0.0` - Environment variables
- `openai>=1.0.0` - AI assistant

See `requirements.txt` for complete list.

## 🚀 Deployment

### Streamlit Cloud

1. Push code to GitHub
2. Connect your GitHub repo to Streamlit Cloud
3. Set `OPENAI_API_KEY` in Secrets Manager
4. Deploy!

### Docker

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY . .
RUN pip install -r requirements.txt
CMD ["streamlit", "run", "streamlit.py"]
```

## 📝 Example Use Cases

1. **Sales Reports** - Create branded monthly sales charts
2. **Financial Analysis** - Generate multi-axis financial reports
3. **Performance Dashboards** - Track KPIs with custom styling
4. **Academic Papers** - Create publication-ready charts
5. **Business Proposals** - Professional-looking chart exports

## 🔐 Security

- Never commit `.env` file with real API keys
- Keep `OPENAI_API_KEY` secret
- Validate all user inputs
- Sanitize data before export

## 📄 License

MIT License - Feel free to use this project for personal or commercial purposes.

## 👥 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## 📧 Support

For issues and questions:

- Check the troubleshooting section above
- Review project structure and file locations
- Check Streamlit and OpenAI documentation

## 🎓 Learning Resources

- [Streamlit Documentation](https://docs.streamlit.io)
- [OpenPyXL Documentation](https://openpyxl.readthedocs.io)
- [OpenAI API Documentation](https://platform.openai.com/docs)
- [Matplotlib Documentation](https://matplotlib.org)
- [Pandas Documentation](https://pandas.pydata.org)

## 🗺️ Roadmap

- [ ] CSV file support
- [ ] More chart types (scatter, histogram, heatmap)
- [ ] Data validation and cleaning tools
- [ ] Custom color palettes
- [ ] Batch export functionality
- [ ] Chart templates library
- [ ] Real-time collaboration
- [ ] Dark mode support

## 📊 Version History

### v1.0.0 (Current)

- Initial release
- Support for Bar, Line, Pie charts
- AI assistant integration
- Excel export with styling
- Template system
- Dual-axis support
- Data editing capabilities

---

**Made with ❤️ using Streamlit, OpenAI, and Python**
