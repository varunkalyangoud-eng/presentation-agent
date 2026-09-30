# Presentation Agent 🎨

An AI-powered presentation generation platform that transforms your data (PDF, Excel, Word, Images, Text) into stunning presentations, reports, visualizations, and infographics.

## 🎯 Features

### 📊 Multi-Format Data Input
- **PDF Documents** - Extract text, tables, and insights
- **Excel Spreadsheets** - Process data and generate analytics
- **Word Documents** - Parse content and structure
- **Images & Photos** - Extract visual context with OCR
- **Text Files** - Process raw text data

### 📈 Output Formats
- **PowerPoint Slides** - Professional PPTX presentations
- **PDF Reports** - Executive summaries and detailed reports
- **Web-Based Presentations** - Interactive HTML5 slideshows
- **Power BI Dashboards** - Interactive business intelligence
- **Tableau Visualizations** - Advanced analytics dashboards
- **Infographics** - Visually stunning data visualizations

### 🎨 Design Features
- **Attractive Infographics** - Auto-generated data visualizations
- **Professional Templates** - Multiple design themes
- **Brand Customization** - Logo, colors, fonts
- **Responsive Design** - Mobile-friendly presentations
- **Interactive Charts** - D3.js, Chart.js, Plotly visualizations

### 📋 Capabilities
- **Data Analysis** - Automatic insights and trends
- **Business Reporting** - Executive summaries
- **Storytelling** - Narrative-driven presentations
- **Sentiment Analysis** - Text and content analysis
- **Predictive Analytics** - Trend forecasting

## 🏗️ Architecture

```
presentation-agent/
├── backend/                 # Python Flask/FastAPI backend
│   ├── app.py
│   ├── parsers/            # Document parsing modules
│   ├── analyzers/          # Data analysis engines
│   ├── generators/         # Slide/report generators
│   ├── exporters/          # PowerPoint, PDF, Power BI exporters
│   ├── visualizers/        # Infographic generation
│   └── requirements.txt
├── frontend/               # React/Vue.js web interface
│   ├── src/
│   ├── components/
│   ├── pages/
│   └── package.json
├── templates/              # Presentation templates
│   ├── powerpoint/
│   ├── html/
│   └── themes/
├── config/                 # Configuration files
│   ├── powerbi_config.json
│   ├── tableau_config.json
│   └── settings.py
└── docker-compose.yml      # Docker orchestration
```

## 🚀 Quick Start

### Prerequisites
- Python 3.9+
- Node.js 14+
- Docker & Docker Compose
- Power BI Desktop (optional)
- Tableau Desktop (optional)

### Installation

```bash
# Clone the repository
git clone https://github.com/varunkalyangoud-eng/presentation-agent.git
cd presentation-agent

# Setup backend
cd backend
pip install -r requirements.txt

# Setup frontend
cd ../frontend
npm install

# Run with Docker
docker-compose up -d
```

### Usage

1. **Upload Data**: Navigate to web interface and upload files
2. **Configure Output**: Select output formats (PowerPoint, PDF, Power BI, Tableau, Infographics)
3. **Customize**: Choose templates and design themes
4. **Generate**: Let AI create your presentation
5. **Export**: Download or share your presentation

## 📊 Supported Visualizations

- Bar Charts & Histograms
- Line & Area Charts
- Pie & Donut Charts
- Scatter Plots & Bubble Charts
- Heatmaps & Correlation matrices
- Sankey Diagrams
- Tree Maps & Sunburst Charts
- Geographic Maps
- Custom Infographics

## 🔌 Integrations

- **Power BI** - Native dataset import & dashboard creation
- **Tableau** - Data source connection & viz generation
- **PowerPoint** - python-pptx automation
- **PDF** - ReportLab and FPDF2
- **OCR** - Tesseract & PyTorch for image analysis
- **AI** - GPT/Claude for content generation & insights

## 📝 Configuration

Edit `config/settings.py`:
```python
POWERPOINT_TEMPLATE = "modern_blue"
TABLEAU_SERVER_URL = "your-tableau-url"
POWERBI_WORKSPACE_ID = "your-workspace-id"
INFOGRAPHIC_STYLE = "material"  # material, flat, 3d, minimalist
```

## 🎨 Templates

- Modern Blue
- Corporate Green
- Startup Minimalist
- Executive Classic
- Creative Colorful
- Dark Mode Professional
- Eco-Friendly Green
- Tech Innovation

## 📚 Documentation

- [Backend API Docs](./backend/README.md)
- [Frontend Guide](./frontend/README.md)
- [Power BI Integration](./docs/powerbi-setup.md)
- [Tableau Integration](./docs/tableau-setup.md)
- [Custom Templates](./docs/custom-templates.md)

## 🤝 Contributing

Contributions welcome! See [CONTRIBUTING.md](CONTRIBUTING.md)

## 📄 License

MIT License - see [LICENSE](LICENSE)

## 💬 Support

Issues? Create a [GitHub Issue](https://github.com/varunkalyangoud-eng/presentation-agent/issues)

---

**Made with ❤️ for data storytellers**
