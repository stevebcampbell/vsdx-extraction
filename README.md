# 📊 VSDX Extraction Tool

A Python tool to extract and analyze Microsoft Visio (.vsdx) files with AI-powered insights using Google Gemini.

## Features

- 📄 **Extract VSDX Files**: Extract all sheets/pages from VSDX files into individual XML files
- 📊 **Visualizations**: Create interactive charts and graphs showing extraction results  
- 🤖 **AI Analysis**: Get intelligent insights about your diagrams using Google Gemini API
- 🌐 **Web Interface**: Easy-to-use Streamlit web application
- 💾 **Export Results**: Download extracted files as ZIP packages

## Installation

1. Clone the repository:
```bash
git clone https://github.com/stevebcampbell/vsdx-extraction.git
cd vsdx-extraction
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Web Interface (Recommended)

Run the Streamlit app:
```bash
streamlit run app.py
```

Then open your browser to `http://localhost:8501` and:
1. Upload a VSDX file
2. Optionally add your Gemini API key for AI analysis
3. Click "Extract and Analyze"
4. View results and download extracted files

### Command Line

Extract a VSDX file using the command line:
```bash
python vsdx_extractor.py your_file.vsdx [output_directory]
```

## What Gets Extracted

The tool extracts and processes:
- **Pages/Sheets**: Each page/sheet becomes a separate XML file
- **Document Properties**: Application and document metadata
- **Master Shapes**: Stencils and templates used in the diagram
- **XML Structure**: Complete XML structure with element counts and analysis

## AI Analysis Features

With a Gemini API key, the tool provides:
- Document structure analysis
- Content type identification
- Complexity assessment
- Technical recommendations
- Comprehensive extraction reports

## File Structure

```
vsdx-extraction/
├── app.py                    # Streamlit web interface
├── vsdx_extractor.py        # Core extraction logic
├── gemini_integration.py    # AI analysis with Gemini
├── visualization.py         # Chart and graph creation
├── requirements.txt         # Python dependencies
└── README.md               # This file
```

## Technical Details

VSDX files are essentially ZIP archives containing XML files. This tool:
1. Extracts the ZIP content to a temporary directory
2. Parses XML files for pages, masters, and document properties
3. Saves individual XML files for each component
4. Creates visualizations of the extraction results
5. Uses AI to analyze and provide insights

## Requirements

- Python 3.7+
- Streamlit
- matplotlib/plotly for visualizations
- Google Generative AI (for Gemini integration)
- Standard libraries: zipfile, xml.etree.ElementTree, etc.

## Getting Gemini API Key

1. Go to [Google AI Studio](https://makersuite.google.com/app/apikey)
2. Create a new API key
3. Add it to the web interface or export as environment variable:
   ```bash
   export GEMINI_API_KEY="your-api-key-here"
   ```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is open source and available under the MIT License.