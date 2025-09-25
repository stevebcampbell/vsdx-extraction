# VSDX Extraction Tool - Usage Guide

## Quick Start

### 1. Web Interface (Recommended)

The easiest way to use the tool is through the web interface:

```bash
streamlit run app.py
```

Then open your browser to `http://localhost:8501` and:
1. Upload a VSDX file
2. Optionally add your Gemini API key for AI analysis  
3. Click "Extract and Analyze"
4. View results and download extracted files

### 2. Command Line

Extract a VSDX file directly from the command line:

```bash
python vsdx_extractor.py your_file.vsdx [output_directory]
```

### 3. Demo Mode

Run the demo to see all features:

```bash
python demo.py
```

This will create a sample VSDX file, extract it, and show visualizations.

## Features Overview

### ✅ What Gets Extracted

- **Pages/Sheets**: Each page becomes a separate XML file
- **Document Properties**: Application metadata (creator, version, etc.)
- **Master Shapes**: Stencils and templates used
- **XML Structure**: Complete structure with element counts

### ✅ Visualizations Created

- **Page Overview**: Bar charts showing element counts per page
- **Complexity Analysis**: Element distribution and trends
- **Interactive Charts**: Hover details and zoom capabilities
- **Export Options**: HTML, PNG, SVG formats

### ✅ AI Analysis (with Gemini API)

- Document structure analysis
- Content type identification  
- Complexity assessment
- Technical recommendations
- Comprehensive reports

## Example Output Structure

```
your_file_extracted/
├── app_properties.xml      # Document metadata
├── document.xml           # Main document structure
├── pages/                 # Individual page files
│   ├── page1.xml
│   ├── page2.xml
│   └── ...
└── masters/               # Master shapes (if any)
    ├── master1.xml
    └── ...
```

## Python API Usage

```python
from vsdx_extractor import VSDXExtractor
from visualization import create_extraction_visualization

# Initialize extractor
extractor = VSDXExtractor()

# Extract VSDX file
result = extractor.extract_vsdx("my_diagram.vsdx", "output_dir")

if result['success']:
    print(f"Extracted {len(result['pages'])} pages")
    
    # Get summary
    summary = extractor.get_extraction_summary()
    
    # Create visualization
    fig = create_extraction_visualization(result['pages'])
    if fig:
        fig.write_html("visualization.html")
        
else:
    print(f"Error: {result['error']}")
```

## Gemini AI Integration

To use AI analysis features:

1. Get API key from [Google AI Studio](https://makersuite.google.com/app/apikey)
2. Add to web interface or set environment variable:
   ```bash
   export GEMINI_API_KEY="your-api-key-here"
   ```

Example AI analysis output:
- Document type identification (flowchart, network diagram, etc.)
- Complexity scoring per page
- Recommendations for optimization
- Technical insights about the content

## Troubleshooting

### Common Issues

**1. "ModuleNotFoundError" errors**
```bash
pip install -r requirements.txt
```

**2. VSDX file not found**
- Check file path is correct
- Ensure file has .vsdx extension
- Verify file is not corrupted

**3. Empty extraction results**
- File might be password protected
- File might be an older .vsd format (not supported)
- File might be corrupted

**4. Visualization errors**
- Check if output directory was created
- Ensure pages were extracted successfully
- Try running demo.py to test basic functionality

### Getting Help

1. Run `python demo.py` to test basic functionality
2. Check the logs for error messages
3. Ensure all dependencies are installed
4. Try with a simple VSDX file first

## Advanced Usage

### Custom Output Processing

```python
# Process specific pages only
def process_page_content(page_path):
    import xml.etree.ElementTree as ET
    tree = ET.parse(page_path)
    root = tree.getroot()
    
    # Custom processing logic here
    shapes = root.findall('.//Shape')
    return len(shapes)

# Use after extraction
for page in result['pages']:
    shape_count = process_page_content(page['output_path'])
    print(f"Page {page['name']}: {shape_count} shapes")
```

### Batch Processing

```python
import os
from pathlib import Path

def batch_extract(input_dir, output_base_dir):
    """Extract multiple VSDX files"""
    vsdx_files = Path(input_dir).glob("*.vsdx")
    
    for vsdx_file in vsdx_files:
        extractor = VSDXExtractor()
        output_dir = os.path.join(output_base_dir, vsdx_file.stem)
        result = extractor.extract_vsdx(str(vsdx_file), output_dir)
        
        if result['success']:
            print(f"✅ {vsdx_file.name}: {len(result['pages'])} pages")
        else:
            print(f"❌ {vsdx_file.name}: {result['error']}")

# Usage
batch_extract("input_vsdx_files/", "extracted_results/")
```