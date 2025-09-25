"""
Streamlit Web Interface for VSDX Extraction Tool
"""

import streamlit as st
import os
import json
import zipfile
from pathlib import Path
import tempfile
import shutil
from vsdx_extractor import VSDXExtractor
from gemini_integration import GeminiAnalyzer
from visualization import create_extraction_visualization
import pandas as pd

# Configure Streamlit page
st.set_page_config(
    page_title="VSDX Extraction Tool",
    page_icon="📊",
    layout="wide"
)

def main():
    st.title("📊 VSDX Extraction Tool")
    st.markdown("Upload a VSDX (Visio) file to extract its contents and analyze with AI")
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("Configuration")
        
        # Gemini API key input
        gemini_api_key = st.text_input(
            "Gemini API Key (Optional)", 
            type="password",
            help="Enter your Google Gemini API key for AI analysis"
        )
        
        if gemini_api_key:
            os.environ["GEMINI_API_KEY"] = gemini_api_key
            st.success("✅ Gemini API key configured")
        
        st.markdown("---")
        st.markdown("### About")
        st.markdown("""
        This tool extracts VSDX files and:
        - 📄 Extracts all sheets to XML files
        - 📊 Creates visualizations
        - 🤖 Provides AI analysis (with Gemini API)
        """)
    
    # Main content area
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.header("📁 File Upload")
        
        uploaded_file = st.file_uploader(
            "Choose a VSDX file",
            type=['vsdx'],
            help="Upload a Microsoft Visio (.vsdx) file"
        )
        
        if uploaded_file is not None:
            # Display file info
            st.info(f"📄 File: {uploaded_file.name} ({uploaded_file.size:,} bytes)")
            
            # Process button
            if st.button("🚀 Extract and Analyze", type="primary"):
                process_vsdx_file(uploaded_file, gemini_api_key)
    
    with col2:
        st.header("ℹ️ Instructions")
        st.markdown("""
        1. **Upload**: Select a VSDX file using the file uploader
        2. **Configure**: Optionally add your Gemini API key for AI analysis
        3. **Extract**: Click "Extract and Analyze" to process the file
        4. **Review**: Examine the extracted XML files and visualizations
        5. **Analyze**: Read the AI analysis if API key is provided
        """)
        
        st.markdown("### Supported Features")
        st.markdown("""
        - ✅ Extract all pages/sheets
        - ✅ Save individual XML files
        - ✅ Generate extraction summary
        - ✅ Create visualizations
        - ✅ AI-powered analysis
        - ✅ Download extracted files
        """)

def process_vsdx_file(uploaded_file, gemini_api_key):
    """Process the uploaded VSDX file"""
    
    # Create progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Step 1: Save uploaded file temporarily
        status_text.text("📁 Saving uploaded file...")
        progress_bar.progress(10)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.vsdx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            temp_vsdx_path = tmp_file.name
        
        # Step 2: Extract VSDX
        status_text.text("🔧 Extracting VSDX file...")
        progress_bar.progress(30)
        
        extractor = VSDXExtractor()
        result = extractor.extract_vsdx(temp_vsdx_path)
        
        if not result['success']:
            st.error(f"❌ Extraction failed: {result['error']}")
            return
        
        # Step 3: Display results
        status_text.text("📊 Creating visualizations...")
        progress_bar.progress(60)
        
        display_extraction_results(result, extractor)
        
        # Step 4: AI Analysis
        if gemini_api_key:
            status_text.text("🤖 Running AI analysis...")
            progress_bar.progress(80)
            
            run_ai_analysis(result, extractor, gemini_api_key)
        
        # Step 5: Complete
        status_text.text("✅ Processing complete!")
        progress_bar.progress(100)
        
        st.success("🎉 VSDX file processed successfully!")
        
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
    
    finally:
        # Clean up temporary file
        if 'temp_vsdx_path' in locals() and os.path.exists(temp_vsdx_path):
            os.unlink(temp_vsdx_path)

def display_extraction_results(result, extractor):
    """Display the extraction results"""
    
    st.header("📊 Extraction Results")
    
    # Summary metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Pages", len(result['pages']))
    
    with col2:
        st.metric("Output Directory", "✅ Created")
    
    with col3:
        total_elements = sum(page.get('elements_count', 0) for page in result['pages'])
        st.metric("Total Elements", total_elements)
    
    with col4:
        st.metric("Status", "✅ Success")
    
    # Pages information
    if result['pages']:
        st.subheader("📄 Pages/Sheets Extracted")
        
        pages_df = pd.DataFrame(result['pages'])
        st.dataframe(pages_df, use_container_width=True)
        
        # Create visualization
        fig = create_extraction_visualization(result['pages'])
        if fig:
            st.plotly_chart(fig, use_container_width=True)
    
    # Extracted data summary
    if extractor.extracted_data:
        st.subheader("📋 Extracted Data Summary")
        
        with st.expander("View Raw Extracted Data"):
            st.json(extractor.extracted_data)
    
    # File download section
    st.subheader("💾 Download Extracted Files")
    
    if st.button("📦 Create Download Package"):
        create_download_package(result['output_dir'])

def run_ai_analysis(result, extractor, api_key):
    """Run AI analysis using Gemini"""
    
    try:
        analyzer = GeminiAnalyzer(api_key)
        
        # Prepare data for analysis
        analysis_data = {
            'summary': extractor.get_extraction_summary(),
            'pages': result['pages']
        }
        
        # Run analysis
        analysis_result = analyzer.analyze_extraction(analysis_data)
        
        if analysis_result:
            st.header("🤖 AI Analysis")
            st.markdown(analysis_result)
    
    except Exception as e:
        st.warning(f"⚠️ AI analysis failed: {str(e)}")

def create_download_package(output_dir):
    """Create a ZIP package of extracted files for download"""
    
    try:
        # Create ZIP file in memory
        import io
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            # Add all files from output directory
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_name = os.path.relpath(file_path, output_dir)
                    zip_file.write(file_path, arc_name)
        
        zip_buffer.seek(0)
        
        # Provide download button
        st.download_button(
            label="📥 Download Extracted Files (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="vsdx_extracted_files.zip",
            mime="application/zip"
        )
        
        st.success("✅ Download package created successfully!")
        
    except Exception as e:
        st.error(f"❌ Error creating download package: {str(e)}")

if __name__ == "__main__":
    main()