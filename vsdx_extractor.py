"""
VSDX Extraction Tool
A Python tool to extract and analyze VSDX (Visio) files
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import tempfile
import shutil

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class VSDXExtractor:
    """
    A class to extract and analyze VSDX files
    VSDX files are essentially ZIP archives containing XML files
    """
    
    def __init__(self):
        self.extracted_data = {}
        self.pages_info = []
        self.temp_dir = None
    
    def extract_vsdx(self, vsdx_path: str, output_dir: str = None) -> Dict:
        """
        Extract VSDX file contents and parse XML data
        
        Args:
            vsdx_path: Path to the VSDX file
            output_dir: Directory to save extracted XML files
            
        Returns:
            Dictionary containing extracted data and metadata
        """
        if not os.path.exists(vsdx_path):
            raise FileNotFoundError(f"VSDX file not found: {vsdx_path}")
        
        if output_dir is None:
            output_dir = f"{vsdx_path}_extracted"
        
        # Create output directory
        os.makedirs(output_dir, exist_ok=True)
        
        # Create temporary directory for extraction
        self.temp_dir = tempfile.mkdtemp()
        
        try:
            # Extract VSDX (ZIP) file
            with zipfile.ZipFile(vsdx_path, 'r') as zip_ref:
                zip_ref.extractall(self.temp_dir)
                logger.info(f"Extracted VSDX to temporary directory: {self.temp_dir}")
            
            # Process extracted files
            self._process_extracted_files(self.temp_dir, output_dir)
            
            return {
                'success': True,
                'output_dir': output_dir,
                'pages': self.pages_info,
                'extracted_data': self.extracted_data
            }
            
        except Exception as e:
            logger.error(f"Error extracting VSDX: {str(e)}")
            return {
                'success': False,
                'error': str(e)
            }
        finally:
            # Clean up temporary directory
            if self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
    
    def _process_extracted_files(self, temp_dir: str, output_dir: str):
        """
        Process the extracted files and save relevant XML files
        """
        # Key directories and files in VSDX structure
        visio_dir = os.path.join(temp_dir, 'visio')
        pages_dir = os.path.join(visio_dir, 'pages')
        
        # Process app.xml for document properties
        app_xml_path = os.path.join(temp_dir, 'docProps', 'app.xml')
        if os.path.exists(app_xml_path):
            self._process_app_xml(app_xml_path, output_dir)
        
        # Process document.xml for document structure
        doc_xml_path = os.path.join(visio_dir, 'document.xml')
        if os.path.exists(doc_xml_path):
            self._process_document_xml(doc_xml_path, output_dir)
        
        # Process pages
        if os.path.exists(pages_dir):
            self._process_pages(pages_dir, output_dir)
        
        # Process masters (stencils/templates)
        masters_dir = os.path.join(visio_dir, 'masters')
        if os.path.exists(masters_dir):
            self._process_masters(masters_dir, output_dir)
    
    def _process_app_xml(self, app_xml_path: str, output_dir: str):
        """Process application properties XML"""
        try:
            tree = ET.parse(app_xml_path)
            root = tree.getroot()
            
            # Save processed XML
            output_path = os.path.join(output_dir, 'app_properties.xml')
            tree.write(output_path, encoding='utf-8', xml_declaration=True)
            
            # Extract basic properties
            self.extracted_data['app_properties'] = {}
            for child in root:
                if child.text:
                    tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    self.extracted_data['app_properties'][tag_name] = child.text
            
            logger.info("Processed app.xml")
            
        except Exception as e:
            logger.error(f"Error processing app.xml: {str(e)}")
    
    def _process_document_xml(self, doc_xml_path: str, output_dir: str):
        """Process main document XML"""
        try:
            tree = ET.parse(doc_xml_path)
            root = tree.getroot()
            
            # Save processed XML
            output_path = os.path.join(output_dir, 'document.xml')
            tree.write(output_path, encoding='utf-8', xml_declaration=True)
            
            # Extract document info
            self.extracted_data['document_info'] = {
                'root_tag': root.tag,
                'namespaces': dict(root.attrib)
            }
            
            logger.info("Processed document.xml")
            
        except Exception as e:
            logger.error(f"Error processing document.xml: {str(e)}")
    
    def _process_pages(self, pages_dir: str, output_dir: str):
        """Process all pages in the VSDX file"""
        pages_output_dir = os.path.join(output_dir, 'pages')
        os.makedirs(pages_output_dir, exist_ok=True)
        
        page_files = [f for f in os.listdir(pages_dir) if f.endswith('.xml')]
        
        for page_file in page_files:
            page_path = os.path.join(pages_dir, page_file)
            self._process_single_page(page_path, pages_output_dir, page_file)
    
    def _process_single_page(self, page_path: str, output_dir: str, page_filename: str):
        """Process a single page XML file"""
        try:
            tree = ET.parse(page_path)
            root = tree.getroot()
            
            # Save the page XML
            output_path = os.path.join(output_dir, page_filename)
            tree.write(output_path, encoding='utf-8', xml_declaration=True)
            
            # Extract page information
            page_info = {
                'filename': page_filename,
                'output_path': output_path,
                'elements_count': len(root.findall('.//*')),
                'root_tag': root.tag
            }
            
            # Try to extract page name and ID
            for elem in root.iter():
                if 'PageSheet' in elem.tag:
                    # Look for name property
                    for cell in elem.findall('.//Cell'):
                        if cell.get('N') == 'PageName':
                            page_info['name'] = cell.get('V', 'Unnamed')
                            break
                    break
            
            if 'name' not in page_info:
                page_info['name'] = page_filename.replace('.xml', '')
            
            self.pages_info.append(page_info)
            logger.info(f"Processed page: {page_filename}")
            
        except Exception as e:
            logger.error(f"Error processing page {page_filename}: {str(e)}")
    
    def _process_masters(self, masters_dir: str, output_dir: str):
        """Process master shapes/stencils"""
        masters_output_dir = os.path.join(output_dir, 'masters')
        os.makedirs(masters_output_dir, exist_ok=True)
        
        master_files = [f for f in os.listdir(masters_dir) if f.endswith('.xml')]
        
        for master_file in master_files:
            master_path = os.path.join(masters_dir, master_file)
            try:
                tree = ET.parse(master_path)
                
                # Save the master XML
                output_path = os.path.join(masters_output_dir, master_file)
                tree.write(output_path, encoding='utf-8', xml_declaration=True)
                
                logger.info(f"Processed master: {master_file}")
                
            except Exception as e:
                logger.error(f"Error processing master {master_file}: {str(e)}")
    
    def get_extraction_summary(self) -> Dict:
        """Get a summary of the extraction results"""
        return {
            'total_pages': len(self.pages_info),
            'pages': self.pages_info,
            'has_app_properties': 'app_properties' in self.extracted_data,
            'has_document_info': 'document_info' in self.extracted_data,
            'extracted_data': self.extracted_data
        }

def main():
    """Main function for command line usage"""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python vsdx_extractor.py <vsdx_file> [output_dir]")
        return
    
    vsdx_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    extractor = VSDXExtractor()
    result = extractor.extract_vsdx(vsdx_file, output_dir)
    
    if result['success']:
        print(f"Extraction successful!")
        print(f"Output directory: {result['output_dir']}")
        print(f"Pages extracted: {len(result['pages'])}")
        
        summary = extractor.get_extraction_summary()
        print("\nSummary:")
        print(json.dumps(summary, indent=2))
    else:
        print(f"Extraction failed: {result['error']}")

if __name__ == "__main__":
    main()