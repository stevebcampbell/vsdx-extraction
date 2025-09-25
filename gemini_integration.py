"""
Gemini AI Integration for VSDX Analysis
"""

import google.generativeai as genai
import json
import logging
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

class GeminiAnalyzer:
    """
    A class to integrate with Google's Gemini AI for analyzing VSDX extraction results
    """
    
    def __init__(self, api_key: str):
        """
        Initialize the Gemini analyzer
        
        Args:
            api_key: Google Gemini API key
        """
        if not api_key:
            raise ValueError("Gemini API key is required")
        
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel('gemini-pro')
        
        logger.info("Gemini AI configured successfully")
    
    def analyze_extraction(self, extraction_data: Dict) -> Optional[str]:
        """
        Analyze the VSDX extraction results using Gemini AI
        
        Args:
            extraction_data: Dictionary containing extraction results and metadata
            
        Returns:
            AI analysis as a string, or None if analysis fails
        """
        try:
            # Prepare the prompt for Gemini
            prompt = self._create_analysis_prompt(extraction_data)
            
            # Generate analysis
            response = self.model.generate_content(prompt)
            
            if response and response.text:
                logger.info("Gemini analysis completed successfully")
                return response.text
            else:
                logger.warning("Gemini returned empty response")
                return None
                
        except Exception as e:
            logger.error(f"Error during Gemini analysis: {str(e)}")
            return None
    
    def _create_analysis_prompt(self, extraction_data: Dict) -> str:
        """
        Create a detailed prompt for Gemini AI analysis
        
        Args:
            extraction_data: Extraction results and metadata
            
        Returns:
            Formatted prompt string
        """
        summary = extraction_data.get('summary', {})
        pages = extraction_data.get('pages', [])
        
        prompt = f"""
        Analyze the following VSDX (Microsoft Visio) file extraction results and provide insights:

        **Extraction Summary:**
        - Total Pages: {summary.get('total_pages', 0)}
        - Has Application Properties: {summary.get('has_app_properties', False)}
        - Has Document Info: {summary.get('has_document_info', False)}

        **Page Details:**
        """
        
        for i, page in enumerate(pages, 1):
            prompt += f"""
        Page {i}:
        - Name: {page.get('name', 'Unknown')}
        - Filename: {page.get('filename', 'Unknown')}
        - Elements Count: {page.get('elements_count', 0)}
        - Root Tag: {page.get('root_tag', 'Unknown')}
        """
        
        prompt += """

        **Analysis Request:**
        Please provide a comprehensive analysis covering:

        1. **Document Structure Analysis:**
           - Overall document organization
           - Page complexity assessment
           - Element distribution patterns

        2. **Content Insights:**
           - What type of diagrams or content this appears to be
           - Complexity level of each page
           - Potential use cases or domains

        3. **Technical Assessment:**
           - XML structure quality
           - Data completeness
           - Any potential issues or anomalies

        4. **Recommendations:**
           - Best practices for working with this extracted data
           - Potential next steps for analysis
           - Tools or techniques that might be useful

        5. **Summary:**
           - Key takeaways about this VSDX file
           - Overall assessment of the extraction success

        Please format your response in clear sections with markdown formatting for readability.
        """
        
        return prompt
    
    def analyze_page_content(self, page_data: Dict, xml_content: str = None) -> Optional[str]:
        """
        Analyze a specific page's content in detail
        
        Args:
            page_data: Page metadata
            xml_content: Raw XML content of the page (optional)
            
        Returns:
            AI analysis of the specific page
        """
        try:
            prompt = f"""
            Analyze this specific page from a VSDX file:

            **Page Information:**
            - Name: {page_data.get('name', 'Unknown')}
            - Elements Count: {page_data.get('elements_count', 0)}
            - Filename: {page_data.get('filename', 'Unknown')}

            """
            
            if xml_content:
                # Truncate XML content if too long
                if len(xml_content) > 5000:
                    xml_content = xml_content[:5000] + "... [truncated]"
                
                prompt += f"""
            **XML Content Sample:**
            ```xml
            {xml_content}
            ```
            """
            
            prompt += """
            Please analyze this page and provide:
            1. What type of content this page likely contains
            2. Complexity assessment
            3. Key elements or patterns identified
            4. Potential insights about the diagram structure
            
            Keep the analysis concise but informative.
            """
            
            response = self.model.generate_content(prompt)
            
            if response and response.text:
                return response.text
            else:
                return None
                
        except Exception as e:
            logger.error(f"Error analyzing page content: {str(e)}")
            return None
    
    def generate_extraction_report(self, extraction_data: Dict) -> Optional[str]:
        """
        Generate a comprehensive report of the extraction process
        
        Args:
            extraction_data: Complete extraction results
            
        Returns:
            Formatted report as markdown string
        """
        try:
            analysis = self.analyze_extraction(extraction_data)
            
            if not analysis:
                return None
            
            # Create a structured report
            report = f"""
# VSDX Extraction Report

## 📊 Extraction Summary
- **Total Pages Extracted:** {extraction_data['summary'].get('total_pages', 0)}
- **Processing Status:** ✅ Successful
- **Analysis Generated:** {len(analysis.split()) if analysis else 0} words

## 🤖 AI Analysis

{analysis}

## 📋 Technical Details

### Pages Processed:
"""
            
            for page in extraction_data.get('pages', []):
                report += f"""
- **{page.get('name', 'Unnamed')}**
  - File: `{page.get('filename', 'unknown.xml')}`
  - Elements: {page.get('elements_count', 0)}
"""
            
            report += """

---
*Report generated using Google Gemini AI*
"""
            
            return report
            
        except Exception as e:
            logger.error(f"Error generating extraction report: {str(e)}")
            return None

def test_gemini_connection(api_key: str) -> bool:
    """
    Test the connection to Gemini AI
    
    Args:
        api_key: Gemini API key to test
        
    Returns:
        True if connection successful, False otherwise
    """
    try:
        analyzer = GeminiAnalyzer(api_key)
        
        # Simple test prompt
        test_prompt = "Hello, can you confirm this connection is working? Please respond with 'Connection successful'."
        response = analyzer.model.generate_content(test_prompt)
        
        if response and response.text and "successful" in response.text.lower():
            return True
        else:
            return False
            
    except Exception as e:
        logger.error(f"Gemini connection test failed: {str(e)}")
        return False