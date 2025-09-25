#!/usr/bin/env python3
"""
Demo script to showcase VSDX extraction capabilities
"""

import os
import json
from vsdx_extractor import VSDXExtractor
from visualization import create_extraction_visualization, create_page_comparison_chart
from create_test_vsdx import create_test_vsdx
import matplotlib.pyplot as plt

def run_demo():
    """Run a complete demo of the VSDX extraction tool"""
    
    print("🎯 VSDX Extraction Tool Demo")
    print("=" * 50)
    
    # Step 1: Create a test VSDX file
    print("\n1. Creating test VSDX file...")
    test_file = create_test_vsdx("demo_diagram.vsdx")
    print(f"   ✅ Created: {test_file}")
    
    # Step 2: Extract the VSDX file
    print("\n2. Extracting VSDX file...")
    extractor = VSDXExtractor()
    result = extractor.extract_vsdx(test_file, "demo_output")
    
    if result['success']:
        print(f"   ✅ Extraction successful!")
        print(f"   📁 Output directory: {result['output_dir']}")
        print(f"   📄 Pages extracted: {len(result['pages'])}")
    else:
        print(f"   ❌ Extraction failed: {result['error']}")
        return
    
    # Step 3: Display extraction summary
    print("\n3. Extraction Summary:")
    summary = extractor.get_extraction_summary()
    
    print(f"   📊 Total pages: {summary['total_pages']}")
    print(f"   📋 App properties: {'✅' if summary['has_app_properties'] else '❌'}")
    print(f"   📄 Document info: {'✅' if summary['has_document_info'] else '❌'}")
    
    print("\n   Page Details:")
    for i, page in enumerate(summary['pages'], 1):
        print(f"   Page {i}: {page['name']} ({page['elements_count']} elements)")
    
    # Step 4: Create visualizations
    print("\n4. Creating visualizations...")
    
    try:
        # Create Plotly visualization
        fig = create_extraction_visualization(result['pages'])
        if fig:
            fig.write_html("demo_visualization.html")
            print("   ✅ Interactive visualization saved as 'demo_visualization.html'")
        
        # Create comparison chart
        comparison_fig = create_page_comparison_chart(result['pages'])
        if comparison_fig:
            comparison_fig.write_html("demo_comparison.html")
            print("   ✅ Page comparison chart saved as 'demo_comparison.html'")
            
    except Exception as e:
        print(f"   ⚠️ Visualization creation failed: {e}")
    
    # Step 5: Show extracted files
    print("\n5. Extracted Files:")
    
    if os.path.exists(result['output_dir']):
        for root, dirs, files in os.walk(result['output_dir']):
            level = root.replace(result['output_dir'], '').count(os.sep)
            indent = ' ' * 2 * level
            print(f"   {indent}📁 {os.path.basename(root)}/")
            subindent = ' ' * 2 * (level + 1)
            for file in files:
                print(f"   {subindent}📄 {file}")
    
    # Step 6: Display extracted content sample
    print("\n6. Sample Extracted Content:")
    
    app_props_file = os.path.join(result['output_dir'], 'app_properties.xml')
    if os.path.exists(app_props_file):
        print("   App Properties (first 200 chars):")
        with open(app_props_file, 'r', encoding='utf-8') as f:
            content = f.read()
            print(f"   {content[:200]}...")
    
    # Step 7: AI Analysis placeholder
    print("\n7. AI Analysis:")
    print("   🤖 To enable AI analysis, run the Streamlit app with your Gemini API key:")
    print("   📝 streamlit run app.py")
    print("   🔑 Add your Gemini API key in the sidebar")
    
    print("\n🎉 Demo completed successfully!")
    print("\nNext steps:")
    print("   • Run 'streamlit run app.py' for the web interface")
    print("   • Upload your own VSDX files")
    print("   • Add Gemini API key for AI analysis")
    print("   • Open demo_visualization.html in your browser")
    
    # Cleanup
    print("\n🧹 Cleaning up demo files...")
    import shutil
    try:
        if os.path.exists("demo_diagram.vsdx"):
            os.remove("demo_diagram.vsdx")
        if os.path.exists("demo_output"):
            shutil.rmtree("demo_output")
        print("   ✅ Demo files cleaned up")
    except Exception as e:
        print(f"   ⚠️ Cleanup warning: {e}")

if __name__ == "__main__":
    run_demo()