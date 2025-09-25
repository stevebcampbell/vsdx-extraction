"""
Create a simple test VSDX file for demonstration purposes
"""

import zipfile
import os
from pathlib import Path

def create_test_vsdx(filename="test_diagram.vsdx"):
    """
    Create a minimal test VSDX file with basic structure
    """
    
    test_dir = "test_vsdx_content"
    os.makedirs(test_dir, exist_ok=True)
    
    # Create directory structure
    os.makedirs(f"{test_dir}/docProps", exist_ok=True)
    os.makedirs(f"{test_dir}/visio", exist_ok=True)
    os.makedirs(f"{test_dir}/visio/pages", exist_ok=True)
    os.makedirs(f"{test_dir}/_rels", exist_ok=True)
    
    # Create app.xml (document properties)
    app_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
    <Application>Microsoft Visio</Application>
    <ScaleCrop>false</ScaleCrop>
    <Company>Test Company</Company>
    <LinksUpToDate>false</LinksUpToDate>
    <SharedDoc>false</SharedDoc>
    <HyperlinksChanged>false</HyperlinksChanged>
    <AppVersion>16.0000</AppVersion>
</Properties>"""
    
    with open(f"{test_dir}/docProps/app.xml", "w", encoding="utf-8") as f:
        f.write(app_xml)
    
    # Create document.xml
    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<VisioDocument xmlns="http://schemas.microsoft.com/office/visio/2012/main">
    <DocumentSettings>
        <GlueSettings>9</GlueSettings>
        <SnapSettings>65847</SnapSettings>
        <SnapExtensions>26</SnapExtensions>
    </DocumentSettings>
    <Pages>
        <Page ID="0" NameU="Page-1" Name="Page-1" ViewScale="1" ViewCenterX="4.25" ViewCenterY="5.5">
            <PageSheet>
                <Cell N="PageWidth" V="8.5" U="IN"/>
                <Cell N="PageHeight" V="11" U="IN"/>
                <Cell N="ShdwOffsetX" V="0.125" U="IN"/>
                <Cell N="ShdwOffsetY" V="-0.125" U="IN"/>
                <Cell N="PageScale" V="1" U="IN_F"/>
                <Cell N="DrawingScale" V="1" U="IN_F"/>
            </PageSheet>
        </Page>
        <Page ID="1" NameU="Page-2" Name="Page-2" ViewScale="1" ViewCenterX="4.25" ViewCenterY="5.5">
            <PageSheet>
                <Cell N="PageWidth" V="8.5" U="IN"/>
                <Cell N="PageHeight" V="11" U="IN"/>
                <Cell N="ShdwOffsetX" V="0.125" U="IN"/>
                <Cell N="ShdwOffsetY" V="-0.125" U="IN"/>
                <Cell N="PageScale" V="1" U="IN_F"/>
                <Cell N="DrawingScale" V="1" U="IN_F"/>
            </PageSheet>
        </Page>
    </Pages>
</VisioDocument>"""
    
    with open(f"{test_dir}/visio/document.xml", "w", encoding="utf-8") as f:
        f.write(document_xml)
    
    # Create page1.xml
    page1_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<PageContents xmlns="http://schemas.microsoft.com/office/visio/2012/main">
    <Shapes>
        <Shape ID="1" Type="Shape" LineStyle="3" FillStyle="3" TextStyle="3">
            <Cell N="PinX" V="2.25"/>
            <Cell N="PinY" V="8.5"/>
            <Cell N="Width" V="1.5"/>
            <Cell N="Height" V="1"/>
            <Text>Rectangle Shape</Text>
        </Shape>
        <Shape ID="2" Type="Shape" LineStyle="3" FillStyle="3" TextStyle="3">
            <Cell N="PinX" V="6"/>
            <Cell N="PinY" V="6"/>
            <Cell N="Width" V="2"/>
            <Cell N="Height" V="1.5"/>
            <Text>Oval Shape</Text>
        </Shape>
        <Shape ID="3" Type="Shape" LineStyle="4" FillStyle="0" TextStyle="3">
            <Cell N="PinX" V="4.25"/>
            <Cell N="PinY" V="3.5"/>
            <Cell N="Width" V="3"/>
            <Cell N="Height" V="0.25"/>
            <Text>Connector Line</Text>
        </Shape>
    </Shapes>
    <Connects>
        <Connect FromSheet="3" FromCell="BeginX" ToSheet="1" ToCell="PinX"/>
        <Connect FromSheet="3" FromCell="EndX" ToSheet="2" ToCell="PinX"/>
    </Connects>
</PageContents>"""
    
    with open(f"{test_dir}/visio/pages/page1.xml", "w", encoding="utf-8") as f:
        f.write(page1_xml)
    
    # Create page2.xml
    page2_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<PageContents xmlns="http://schemas.microsoft.com/office/visio/2012/main">
    <Shapes>
        <Shape ID="4" Type="Shape" LineStyle="5" FillStyle="5" TextStyle="3">
            <Cell N="PinX" V="3"/>
            <Cell N="PinY" V="7"/>
            <Cell N="Width" V="2"/>
            <Cell N="Height" V="1"/>
            <Text>Process Box</Text>
        </Shape>
        <Shape ID="5" Type="Shape" LineStyle="6" FillStyle="6" TextStyle="3">
            <Cell N="PinX" V="6"/>
            <Cell N="PinY" V="4"/>
            <Cell N="Width" V="1.5"/>
            <Cell N="Height" V="1.5"/>
            <Text>Decision Diamond</Text>
        </Shape>
        <Shape ID="6" Type="Shape" LineStyle="2" FillStyle="2" TextStyle="3">
            <Cell N="PinX" V="2"/>
            <Cell N="PinY" V="2"/>
            <Cell N="Width" V="4"/>
            <Cell N="Height" V="0.5"/>
            <Text>Data Storage</Text>
        </Shape>
    </Shapes>
</PageContents>"""
    
    with open(f"{test_dir}/visio/pages/page2.xml", "w", encoding="utf-8") as f:
        f.write(page2_xml)
    
    # Create [Content_Types].xml
    content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    <Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.drawing.main+xml"/>
    <Override PartName="/visio/pages/page1.xml" ContentType="application/vnd.ms-visio.page+xml"/>
    <Override PartName="/visio/pages/page2.xml" ContentType="application/vnd.ms-visio.page+xml"/>
</Types>"""
    
    with open(f"{test_dir}/[Content_Types].xml", "w", encoding="utf-8") as f:
        f.write(content_types_xml)
    
    # Create the VSDX file (ZIP archive)
    with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(test_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, test_dir)
                zipf.write(file_path, arcname)
    
    # Clean up temporary directory
    import shutil
    shutil.rmtree(test_dir)
    
    print(f"✅ Created test VSDX file: {filename}")
    return filename

if __name__ == "__main__":
    create_test_vsdx()