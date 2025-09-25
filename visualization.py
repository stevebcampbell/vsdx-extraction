"""
Visualization module for VSDX extraction results
"""

import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import pandas as pd
import matplotlib.pyplot as plt
from typing import Dict, List, Optional
import logging

# Optional seaborn import
try:
    import seaborn as sns
    HAS_SEABORN = True
except ImportError:
    HAS_SEABORN = False

logger = logging.getLogger(__name__)

def create_extraction_visualization(pages_data: List[Dict]) -> Optional[go.Figure]:
    """
    Create interactive visualizations for VSDX extraction results
    
    Args:
        pages_data: List of page information dictionaries
        
    Returns:
        Plotly figure object or None if creation fails
    """
    try:
        if not pages_data:
            logger.warning("No pages data provided for visualization")
            return None
        
        # Convert to DataFrame for easier manipulation
        df = pd.DataFrame(pages_data)
        
        # Create subplots
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=[
                "Pages Overview",
                "Elements per Page", 
                "Page Size Distribution",
                "File Type Analysis"
            ],
            specs=[[{"type": "bar"}, {"type": "scatter"}],
                   [{"type": "histogram"}, {"type": "pie"}]]
        )
        
        # 1. Pages Overview Bar Chart
        fig.add_trace(
            go.Bar(
                x=df['name'] if 'name' in df.columns else df.index,
                y=df['elements_count'] if 'elements_count' in df.columns else [1] * len(df),
                name="Elements Count",
                marker_color='lightblue'
            ),
            row=1, col=1
        )
        
        # 2. Elements per Page Scatter Plot
        if 'elements_count' in df.columns:
            fig.add_trace(
                go.Scatter(
                    x=df.index,
                    y=df['elements_count'],
                    mode='markers+lines',
                    name="Elements Trend",
                    marker=dict(size=10, color='orange')
                ),
                row=1, col=2
            )
        
        # 3. Elements Distribution Histogram
        if 'elements_count' in df.columns:
            fig.add_trace(
                go.Histogram(
                    x=df['elements_count'],
                    name="Element Distribution",
                    marker_color='lightgreen'
                ),
                row=2, col=1
            )
        
        # 4. File Type Pie Chart
        file_extensions = df['filename'].str.extract(r'\.(\w+)$')[0].value_counts() if 'filename' in df.columns else pd.Series({'xml': len(df)})
        
        fig.add_trace(
            go.Pie(
                labels=file_extensions.index,
                values=file_extensions.values,
                name="File Types"
            ),
            row=2, col=2
        )
        
        # Update layout
        fig.update_layout(
            title="VSDX Extraction Analysis Dashboard",
            showlegend=True,
            height=800,
            template="plotly_white"
        )
        
        # Update x-axis labels
        fig.update_xaxes(title_text="Pages", row=1, col=1)
        fig.update_xaxes(title_text="Page Index", row=1, col=2)
        fig.update_xaxes(title_text="Element Count", row=2, col=1)
        
        # Update y-axis labels
        fig.update_yaxes(title_text="Element Count", row=1, col=1)
        fig.update_yaxes(title_text="Element Count", row=1, col=2)
        fig.update_yaxes(title_text="Frequency", row=2, col=1)
        
        logger.info("Visualization created successfully")
        return fig
        
    except Exception as e:
        logger.error(f"Error creating visualization: {str(e)}")
        return None

def create_page_comparison_chart(pages_data: List[Dict]) -> Optional[go.Figure]:
    """
    Create a comparison chart for pages
    
    Args:
        pages_data: List of page information dictionaries
        
    Returns:
        Plotly figure for page comparison
    """
    try:
        if not pages_data:
            return None
        
        df = pd.DataFrame(pages_data)
        
        fig = go.Figure()
        
        # Bar chart comparing pages
        fig.add_trace(
            go.Bar(
                x=df['name'] if 'name' in df.columns else [f"Page {i+1}" for i in range(len(df))],
                y=df['elements_count'] if 'elements_count' in df.columns else [1] * len(df),
                text=df['elements_count'] if 'elements_count' in df.columns else [1] * len(df),
                textposition='auto',
                marker_color='steelblue'
            )
        )
        
        fig.update_layout(
            title="Page Complexity Comparison",
            xaxis_title="Page Name",
            yaxis_title="Number of Elements",
            template="plotly_white",
            height=500
        )
        
        return fig
        
    except Exception as e:
        logger.error(f"Error creating page comparison chart: {str(e)}")
        return None

def create_extraction_summary_table(extraction_data: Dict) -> Optional[pd.DataFrame]:
    """
    Create a summary table of extraction results
    
    Args:
        extraction_data: Complete extraction data
        
    Returns:
        Pandas DataFrame with summary information
    """
    try:
        pages = extraction_data.get('pages', [])
        
        if not pages:
            return None
        
        # Create summary data
        summary_data = []
        
        for i, page in enumerate(pages, 1):
            summary_data.append({
                'Page #': i,
                'Page Name': page.get('name', 'Unnamed'),
                'Filename': page.get('filename', 'unknown.xml'),
                'Elements': page.get('elements_count', 0),
                'Root Tag': page.get('root_tag', 'Unknown'),
                'Status': '✅ Extracted'
            })
        
        df = pd.DataFrame(summary_data)
        
        # Add total row
        total_row = {
            'Page #': 'TOTAL',
            'Page Name': f"{len(pages)} pages",
            'Filename': f"{len(pages)} files",
            'Elements': sum(page.get('elements_count', 0) for page in pages),
            'Root Tag': 'Various',
            'Status': '✅ Complete'
        }
        
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
        
        return df
        
    except Exception as e:
        logger.error(f"Error creating summary table: {str(e)}")
        return None

def create_matplotlib_visualization(pages_data: List[Dict]) -> Optional[plt.Figure]:
    """
    Create matplotlib visualization as alternative to Plotly
    
    Args:
        pages_data: List of page information dictionaries
        
    Returns:
        Matplotlib figure object
    """
    try:
        if not pages_data:
            return None
        
        df = pd.DataFrame(pages_data)
        
        # Set up the matplotlib figure
        fig, axes = plt.subplots(2, 2, figsize=(15, 10))
        fig.suptitle('VSDX Extraction Analysis', fontsize=16)
        
        # 1. Bar chart of elements per page
        if 'name' in df.columns and 'elements_count' in df.columns:
            axes[0, 0].bar(range(len(df)), df['elements_count'], color='skyblue')
            axes[0, 0].set_title('Elements per Page')
            axes[0, 0].set_xlabel('Page Index')
            axes[0, 0].set_ylabel('Element Count')
            axes[0, 0].set_xticks(range(len(df)))
            axes[0, 0].set_xticklabels([f"P{i+1}" for i in range(len(df))], rotation=45)
        
        # 2. Line plot of element trend
        if 'elements_count' in df.columns:
            axes[0, 1].plot(df.index, df['elements_count'], marker='o', color='orange')
            axes[0, 1].set_title('Element Count Trend')
            axes[0, 1].set_xlabel('Page Index')
            axes[0, 1].set_ylabel('Element Count')
        
        # 3. Histogram of element distribution
        if 'elements_count' in df.columns:
            axes[1, 0].hist(df['elements_count'], bins=max(5, len(df)//2), color='lightgreen', alpha=0.7)
            axes[1, 0].set_title('Element Count Distribution')
            axes[1, 0].set_xlabel('Element Count')
            axes[1, 0].set_ylabel('Frequency')
        
        # 4. Pie chart of file types
        if 'filename' in df.columns:
            file_extensions = df['filename'].str.extract(r'\.(\w+)$')[0].value_counts()
            axes[1, 1].pie(file_extensions.values, labels=file_extensions.index, autopct='%1.1f%%')
            axes[1, 1].set_title('File Types')
        
        plt.tight_layout()
        return fig
        
    except Exception as e:
        logger.error(f"Error creating matplotlib visualization: {str(e)}")
        return None

def export_visualization(fig, filename: str, format: str = 'png'):
    """
    Export visualization to file
    
    Args:
        fig: Figure object (Plotly or Matplotlib)
        filename: Output filename
        format: Export format ('png', 'html', 'svg', etc.)
    """
    try:
        if hasattr(fig, 'write_html'):  # Plotly figure
            if format.lower() == 'html':
                fig.write_html(filename)
            else:
                fig.write_image(filename)
        elif hasattr(fig, 'savefig'):  # Matplotlib figure
            fig.savefig(filename, format=format, dpi=300, bbox_inches='tight')
        
        logger.info(f"Visualization exported to {filename}")
        
    except Exception as e:
        logger.error(f"Error exporting visualization: {str(e)}")

def create_interactive_dashboard(extraction_data: Dict) -> Optional[go.Figure]:
    """
    Create an interactive dashboard combining multiple visualizations
    
    Args:
        extraction_data: Complete extraction data
        
    Returns:
        Interactive Plotly dashboard
    """
    try:
        pages = extraction_data.get('pages', [])
        
        if not pages:
            return None
        
        # Create main visualization
        main_fig = create_extraction_visualization(pages)
        
        if main_fig:
            # Add interactivity
            main_fig.update_traces(
                hovertemplate="<b>%{x}</b><br>" +
                            "Elements: %{y}<br>" +
                            "<extra></extra>"
            )
        
        return main_fig
        
    except Exception as e:
        logger.error(f"Error creating interactive dashboard: {str(e)}")
        return None