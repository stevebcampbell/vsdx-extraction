#!/usr/bin/env python3
"""
Simple script to run the Streamlit app
"""

import subprocess
import sys
import os

def main():
    """Run the Streamlit app"""
    
    # Ensure we're in the right directory
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    print("🚀 Starting VSDX Extraction Tool...")
    print("📊 The web interface will open in your browser")
    print("🌐 URL: http://localhost:8501")
    print()
    print("Instructions:")
    print("1. Upload a VSDX file")
    print("2. Optionally add Gemini API key for AI analysis")
    print("3. Click 'Extract and Analyze'")
    print()
    print("Press Ctrl+C to stop the server")
    print("-" * 50)
    
    try:
        # Run streamlit
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", "app.py",
            "--server.port", "8501",
            "--server.address", "0.0.0.0",
            "--browser.gatherUsageStats", "false"
        ])
    except KeyboardInterrupt:
        print("\n👋 Shutting down VSDX Extraction Tool...")
    except Exception as e:
        print(f"❌ Error running app: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())