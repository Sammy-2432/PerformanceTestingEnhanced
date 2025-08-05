"""
Truist Smart Compliance Checker
Main application with improved performance and architecture
"""

import streamlit as st
import sys
from pathlib import Path

# Add src to Python path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.components.app_core import OptimizedComplianceApp
from src.config import APP, UI
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)


def main():
    """Main application entry point"""
    try:
        # Set page config with optimized settings
        st.set_page_config(
            page_title=APP.title,
            page_icon=APP.icon,
            layout="wide",
            initial_sidebar_state="expanded",
            menu_items={
                'About': f"{APP.title} v{APP.version} - Optimized for performance and efficiency"
            }
        )
        
        # Initialize and run the optimized app
        app = OptimizedComplianceApp()
        app.run()
        
    except Exception as e:
        logger.error(f"Application error: {e}")
        st.error(f"❌ Application Error: {str(e)}")
        st.error("Please check the logs for more details.")


if __name__ == "__main__":
    main()
