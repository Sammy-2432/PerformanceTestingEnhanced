"""
Optimized Application Core
Main application logic with performance optimizations and clean architecture
"""

import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, Any, Optional
import logging
import gc
from functools import lru_cache

# Import optimized components
from ..utils.excel_reader import OptimizedExcelReader
from ..analyzers.docx_analyzer import OptimizedDocxAnalyzer
from ..analyzers.ppt_analyzer import OptimizedPowerPointAnalyzer
from ..compliance.docx_compliance import OptimizedDocxComplianceChecker
from ..compliance.pptx_compliance import OptimizedPptxComplianceChecker
from ..config import (
    APP, UI, EXCEL_FILE_PATH, UPDATE_DAY, 
    DATA_DIR, TEMP_DIR, ALTERNATIVE_EXCEL_PATHS
)

logger = logging.getLogger(__name__)


class OptimizedComplianceApp:
    """
    Optimized main application class with:
    - Efficient session state management
    - Memory optimization
    - Performance monitoring
    - Clean separation of concerns
    """
    
    def __init__(self):
        """Initialize the optimized application"""
        self._setup_custom_css()
        self._initialize_session_state()
        self._setup_performance_monitoring()
    
    def _setup_custom_css(self):
        """Setup optimized CSS styles"""
        st.markdown("""
        <style>
            .main-header {
                font-size: 2.5rem;
                font-weight: bold;
                text-align: center;
                background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                margin-bottom: 1.5rem;
            }
            
            .compliance-card {
                padding: 1rem;
                border-radius: 8px;
                margin: 0.8rem 0;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
                transition: transform 0.2s ease;
            }
            
            .compliance-card:hover {
                transform: translateY(-2px);
            }
            
            .compliant { background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white; }
            .non-compliant { background: linear-gradient(135deg, #dc3545 0%, #fd7e14 100%); color: white; }
            .partial-compliant { background: linear-gradient(135deg, #ffc107 0%, #fd7e14 100%); color: black; }
            
            .metric-container {
                background: #f8f9fa;
                padding: 1rem;
                border-radius: 8px;
                border-left: 4px solid #007bff;
            }
            
            .performance-indicator {
                position: fixed;
                top: 10px;
                right: 10px;
                background: rgba(0, 0, 0, 0.8);
                color: white;
                padding: 5px 10px;
                border-radius: 5px;
                font-size: 0.8rem;
                z-index: 1000;
            }
            
            /* Optimize for mobile */
            @media (max-width: 768px) {
                .main-header { font-size: 2rem; }
                .compliance-card { margin: 0.5rem 0; padding: 0.8rem; }
            }
        </style>
        """, unsafe_allow_html=True)
    
    def _initialize_session_state(self):
        """Initialize session state with optimized defaults"""
        defaults = {
            'selected_release': None,
            'selected_project': None,
            'selected_enterprise_release_id': None,
            'excel_reader': None,
            'performance_stats': {},
            'last_analysis_time': None,
            'memory_usage': 0
        }
        
        for key, value in defaults.items():
            if key not in st.session_state:
                st.session_state[key] = value
    
    def _setup_performance_monitoring(self):
        """Setup performance monitoring"""
        if 'performance_stats' not in st.session_state:
            st.session_state.performance_stats = {
                'total_analyses': 0,
                'avg_processing_time': 0.0,
                'memory_usage_mb': 0.0,
                'cache_hits': 0
            }
    
    @lru_cache(maxsize=1)
    def _get_last_wednesday(self) -> datetime:
        """Get the date of the last Wednesday (cached)"""
        today = datetime.now()
        days_after_wednesday = (today.weekday() - UPDATE_DAY) % 7
        return today - timedelta(days=days_after_wednesday)
    
    def _is_excel_updated(self, file_path: Path) -> bool:
        """Check if Excel file is updated efficiently"""
        if not file_path.exists():
            return False
        
        file_modified = datetime.fromtimestamp(file_path.stat().st_mtime)
        last_wednesday = self._get_last_wednesday()
        
        return file_modified >= last_wednesday
    
    @st.cache_resource(ttl=3600)  # Cache for 1 hour
    def _load_excel_data(_self, file_path: str) -> Optional[OptimizedExcelReader]:
        """Load Excel data with caching"""
        try:
            excel_reader = OptimizedExcelReader(file_path)
            if excel_reader.load_data(optimize_memory=True):
                return excel_reader
            return None
        except Exception as e:
            logger.error(f"Error loading Excel data: {e}")
            return None
    
    def _handle_excel_file(self) -> Optional[str]:
        """Handle Excel file loading with fallback options"""
        # Check main Excel file
        if EXCEL_FILE_PATH.exists() and self._is_excel_updated(EXCEL_FILE_PATH):
            st.success("✅ Using up-to-date Excel file")
            return str(EXCEL_FILE_PATH)
        
        # Check alternative paths
        for alt_path in ALTERNATIVE_EXCEL_PATHS:
            alt_path = Path(alt_path)
            if alt_path.exists() and self._is_excel_updated(alt_path):
                st.info(f"📁 Using Excel file from: {alt_path}")
                return str(alt_path)
        
        # Show upload option
        st.warning("⚠️ Excel file is outdated or not found. Please upload a current file.")
        
        uploaded_excel = st.file_uploader(
            "Upload Updated Excel File",
            type=['xlsx', 'xls'],
            help="Upload the latest project data Excel file"
        )
        
        if uploaded_excel:
            # Save uploaded file temporarily
            temp_path = TEMP_DIR / f"temp_excel_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            with open(temp_path, 'wb') as f:
                f.write(uploaded_excel.getvalue())
            
            st.success("✅ Excel file uploaded successfully!")
            return str(temp_path)
        
        return None
    
    def _render_project_selection_ui(self, excel_reader: OptimizedExcelReader):
        """Render optimized project selection UI"""
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Release dropdown with caching
            releases = excel_reader.get_releases()
            if releases:
                release_index = 0
                if st.session_state.selected_release and st.session_state.selected_release in releases:
                    release_index = releases.index(st.session_state.selected_release) + 1
                
                selected_release = st.selectbox(
                    "🔄 Select Release",
                    options=[""] + releases,
                    index=release_index,
                    help="Select the release version",
                    key="release_selector"
                )
                
                if selected_release != st.session_state.selected_release:
                    st.session_state.selected_release = selected_release if selected_release else None
                    st.session_state.selected_project = None
                    st.session_state.selected_enterprise_release_id = None
                    st.rerun()
            else:
                st.warning("⚠️ No releases found in Excel data")
                return
        
        with col2:
            # Project dropdown
            if st.session_state.selected_release:
                projects = excel_reader.get_projects_by_release(st.session_state.selected_release)
                if projects:
                    project_index = 0
                    if st.session_state.selected_project and st.session_state.selected_project in projects:
                        project_index = projects.index(st.session_state.selected_project) + 1
                    
                    selected_project = st.selectbox(
                        "📁 Select Project",
                        options=[""] + projects,
                        index=project_index,
                        help="Select the project name",
                        key="project_selector"
                    )
                    
                    if selected_project != st.session_state.selected_project:
                        st.session_state.selected_project = selected_project if selected_project else None
                        st.session_state.selected_enterprise_release_id = None
                        st.rerun()
                else:
                    st.selectbox("📁 Select Project", options=["No projects available"], disabled=True)
            else:
                st.selectbox("📁 Select Project", options=["Select Release first"], disabled=True)
        
        with col3:
            # Enterprise Release ID dropdown
            if st.session_state.selected_release and st.session_state.selected_project:
                enterprise_release_ids = excel_reader.get_enterprise_release_ids_by_release_and_project(
                    st.session_state.selected_release, 
                    st.session_state.selected_project
                )
                if enterprise_release_ids:
                    eid_index = 0
                    if (st.session_state.selected_enterprise_release_id and 
                        st.session_state.selected_enterprise_release_id in enterprise_release_ids):
                        eid_index = enterprise_release_ids.index(st.session_state.selected_enterprise_release_id) + 1
                    
                    selected_eid = st.selectbox(
                        "🆔 Select Enterprise Release ID",
                        options=[""] + enterprise_release_ids,
                        index=eid_index,
                        help="Select the Enterprise Release ID",
                        key="eid_selector"
                    )
                    
                    if selected_eid != st.session_state.selected_enterprise_release_id:
                        st.session_state.selected_enterprise_release_id = selected_eid if selected_eid else None
                else:
                    st.selectbox("🆔 Select Enterprise Release ID", options=["No Enterprise Release IDs available"], disabled=True)
            else:
                st.selectbox("🆔 Select Enterprise Release ID", options=["Select Project first"], disabled=True)
    
    def _display_project_info(self, excel_reader: OptimizedExcelReader):
        """Display selected project information efficiently"""
        if all([st.session_state.selected_release, 
                st.session_state.selected_project, 
                st.session_state.selected_enterprise_release_id]):
            
            project_data = excel_reader.get_project_data_by_criteria(
                st.session_state.selected_release,
                st.session_state.selected_project,
                st.session_state.selected_enterprise_release_id
            )
            
            if project_data:
                st.markdown("### 📊 Selected Project Information")
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.info(f"**Release:** {project_data.get('Release', 'N/A')}")
                with col2:
                    st.info(f"**Project:** {project_data.get('Project Name', 'N/A')}")
                with col3:
                    st.info(f"**Release ID:** {project_data.get('Enterprise Release ID', 'N/A')}")
                with col4:
                    st.info(f"**End Date:** {project_data.get('End Date', 'N/A')}")
    
    def _show_performance_stats(self):
        """Show performance statistics"""
        stats = st.session_state.performance_stats
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Analyses", stats.get('total_analyses', 0))
        with col2:
            st.metric("Avg Time (s)", f"{stats.get('avg_processing_time', 0):.2f}")
        with col3:
            st.metric("Memory (MB)", f"{stats.get('memory_usage_mb', 0):.1f}")
        with col4:
            st.metric("Cache Hits", stats.get('cache_hits', 0))
    
    def run(self):
        """Main application run method"""
        # Header
        st.markdown(f'<h1 class="main-header">{APP.icon} {APP.title}</h1>', unsafe_allow_html=True)
        
        # Sidebar navigation
        with st.sidebar:
            st.markdown("## 🧭 Navigation")
            
            page = st.selectbox(
                "Choose a section:",
                ["🏠 Home", "📄 Test Plan Review", "📊 Test Report Review", "⚡ Performance"]
            )
            
            # Performance indicator
            if st.session_state.performance_stats.get('total_analyses', 0) > 0:
                st.markdown("### 📈 Quick Stats")
                self._show_performance_stats()
        
        # Route to appropriate page
        if page == "🏠 Home":
            self._show_home_page()
        elif page == "📄 Test Plan Review":
            self._show_test_plan_page()
        elif page == "📊 Test Report Review":
            self._show_test_report_page()
        elif page == "⚡ Performance":
            self._show_performance_page()
    
    def _show_home_page(self):
        """Show optimized home page"""
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.markdown("""
            ### Welcome to the Truist Smart Compliance Checker! 🎯
            
            This optimized application helps you validate test plan and test report documents 
            against project data with improved performance and efficiency.
            
            #### 🚀 Features:
            
            - **High Performance**: Optimized for speed and memory efficiency
            - **Intelligent Caching**: Reduces processing time for repeated operations
            - **Parallel Processing**: Faster compliance checking
            - **Enterprise Release ID Support**: Updated to use Enterprise Release IDs
            - **Real-time Performance Monitoring**: Track application performance
            
            #### 📋 How to Use:
            
            1. **Navigate**: Use the sidebar to select Test Plan or Test Report Review
            2. **Select Project**: Choose Release → Project → Enterprise Release ID
            3. **Upload Document**: Upload your DOCX (test plan) or PPTX (test report)
            4. **Run Analysis**: Click the compliance check button
            5. **Review Results**: View detailed compliance results and recommendations
            
            #### 🔧 Optimization Features:
            
            - **Memory Management**: Efficient handling of large files
            - **Vectorized Operations**: Fast data processing using numpy
            - **Smart Caching**: Reduced redundant calculations
            - **Parallel Execution**: Multiple compliance checks run simultaneously
            """)
        
        with col2:
            st.markdown("### 📊 System Status")
            
            # Excel file status
            excel_path = self._handle_excel_file()
            if excel_path:
                st.success("✅ Excel Data Available")
            else:
                st.error("❌ Excel Data Required")
            
            # Memory usage
            import psutil
            import os
            process = psutil.Process(os.getpid())
            memory_mb = process.memory_info().rss / 1024 / 1024
            st.metric("Memory Usage", f"{memory_mb:.1f} MB")
            
            # Version info
            st.info(f"**Version:** {APP.version}")
    
    def _show_test_plan_page(self):
        """Show optimized test plan review page"""
        st.markdown("## 📄 Test Plan Review")
        
        # Handle Excel file
        excel_file_path = self._handle_excel_file()
        if not excel_file_path:
            return
        
        # Load Excel data
        excel_reader = self._load_excel_data(excel_file_path)
        if not excel_reader:
            st.error("❌ Failed to load Excel data")
            return
        
        # Project selection
        st.markdown("### 📋 Project Selection")
        self._render_project_selection_ui(excel_reader)
        
        # Display project info
        self._display_project_info(excel_reader)
        
        st.markdown("---")
        
        # File upload
        st.markdown("### 📄 Upload Test Plan Document")
        uploaded_file = st.file_uploader(
            "Choose a DOCX file",
            type=['docx'],
            help="Upload your test plan document in DOCX format"
        )
        
        # Compliance check
        compliance_enabled = all([
            st.session_state.selected_release,
            st.session_state.selected_project,
            st.session_state.selected_enterprise_release_id,
            uploaded_file
        ])
        
        if uploaded_file:
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button(
                    "🔍 Check Compliance",
                    type="primary",
                    use_container_width=True,
                    disabled=not compliance_enabled
                ):
                    if compliance_enabled:
                        self._run_docx_compliance_check(uploaded_file, excel_reader)
                    else:
                        st.warning("⚠️ Please complete all selections before checking compliance.")
        else:
            st.info("📤 Please upload a DOCX file to begin compliance checking")
    
    def _show_test_report_page(self):
        """Show optimized test report review page"""
        st.markdown("## 📊 Test Report Review")
        
        # Handle Excel file
        excel_file_path = self._handle_excel_file()
        if not excel_file_path:
            return
        
        # Load Excel data  
        excel_reader = self._load_excel_data(excel_file_path)
        if not excel_reader:
            st.error("❌ Failed to load Excel data")
            return
        
        # Project selection (reuse same UI)
        st.markdown("### 📋 Project Selection")
        self._render_project_selection_ui(excel_reader)
        
        # Display project info
        self._display_project_info(excel_reader)
        
        st.markdown("---")
        
        # File upload
        st.markdown("### 📄 Upload Test Report Document")
        uploaded_file = st.file_uploader(
            "Choose a PPTX file",
            type=['pptx'],
            help="Upload your test report document in PPTX format"
        )
        
        # Compliance check
        compliance_enabled = all([
            st.session_state.selected_release,
            st.session_state.selected_project,
            st.session_state.selected_enterprise_release_id,
            uploaded_file
        ])
        
        if uploaded_file:
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button(
                    "🔍 Check Test Report Compliance",
                    type="primary",
                    use_container_width=True,
                    disabled=not compliance_enabled
                ):
                    if compliance_enabled:
                        self._run_pptx_compliance_check(uploaded_file, excel_reader)
                    else:
                        st.warning("⚠️ Please complete all selections before checking compliance.")
        else:
            st.info("📤 Please upload a PPTX file to begin compliance checking")
    
    def _show_performance_page(self):
        """Show performance monitoring page"""
        st.markdown("## ⚡ Performance Monitoring")
        
        # Performance statistics
        st.markdown("### 📊 Application Statistics")
        self._show_performance_stats()
        
        # Memory management
        st.markdown("### 🧹 Memory Management")
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("🗑️ Clear Cache"):
                # Clear various caches
                st.cache_data.clear()
                if hasattr(st.session_state, 'excel_reader') and st.session_state.excel_reader:
                    st.session_state.excel_reader.clear_cache()
                gc.collect()
                st.success("✅ Cache cleared successfully")
        
        with col2:
            if st.button("📊 Collect Garbage"):
                gc.collect()
                st.success("✅ Garbage collection completed")
        
        # System information
        st.markdown("### 💻 System Information")
        try:
            import psutil
            import os
            
            process = psutil.Process(os.getpid())
            memory_info = process.memory_info()
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("RSS Memory", f"{memory_info.rss / 1024 / 1024:.1f} MB")
            with col2:
                st.metric("VMS Memory", f"{memory_info.vms / 1024 / 1024:.1f} MB")
            with col3:
                st.metric("CPU Usage", f"{process.cpu_percent():.1f}%")
                
        except ImportError:
            st.warning("⚠️ psutil not installed. Install with: pip install psutil")
    
    def _run_docx_compliance_check(self, uploaded_file, excel_reader: OptimizedExcelReader):
        """Run optimized DOCX compliance check"""
        start_time = datetime.now()
        
        try:
            # Progress tracking
            progress = st.progress(0)
            status = st.empty()
            
            # Step 1: Get project data
            status.text("📊 Retrieving project data...")
            progress.progress(20)
            
            project_data = excel_reader.get_project_data_by_criteria(
                st.session_state.selected_release,
                st.session_state.selected_project,
                st.session_state.selected_enterprise_release_id
            )
            
            # Step 2: Analyze document
            status.text("📄 Analyzing DOCX document...")
            progress.progress(40)
            
            docx_analyzer = OptimizedDocxAnalyzer(uploaded_file, cache_enabled=True)
            docx_data = docx_analyzer.analyze()
            
            # Step 3: Run compliance checks
            status.text("🔍 Running compliance checks...")
            progress.progress(60)
            
            # Create filtered Excel data for compliance
            excel_data = {
                'release': [project_data.get('Release')],
                'business_app_id': [project_data.get('Business Application ID')],
                'enterprise_release_id': [project_data.get('Enterprise Release ID')],
                'project_name': [project_data.get('Project Name')],
                'task_id': [project_data.get('Task ID')],
                'end_date': [project_data.get('End Date')]
            }
            
            checker = OptimizedDocxComplianceChecker(excel_data, docx_data)
            results = checker.run_all_checks(parallel=True)
            
            # Step 4: Display results
            status.text("📋 Preparing results...")
            progress.progress(80)
            
            self._display_compliance_results(results, "Test Plan")
            
            progress.progress(100)
            status.text("✅ Analysis complete!")
            
            # Update performance stats
            processing_time = (datetime.now() - start_time).total_seconds()
            self._update_performance_stats(processing_time)
            
        except Exception as e:
            st.error(f"❌ Error during compliance check: {str(e)}")
            logger.error(f"DOCX compliance check error: {e}")
    
    def _run_pptx_compliance_check(self, uploaded_file, excel_reader: OptimizedExcelReader):
        """Run optimized PPTX compliance check"""
        start_time = datetime.now()
        
        try:
            # Progress tracking
            progress = st.progress(0)
            status = st.empty()
            
            # Step 1: Get project data
            status.text("📊 Retrieving project data...")
            progress.progress(20)
            
            project_data = excel_reader.get_project_data_by_criteria(
                st.session_state.selected_release,
                st.session_state.selected_project,
                st.session_state.selected_enterprise_release_id
            )
            
            # Step 2: Analyze document
            status.text("📊 Analyzing PPTX document...")
            progress.progress(40)
            
            ppt_analyzer = OptimizedPowerPointAnalyzer(uploaded_file, cache_enabled=True)
            ppt_data = ppt_analyzer.analyze()
            
            # Step 3: Run compliance checks
            status.text("🔍 Running compliance checks...")
            progress.progress(60)
            
            # Create project data for compliance
            excel_data = {
                'release': [project_data.get('Release')],
                'business_app_id': [project_data.get('Business Application ID')],
                'enterprise_release_id': [project_data.get('Enterprise Release ID')],
                'project_name': [project_data.get('Project Name')],
                'task_id': [project_data.get('Task ID')],
                'end_date': [project_data.get('End Date')]
            }
            
            checker = OptimizedPptxComplianceChecker(excel_data, ppt_data)
            results = checker.run_all_checks(parallel=True)
            
            # Step 4: Display results
            status.text("📋 Preparing results...")
            progress.progress(80)
            
            self._display_compliance_results(results, "Test Report")
            
            progress.progress(100)
            status.text("✅ Analysis complete!")
            
            # Update performance stats
            processing_time = (datetime.now() - start_time).total_seconds()
            self._update_performance_stats(processing_time)
            
        except Exception as e:
            st.error(f"❌ Error during compliance check: {str(e)}")
            logger.error(f"PPTX compliance check error: {e}")
    
    def _display_compliance_results(self, results: Dict[str, Any], document_type: str):
        """Display compliance results with improved formatting"""
        st.markdown(f"### 📋 {document_type} Compliance Results")
        
        # Overall statistics
        overall_stats = results.get('overall_statistics', {})
        overall_score = overall_stats.get('overall_score', 0)
        compliance_level = overall_stats.get('compliance_level', 'Unknown')
        
        # Color-coded header
        if overall_score >= 0.8:
            header_class = "compliant"
        elif overall_score >= 0.6:
            header_class = "partial-compliant" 
        else:
            header_class = "non-compliant"
        
        st.markdown(f"""
        <div class="compliance-card {header_class}">
            <h3>Overall Compliance: {compliance_level}</h3>
            <h2>Score: {overall_score:.1%}</h2>
            <p>Passed: {overall_stats.get('passed_checks', 0)} / {overall_stats.get('total_checks', 0)} checks</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Individual check results
        individual_checks = results.get('individual_checks', {})
        
        if individual_checks:
            st.markdown("### 📝 Detailed Check Results")
            
            for check_name, check_result in individual_checks.items():
                check_title = check_name.replace('_', ' ').title()
                passed = check_result.get('passed', False)
                score = check_result.get('score', 0)
                
                # Create expandable section for each check
                with st.expander(f"{'✅' if passed else '❌'} {check_title} (Score: {score:.1%})", expanded=not passed):
                    
                    # Show overall status
                    if passed:
                        st.success(f"✅ **PASSED** - This check meets compliance requirements")
                    else:
                        st.error(f"❌ **FAILED** - This check requires attention")
                    
                    # Show details as bullet points instead of JSON
                    details = check_result.get('details', {})
                    if details:
                        st.markdown("**Check Details:**")
                        self._display_check_details_as_bullets(details)
                    
                    # Show errors if any
                    errors = check_result.get('errors', [])
                    if errors:
                        st.warning("⚠️ **Issues Found:**")
                        for error in errors:
                            st.markdown(f"• {error}")
        else:
            st.info("No detailed check results available.")
    
    def _display_check_details_as_bullets(self, details: Dict[str, Any]):
        """Display check details as formatted bullet points with tick/cross icons"""
        for key, value in details.items():
            if key == 'field_checks' and isinstance(value, dict):
                # Handle field validation checks specifically
                st.markdown("**Field Validation Results:**")
                for field_name, field_details in value.items():
                    field_title = field_name.replace('_validation', '').replace('_', ' ').title()
                    self._display_field_check_result(field_title, field_details)
            elif key in ['total_fields_checked', 'fields_passed']:
                # Handle summary statistics
                icon = "📊"
                formatted_key = key.replace('_', ' ').title()
                st.markdown(f"{icon} **{formatted_key}:** {value}")
            elif isinstance(value, dict):
                # Handle other nested dictionaries
                st.markdown(f"**{key.replace('_', ' ').title()}:**")
                for sub_key, sub_value in value.items():
                    icon = self._get_status_icon(sub_value)
                    formatted_key = sub_key.replace('_', ' ').title()
                    st.markdown(f"  {icon} {formatted_key}: {self._format_value(sub_value)}")
            elif isinstance(value, list):
                # Handle lists
                st.markdown(f"**{key.replace('_', ' ').title()}:**")
                for item in value:
                    if isinstance(item, dict):
                        for item_key, item_value in item.items():
                            icon = self._get_status_icon(item_value)
                            formatted_key = item_key.replace('_', ' ').title()
                            st.markdown(f"  {icon} {formatted_key}: {self._format_value(item_value)}")
                    else:
                        icon = self._get_status_icon(item)
                        st.markdown(f"  {icon} {self._format_value(item)}")
            else:
                # Handle simple key-value pairs
                icon = self._get_status_icon(value)
                formatted_key = key.replace('_', ' ').title()
                st.markdown(f"{icon} **{formatted_key}:** {self._format_value(value)}")
    
    def _display_field_check_result(self, field_name: str, field_details: Dict[str, Any]):
        """Display individual field check result with specific formatting"""
        document_value = field_details.get('document_value', 'Not found')
        excel_values = field_details.get('excel_values', [])
        best_match = field_details.get('best_match', 'No match')
        similarity_score = field_details.get('similarity_score', 0)
        
        # Determine if this is a match
        is_match = similarity_score >= 0.8  # Using threshold from config
        icon = "✅" if is_match else "❌"
        
        st.markdown(f"  {icon} **{field_name}:**")
        st.markdown(f"    📄 **Document Value:** `{document_value}`")
        
        if excel_values:
            st.markdown(f"    📊 **Expected Values:** `{', '.join(str(v) for v in excel_values)}`")
        
        if best_match and best_match != document_value:
            match_icon = "✅" if is_match else "⚠️"
            st.markdown(f"    {match_icon} **Best Match:** `{best_match}` (Similarity: {similarity_score:.1%})")
        elif is_match:
            st.markdown(f"    ✅ **Status:** Exact match found!")
        else:
            st.markdown(f"    ❌ **Status:** No sufficient match (Similarity: {similarity_score:.1%})")
    
    def _get_status_icon(self, value: Any) -> str:
        """Get appropriate icon based on value type and content"""
        if isinstance(value, bool):
            return "✅" if value else "❌"
        elif isinstance(value, str):
            if value.lower() in ['pass', 'passed', 'success', 'found', 'valid', 'compliant']:
                return "✅"
            elif value.lower() in ['fail', 'failed', 'error', 'missing', 'invalid', 'non-compliant']:
                return "❌"
            elif value.lower() in ['warning', 'partial', 'incomplete']:
                return "⚠️"
            else:
                return "📄"
        elif isinstance(value, (int, float)):
            if value > 0.8:
                return "✅"
            elif value > 0.5:
                return "⚠️"
            else:
                return "❌"
        else:
            return "📄"
    
    def _format_value(self, value: Any) -> str:
        """Format value for display"""
        if isinstance(value, bool):
            return "Yes" if value else "No"
        elif isinstance(value, float):
            if 0 <= value <= 1:
                return f"{value:.1%}"
            else:
                return f"{value:.2f}"
        elif value is None:
            return "Not Available"
        else:
            return str(value)
    
    def _update_performance_stats(self, processing_time: float):
        """Update performance statistics"""
        stats = st.session_state.performance_stats
        
        stats['total_analyses'] = stats.get('total_analyses', 0) + 1
        
        # Update average processing time
        current_avg = stats.get('avg_processing_time', 0)
        total_analyses = stats['total_analyses']
        stats['avg_processing_time'] = (current_avg * (total_analyses - 1) + processing_time) / total_analyses
        
        # Update memory usage
        try:
            import psutil
            import os
            process = psutil.Process(os.getpid())
            stats['memory_usage_mb'] = process.memory_info().rss / 1024 / 1024
        except ImportError:
            pass
        
        st.session_state.performance_stats = stats
