import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from utils.data_analyzer import DataAnalyzer
from utils.kpi_generator import KPIGenerator
from utils.visualizations import DataVisualizer

# Page configuration
st.set_page_config(
    page_title="Excel Data Insights & KPI Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern UI
st.markdown("""
<style>
    /* Import modern font */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Global font styling */
    .stApp {
        font-family: 'Inter', sans-serif;
    }
    
    /* Header styling */
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
    }
    
    .main-header h1 {
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }
    
    .main-header p {
        font-size: 1.1rem;
        opacity: 0.9;
        margin: 0;
    }
    
    /* Card styling */
    .metric-card {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1.5rem;
        border-radius: 12px;
        margin: 0.5rem 0;
        border-left: 4px solid #00d4aa;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #00d4aa 0%, #00b894 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        font-weight: 500;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,212,170,0.3);
    }
    
    /* File uploader styling */
    .uploadedFile {
        border-radius: 8px;
        border: 2px dashed #00d4aa;
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 0.5rem 1rem;
        font-weight: 500;
    }
    
    /* Sidebar styling */
    .css-1d391kg {
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Success/Error message styling */
    .stSuccess {
        border-radius: 8px;
        border-left: 4px solid #00d4aa;
    }
    
    .stError {
        border-radius: 8px;
        border-left: 4px solid #e74c3c;
    }
    
    /* KPI expandable sections */
    .streamlit-expanderHeader {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        border-radius: 8px;
        border: 1px solid #dee2e6;
    }
    
    /* Welcome section styling */
    .welcome-section {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 3rem;
        border-radius: 15px;
        margin: 2rem 0;
        text-align: center;
    }
    
    .feature-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
        gap: 1.5rem;
        margin: 2rem 0;
    }
    
    .feature-card {
        background: rgba(255,255,255,0.1);
        padding: 1.5rem;
        border-radius: 12px;
        backdrop-filter: blur(10px);
        border: 1px solid rgba(255,255,255,0.2);
    }
    
    .feature-card h3 {
        font-size: 1.3rem;
        margin-bottom: 0.5rem;
        font-weight: 600;
    }
    
    .feature-card p {
        opacity: 0.9;
        line-height: 1.5;
    }
    
    /* Chart container styling */
    .stPlotlyChart {
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        overflow: hidden;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'data' not in st.session_state:
    st.session_state.data = None
if 'analysis_complete' not in st.session_state:
    st.session_state.analysis_complete = False

def main():
    # Modern header
    st.markdown("""
    <div class="main-header">
        <h1>📊 Excel Data Insights & KPI Analyzer</h1>
        <p>Transform your Excel data into actionable business intelligence with AI-powered analysis</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar for file upload and navigation
    with st.sidebar:
        st.markdown("### 📁 Data Upload")
        st.markdown("Upload your Excel file to begin analysis")
        
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Supported formats: .xlsx, .xls"
        )
        
        if uploaded_file is not None:
            try:
                # Read Excel file and handle multiple sheets
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names
                
                if len(sheet_names) > 1:
                    st.markdown("### 📄 Sheet Selection")
                    selected_sheet = st.selectbox(
                        "Choose sheet to analyze",
                        sheet_names,
                        help="Select which worksheet to analyze"
                    )
                else:
                    selected_sheet = sheet_names[0]
                
                # Load data
                data = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
                st.session_state.data = data
                st.session_state.analysis_complete = True
                
                st.success(f"✅ Successfully loaded sheet: {selected_sheet}")
                st.info(f"📊 Data shape: {data.shape[0]} rows × {data.shape[1]} columns")
                
            except Exception as e:
                st.error(f"❌ Error reading Excel file: {str(e)}")
                st.session_state.data = None
                st.session_state.analysis_complete = False
    
    # Main content area
    if st.session_state.data is not None and st.session_state.analysis_complete:
        data = st.session_state.data
        
        # Initialize analyzers
        analyzer = DataAnalyzer(data)
        kpi_generator = KPIGenerator(data)
        visualizer = DataVisualizer(data)
        
        # Create modern tabs for different analysis sections
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📋 Overview", 
            "📈 Insights", 
            "🎯 KPIs", 
            "📊 Charts", 
            "🔍 Correlations",
            "📄 Report"
        ])
        
        with tab1:
            st.markdown("### 📊 Dataset Overview")
            
            # Metrics in a modern grid
            col1, col2, col3, col4 = st.columns(4)
            basic_info = analyzer.get_basic_info()
            
            with col1:
                st.metric("📋 Total Rows", f"{basic_info['Total Rows']:,}")
            with col2:
                st.metric("📝 Columns", basic_info['Total Columns'])
            with col3:
                st.metric("💾 Size (MB)", basic_info['Memory Usage (MB)'])
            with col4:
                st.metric("🔢 Numerical", basic_info['Numerical Columns'])
            
            st.markdown("---")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("### 📋 Data Types Summary")
                col_info = analyzer.get_column_info()
                dtype_summary = col_info['Data Type'].value_counts()
                for dtype, count in dtype_summary.items():
                    st.markdown(f"**{dtype}:** {count} columns")
            
            with col2:
                st.markdown("### 📊 Column Details")
                column_info = analyzer.get_column_info()
                st.dataframe(column_info, use_container_width=True, height=300)
            
            st.markdown("### 👀 Data Preview")
            col1, col2 = st.columns([3, 1])
            with col2:
                show_rows = st.slider("Rows to display", 5, min(50, len(data)), 10)
            st.dataframe(data.head(show_rows), use_container_width=True, height=400)
            
            # Data types distribution
            st.markdown("### 🔧 Data Types Distribution")
            dtype_counts = data.dtypes.value_counts()
            # Convert to simple strings to avoid serialization issues
            fig = px.pie(
                values=dtype_counts.values, 
                names=[str(dtype) for dtype in dtype_counts.index], 
                title="Distribution of Data Types",
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            st.markdown("### 📈 Advanced Data Insights")
            
            # Quality metrics overview
            quality_metrics = analyzer.get_data_quality_metrics()
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("🎯 Completeness", f"{quality_metrics['Data Completeness']:.1f}%")
            with col2:
                st.metric("📊 Missing Cells", f"{quality_metrics['Missing Cells']:,}")
            with col3:
                st.metric("🔄 Duplicates", f"{quality_metrics['Duplicate Rows']:,}")
            with col4:
                st.metric("📦 Total Cells", f"{quality_metrics['Total Cells']:,}")
            
            st.markdown("---")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("### 📊 Statistical Summary")
                summary_stats = analyzer.get_summary_statistics()
                if not summary_stats.empty:
                    st.dataframe(summary_stats, use_container_width=True, height=350)
                else:
                    st.info("📝 No numerical columns for statistical analysis")
            
            with col2:
                st.markdown("### 🔍 Missing Values Analysis")
                missing_data = analyzer.analyze_missing_values()
                if not missing_data.empty:
                    fig = px.bar(
                        x=missing_data.index, 
                        y=missing_data['Missing_Count'],
                        title="Missing Values by Column",
                        color_discrete_sequence=['#00d4aa']
                    )
                    fig.update_layout(height=350)
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.success("✅ Perfect data quality - no missing values detected!")
            
            # Outliers detection
            st.markdown("### 🚨 Outlier Detection Results")
            outliers_info = analyzer.detect_outliers()
            if outliers_info:
                cols = st.columns(min(3, len(outliers_info)))
                for idx, (col, outliers) in enumerate(outliers_info.items()):
                    with cols[idx % 3]:
                        if len(outliers) > 0:
                            st.metric(f"📊 {col}", f"{len(outliers)} outliers")
                            with st.expander(f"View {col} outliers"):
                                st.write(outliers.tolist())
                        else:
                            st.metric(f"✅ {col}", "No outliers")
            else:
                st.info("📊 No numerical columns available for outlier analysis")
        
        with tab3:
            st.markdown("### 🎯 AI-Generated KPI Recommendations")
            st.markdown("Intelligent suggestions based on your data structure and business patterns")
            
            kpi_suggestions = kpi_generator.generate_kpi_suggestions()
            
            if kpi_suggestions:
                for category, kpis in kpi_suggestions.items():
                    st.markdown(f"#### {category}")
                    
                    for kpi in kpis:
                        with st.expander(f"🎯 {kpi['name']}", expanded=False):
                            col1, col2 = st.columns([2, 1])
                            
                            with col1:
                                st.markdown(f"**📝 Description:**  \n{kpi['description']}")
                                st.markdown(f"**🧮 Formula:**  \n`{kpi['formula']}`")
                                st.markdown(f"**💡 Business Value:**  \n{kpi['business_value']}")
                            
                            with col2:
                                # Calculate KPI if possible
                                if kpi.get('calculation'):
                                    try:
                                        result = kpi['calculation'](data)
                                        if isinstance(result, (int, float)):
                                            st.metric("💹 Current Value", f"{result:.2f}")
                                        else:
                                            st.metric("💹 Current Value", str(result))
                                    except Exception as e:
                                        st.warning(f"⚠️ Calculation unavailable")
                    
                    st.markdown("---")
            else:
                st.info("📤 Upload data to see personalized KPI recommendations")
        
        with tab4:
            st.markdown("### 📊 Interactive Data Visualizations")
            
            # Create visualization tabs
            viz_tab1, viz_tab2, viz_tab3 = st.tabs(["📈 Numerical", "📋 Categorical", "⏰ Time Series"])
            
            with viz_tab1:
                numerical_cols = data.select_dtypes(include=[np.number]).columns.tolist()
                if numerical_cols:
                    col1, col2 = st.columns([2, 1])
                    with col2:
                        selected_num_col = st.selectbox("📊 Choose column", numerical_cols, key="num_viz")
                    
                    if selected_num_col:
                        fig = visualizer.create_distribution_plot(selected_num_col)
                        st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("📊 No numerical columns available for visualization")
            
            with viz_tab2:
                categorical_cols = data.select_dtypes(include=['object']).columns.tolist()
                if categorical_cols:
                    col1, col2 = st.columns([2, 1])
                    with col2:
                        selected_cat_col = st.selectbox("📋 Choose column", categorical_cols, key="cat_viz")
                    
                    if selected_cat_col:
                        fig = visualizer.create_categorical_plot(selected_cat_col)
                        st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("📋 No categorical columns available for visualization")
            
            with viz_tab3:
                datetime_cols = data.select_dtypes(include=['datetime64']).columns.tolist()
                numerical_cols = data.select_dtypes(include=[np.number]).columns.tolist()
                
                if datetime_cols and numerical_cols:
                    col1, col2 = st.columns(2)
                    with col1:
                        selected_date_col = st.selectbox("⏰ Date column", datetime_cols, key="date_viz")
                    with col2:
                        selected_value_col = st.selectbox("📊 Value column", numerical_cols, key="value_viz")
                    
                    if selected_date_col and selected_value_col:
                        fig = visualizer.create_time_series_plot(selected_date_col, selected_value_col)
                        st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("⏰ Time series visualization requires both datetime and numerical columns")
        
        with tab5:
            st.markdown("### 🔍 Correlation Analysis")
            
            numerical_cols = data.select_dtypes(include=[np.number]).columns.tolist()
            if len(numerical_cols) > 1:
                correlation_matrix = analyzer.calculate_correlation_matrix()
                
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    st.markdown("#### 🌡️ Correlation Heatmap")
                    fig = visualizer.create_correlation_heatmap(correlation_matrix)
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    st.markdown("#### 🎯 Strong Correlations")
                    strong_corr = analyzer.find_strong_correlations()
                    if not strong_corr.empty:
                        for idx, row in strong_corr.iterrows():
                            strength_color = "🟢" if row['Correlation'] > 0 else "🔴"
                            st.markdown(f"""
                            **{strength_color} {row['Variable 1']} ↔ {row['Variable 2']}**  
                            Correlation: `{row['Correlation']:.3f}`  
                            {row['Strength']}
                            """)
                            st.markdown("---")
                    else:
                        st.info("🎯 No strong correlations detected (threshold: 0.7)")
            else:
                st.info("🔍 Correlation analysis requires at least 2 numerical columns")
        
        with tab6:
            st.markdown("### 📄 Analysis Report Generation")
            st.markdown("Generate a comprehensive report of your data analysis")
            
            col1, col2 = st.columns([2, 1])
            
            with col2:
                generate_btn = st.button("🚀 Generate Report", type="primary")
            
            if generate_btn:
                with st.spinner("🔄 Generating comprehensive analysis report..."):
                    report = generate_analysis_report(data, analyzer, kpi_generator)
                
                # Success message and download
                st.success("✅ Report generated successfully!")
                
                col1, col2 = st.columns([1, 1])
                with col1:
                    st.download_button(
                        label="📥 Download Report",
                        data=report,
                        file_name=f"data_analysis_report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.txt",
                        mime="text/plain",
                        type="primary"
                    )
                
                # Report preview with better styling
                st.markdown("#### 👀 Report Preview")
                st.text_area("", report, height=500, disabled=True)
    
    else:
        # Modern welcome screen
        st.markdown("""
        <div class="welcome-section">
            <h2>🚀 Transform Your Excel Data into Business Intelligence</h2>
            <p style="font-size: 1.2rem; margin: 1rem 0;">
                Upload your Excel file and get instant AI-powered insights, automated KPI recommendations, 
                and interactive visualizations in seconds.
            </p>
            
            <div class="feature-grid">
                <div class="feature-card">
                    <h3>📊 Smart Analysis</h3>
                    <p>Comprehensive data overview with automatic data type detection and quality assessment</p>
                </div>
                
                <div class="feature-card">
                    <h3>🎯 AI-Powered KPIs</h3>
                    <p>Intelligent KPI recommendations tailored to your specific data structure and business context</p>
                </div>
                
                <div class="feature-card">
                    <h3>📈 Interactive Charts</h3>
                    <p>Beautiful, interactive visualizations that help you understand patterns and trends instantly</p>
                </div>
                
                <div class="feature-card">
                    <h3>🔍 Deep Insights</h3>
                    <p>Advanced correlation analysis, outlier detection, and missing value assessment</p>
                </div>
                
                <div class="feature-card">
                    <h3>📄 Export Reports</h3>
                    <p>Generate comprehensive analysis reports that you can download and share with your team</p>
                </div>
                
                <div class="feature-card">
                    <h3>⚡ Real-time Processing</h3>
                    <p>Lightning-fast analysis with support for large datasets and multi-sheet workbooks</p>
                </div>
            </div>
            
            <div style="margin-top: 2rem; padding: 1.5rem; background: rgba(255,255,255,0.1); border-radius: 10px;">
                <h3>🎯 Getting Started</h3>
                <ol style="text-align: left; display: inline-block; margin: 0;">
                    <li>Click "Browse files" in the sidebar</li>
                    <li>Upload your Excel file (.xlsx or .xls)</li>
                    <li>Select a sheet if multiple are available</li>
                    <li>Explore insights across different analysis tabs</li>
                </ol>
            </div>
        </div>
        """, unsafe_allow_html=True)

def generate_analysis_report(data, analyzer, kpi_generator):
    """Generate a comprehensive text report of the analysis."""
    report = []
    report.append("=" * 60)
    report.append("EXCEL DATA ANALYSIS REPORT")
    report.append("=" * 60)
    report.append(f"Generated on: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report.append("")
    
    # Basic Info
    report.append("1. BASIC INFORMATION")
    report.append("-" * 25)
    basic_info = analyzer.get_basic_info()
    for key, value in basic_info.items():
        report.append(f"{key}: {value}")
    report.append("")
    
    # Data Quality
    report.append("2. DATA QUALITY METRICS")
    report.append("-" * 30)
    quality_metrics = analyzer.get_data_quality_metrics()
    for metric, value in quality_metrics.items():
        report.append(f"{metric}: {value}")
    report.append("")
    
    # Missing Values
    report.append("3. MISSING VALUES ANALYSIS")
    report.append("-" * 35)
    missing_data = analyzer.analyze_missing_values()
    if not missing_data.empty:
        for idx, row in missing_data.iterrows():
            report.append(f"{idx}: {row['Missing_Count']} missing ({row['Missing_Percentage']:.2f}%)")
    else:
        report.append("No missing values found.")
    report.append("")
    
    # KPI Recommendations
    report.append("4. KPI RECOMMENDATIONS")
    report.append("-" * 25)
    kpi_suggestions = kpi_generator.generate_kpi_suggestions()
    for category, kpis in kpi_suggestions.items():
        report.append(f"\n{category.upper()}:")
        for kpi in kpis:
            report.append(f"  • {kpi['name']}: {kpi['description']}")
    
    return "\n".join(report)

if __name__ == "__main__":
    main()
