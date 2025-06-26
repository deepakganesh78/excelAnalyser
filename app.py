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

# Initialize session state
if 'data' not in st.session_state:
    st.session_state.data = None
if 'analysis_complete' not in st.session_state:
    st.session_state.analysis_complete = False

def main():
    st.title("📊 Excel Data Insights & KPI Analyzer")
    st.markdown("Upload your Excel file to get automated insights, data analysis, and KPI recommendations.")
    
    # Sidebar for file upload and navigation
    with st.sidebar:
        st.header("📁 File Upload")
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload Excel files (.xlsx or .xls format)"
        )
        
        if uploaded_file is not None:
            try:
                # Read Excel file and handle multiple sheets
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names
                
                if len(sheet_names) > 1:
                    selected_sheet = st.selectbox(
                        "Select Sheet",
                        sheet_names,
                        help="Choose which sheet to analyze"
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
        
        # Create tabs for different analysis sections
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📋 Data Overview", 
            "📈 Data Insights", 
            "🎯 KPI Recommendations", 
            "📊 Visualizations", 
            "🔍 Correlation Analysis",
            "📄 Export Report"
        ])
        
        with tab1:
            st.header("Data Overview")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Basic Information")
                basic_info = analyzer.get_basic_info()
                for key, value in basic_info.items():
                    st.metric(key, value)
            
            with col2:
                st.subheader("Column Information")
                column_info = analyzer.get_column_info()
                st.dataframe(column_info, use_container_width=True)
            
            st.subheader("Data Preview")
            show_rows = st.slider("Number of rows to display", 5, min(50, len(data)), 10)
            st.dataframe(data.head(show_rows), use_container_width=True)
            
            # Data types distribution
            st.subheader("Data Types Distribution")
            dtype_counts = data.dtypes.value_counts()
            # Convert to simple strings to avoid serialization issues
            fig = px.pie(
                values=dtype_counts.values, 
                names=[str(dtype) for dtype in dtype_counts.index], 
                title="Distribution of Data Types"
            )
            st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            st.header("Data Insights")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Summary Statistics")
                summary_stats = analyzer.get_summary_statistics()
                st.dataframe(summary_stats, use_container_width=True)
            
            with col2:
                st.subheader("Data Quality Assessment")
                quality_metrics = analyzer.get_data_quality_metrics()
                
                for metric, value in quality_metrics.items():
                    if isinstance(value, (int, float)):
                        # Skip NaN and infinite values
                        if not (np.isnan(value) or np.isinf(value)):
                            st.metric(metric, f"{value:.2f}" if isinstance(value, float) else value)
                    else:
                        st.write(f"**{metric}:** {value}")
            
            # Missing values analysis
            st.subheader("Missing Values Analysis")
            missing_data = analyzer.analyze_missing_values()
            if not missing_data.empty:
                fig = px.bar(
                    x=missing_data.index, 
                    y=missing_data['Missing_Count'],
                    title="Missing Values by Column"
                )
                st.plotly_chart(fig, use_container_width=True)
                st.dataframe(missing_data, use_container_width=True)
            else:
                st.success("✅ No missing values found!")
            
            # Outliers detection
            st.subheader("Outliers Detection")
            outliers_info = analyzer.detect_outliers()
            if outliers_info:
                for col, outliers in outliers_info.items():
                    if len(outliers) > 0:
                        st.write(f"**{col}:** {len(outliers)} outliers detected")
                        with st.expander(f"View outliers for {col}"):
                            st.write(outliers.tolist())
            else:
                st.info("No numerical columns found for outlier detection.")
        
        with tab3:
            st.header("KPI Recommendations")
            
            kpi_suggestions = kpi_generator.generate_kpi_suggestions()
            
            if kpi_suggestions:
                # Filter and show only critical KPIs
                critical_categories = ['Business KPIs', 'Data Quality KPIs', 'Performance KPIs']
                
                for category in critical_categories:
                    if category in kpi_suggestions:
                        kpis = kpi_suggestions[category]
                        if kpis:  # Only show if there are KPIs in this category
                            st.subheader(f"{category}")
                            
                            # Limit to top 3 KPIs per category for focus
                            for kpi in kpis[:3]:
                                with st.expander(f"🎯 {kpi['name']}"):
                                    st.write(f"**Description:** {kpi['description']}")
                                    st.write(f"**Formula:** {kpi['formula']}")
                                    st.write(f"**Business Value:** {kpi['business_value']}")
                                    
                                    # Calculate KPI if possible
                                    if kpi.get('calculation'):
                                        try:
                                            result = kpi['calculation'](data)
                                            # Only show valid results (not NaN or infinite)
                                            if isinstance(result, (int, float)) and not (np.isnan(result) or np.isinf(result)):
                                                st.metric("Calculated Value", f"{result:.2f}")
                                            elif not isinstance(result, (int, float)):
                                                st.metric("Calculated Value", str(result))
                                        except Exception as e:
                                            st.warning(f"Could not calculate: {str(e)}")
            else:
                st.info("Upload data to see KPI recommendations.")
        
        with tab4:
            st.header("Data Visualizations")
            
            # Numerical columns distribution
            numerical_cols = data.select_dtypes(include=[np.number]).columns.tolist()
            if numerical_cols:
                st.subheader("Numerical Data Distributions")
                selected_num_col = st.selectbox("Select numerical column", numerical_cols)
                if selected_num_col:
                    fig = visualizer.create_distribution_plot(selected_num_col)
                    st.plotly_chart(fig, use_container_width=True)
            
            # Categorical columns
            categorical_cols = data.select_dtypes(include=['object']).columns.tolist()
            if categorical_cols:
                st.subheader("Categorical Data Analysis")
                selected_cat_col = st.selectbox("Select categorical column", categorical_cols)
                if selected_cat_col:
                    fig = visualizer.create_categorical_plot(selected_cat_col)
                    st.plotly_chart(fig, use_container_width=True)
            
            # Time series analysis if datetime columns exist
            datetime_cols = data.select_dtypes(include=['datetime64']).columns.tolist()
            if datetime_cols:
                st.subheader("Time Series Analysis")
                selected_date_col = st.selectbox("Select datetime column", datetime_cols)
                if selected_date_col and numerical_cols:
                    selected_value_col = st.selectbox("Select value column", numerical_cols)
                    fig = visualizer.create_time_series_plot(selected_date_col, selected_value_col)
                    st.plotly_chart(fig, use_container_width=True)
        
        with tab5:
            st.header("Correlation Analysis")
            
            if len(numerical_cols) > 1:
                correlation_matrix = analyzer.calculate_correlation_matrix()
                
                # Correlation heatmap
                fig = visualizer.create_correlation_heatmap(correlation_matrix)
                st.plotly_chart(fig, use_container_width=True)
                
                # Strong correlations
                st.subheader("Strong Correlations")
                strong_corr = analyzer.find_strong_correlations()
                if not strong_corr.empty:
                    st.dataframe(strong_corr, use_container_width=True)
                else:
                    st.info("No strong correlations found (threshold: 0.7)")
            else:
                st.info("Need at least 2 numerical columns for correlation analysis.")
        
        with tab6:
            st.header("Export Analysis Report")
            
            if st.button("📄 Generate Report"):
                report = generate_analysis_report(data, analyzer, kpi_generator)
                
                # Create download button
                st.download_button(
                    label="📥 Download Report",
                    data=report,
                    file_name=f"data_analysis_report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.txt",
                    mime="text/plain"
                )
                
                # Display report preview
                st.subheader("Report Preview")
                st.text_area("Report Content", report, height=400)
    
    else:
        # Welcome screen
        st.markdown("""
        ## Welcome to Excel Data Insights & KPI Analyzer! 🎉
        
        This application helps you analyze your Excel data and provides:
        
        - 📊 **Comprehensive Data Overview**: Basic statistics, data types, and structure analysis
        - 🔍 **Deep Data Insights**: Missing values, outliers, and data quality assessment
        - 🎯 **Smart KPI Recommendations**: Automated suggestions based on your data structure
        - 📈 **Interactive Visualizations**: Charts and graphs to understand your data better
        - 🔗 **Correlation Analysis**: Discover relationships between variables
        - 📄 **Exportable Reports**: Download comprehensive analysis reports
        
        ### How to Get Started:
        1. Click "Browse files" in the sidebar
        2. Upload your Excel file (.xlsx or .xls)
        3. Select a sheet if your file has multiple sheets
        4. Explore the analysis tabs above!
        
        ### Supported Features:
        - ✅ Multiple Excel formats (.xlsx, .xls)
        - ✅ Multi-sheet workbooks
        - ✅ Large dataset handling
        - ✅ Real-time analysis
        - ✅ Interactive visualizations
        """)

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
