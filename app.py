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
from utils.ppt_analyzer import PowerPointAnalyzer

# Page configuration
st.set_page_config(
    page_title="Data Insights & Analysis Platform",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state
if 'data' not in st.session_state:
    st.session_state.data = None
if 'ppt_analyzer' not in st.session_state:
    st.session_state.ppt_analyzer = None
if 'analysis_complete' not in st.session_state:
    st.session_state.analysis_complete = False
if 'file_type' not in st.session_state:
    st.session_state.file_type = None

def main():
    st.title("📊 Data Insights & Analysis Platform")
    st.markdown("Upload Excel files for data analysis or PowerPoint presentations for content insights.")
    
    # Sidebar for file upload and navigation
    with st.sidebar:
        st.header("📁 File Upload")
        
        # File type selection
        file_type = st.radio(
            "Select file type:",
            ["Excel Files", "PowerPoint Presentations"],
            help="Choose the type of file you want to analyze"
        )
        
        if file_type == "Excel Files":
            uploaded_file = st.file_uploader(
                "Choose an Excel file",
                type=['xlsx', 'xls'],
                help="Upload Excel files (.xlsx or .xls format)"
            )
        else:
            uploaded_file = st.file_uploader(
                "Choose a PowerPoint file",
                type=['pptx', 'ppt'],
                help="Upload PowerPoint files (.pptx or .ppt format)"
            )
        
        if uploaded_file is not None:
            try:
                if file_type == "Excel Files":
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
                    st.session_state.ppt_analyzer = None
                    st.session_state.file_type = "excel"
                    st.session_state.analysis_complete = True
                    
                    st.success(f"✅ Successfully loaded sheet: {selected_sheet}")
                    st.info(f"📊 Data shape: {data.shape[0]} rows × {data.shape[1]} columns")
                    
                else:
                    # Process PowerPoint file
                    file_content = io.BytesIO(uploaded_file.read())
                    ppt_analyzer = PowerPointAnalyzer(file_content)
                    
                    st.session_state.ppt_analyzer = ppt_analyzer
                    st.session_state.data = None
                    st.session_state.file_type = "powerpoint"
                    st.session_state.analysis_complete = True
                    
                    overview = ppt_analyzer.get_presentation_overview()
                    st.success(f"✅ Successfully loaded PowerPoint presentation")
                    st.info(f"📊 Presentation: {overview['total_slides']} slides, {overview['total_text_length']} characters")
                    
            except Exception as e:
                st.error(f"❌ Error reading file: {str(e)}")
                st.session_state.data = None
                st.session_state.ppt_analyzer = None
                st.session_state.analysis_complete = False
    
    # Main content area
    if st.session_state.ppt_analyzer is not None and st.session_state.analysis_complete:
        # PowerPoint Analysis Interface
        ppt_analyzer = st.session_state.ppt_analyzer
        
        # Create tabs for PowerPoint analysis
        tab1, tab2, tab3, tab4 = st.tabs([
            "📊 Presentation Overview",
            "📋 Content Analysis", 
            "🎯 Presentation KPIs",
            "🤖 AI Insights"
        ])
        
        with tab1:
            st.header("Presentation Overview")
            
            overview = ppt_analyzer.get_presentation_overview()
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Slides", overview['total_slides'])
                st.metric("Slides with Titles", overview['slides_with_titles'])
            with col2:
                st.metric("Total Text Length", f"{overview['total_text_length']:,} chars")
                st.metric("Average Text per Slide", f"{overview['average_text_per_slide']:.1f} chars")
            with col3:
                st.metric("Slides with Content", overview['slides_with_content'])
                st.metric("Total Bullet Points", overview['total_bullet_points'])
        
        with tab2:
            st.header("Content Structure Analysis")
            
            structure = ppt_analyzer.analyze_content_structure()
            
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Text Distribution")
                dist = structure['text_distribution']
                st.write(f"**Minimum text length:** {dist['min_text_length']} characters")
                st.write(f"**Maximum text length:** {dist['max_text_length']} characters")
                st.write(f"**Average text length:** {dist['avg_text_length']:.1f} characters")
            
            with col2:
                st.subheader("Content Consistency")
                consistency = structure['content_consistency']
                st.write(f"**Title consistency:** {consistency['title_consistency']*100:.1f}%")
                st.write(f"**Content consistency score:** {consistency['content_consistency_score']*100:.1f}%")
                st.write(f"**Has consistent structure:** {'✅ Yes' if consistency['consistent_structure'] else '❌ No'}")
            
            st.subheader("Slide Categories")
            slide_types = structure['slide_types']
            for category, count in slide_types.items():
                st.write(f"**{category.replace('_', ' ').title()}:** {count}")
            
            # Display slide content
            st.subheader("Slide Content Details")
            slides_data = ppt_analyzer.extract_slide_content()
            for slide in slides_data:
                with st.expander(f"Slide {slide['slide_number']}: {slide['title'] or 'No Title'}"):
                    if slide['title']:
                        st.write(f"**Title:** {slide['title']}")
                    if slide['content']:
                        st.write("**Content:**")
                        for content in slide['content']:
                            st.write(f"- {content}")
                    if slide['bullet_points']:
                        st.write("**Bullet Points:**")
                        for bullet in slide['bullet_points']:
                            st.write(f"• {bullet}")
                    st.write(f"**Text Length:** {slide['text_length']} characters")
                    st.write(f"**Shape Count:** {slide['shape_count']}")
        
        with tab3:
            st.header("Presentation KPIs")
            
            kpis = ppt_analyzer.generate_presentation_kpis()
            
            for kpi in kpis:
                with st.expander(f"📊 {kpi['name']}"):
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        st.metric(kpi['name'], f"{kpi['value']} {kpi['unit']}")
                    with col2:
                        st.write(f"**Category:** {kpi['category']}")
                        st.write(f"**Description:** {kpi['description']}")
                        st.write(f"**Recommendation:** {kpi['recommendation']}")
        
        with tab4:
            st.header("AI-Powered Insights")
            
            if st.button("🤖 Generate AI Analysis"):
                with st.spinner("Analyzing presentation content..."):
                    insights = ppt_analyzer.generate_ai_insights()
                
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Summary")
                    st.write(insights['summary'])
                    
                    st.subheader("Key Themes")
                    if isinstance(insights['key_themes'], list):
                        for theme in insights['key_themes']:
                            st.write(f"• {theme}")
                    else:
                        st.write(insights['key_themes'])
                
                with col2:
                    st.subheader("Content Quality")
                    st.write(insights['content_quality'])
                    
                    st.subheader("Recommendations")
                    if isinstance(insights['recommendations'], list):
                        for rec in insights['recommendations']:
                            st.write(f"• {rec}")
                    else:
                        st.write(insights['recommendations'])
    
    elif st.session_state.data is not None and st.session_state.analysis_complete:
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
            
            # Create sub-tabs for different visualization types
            viz_tab1, viz_tab2 = st.tabs(["📊 Standard Charts", "🎯 Custom Metrics"])
            
            with viz_tab1:
                # Numerical columns distribution
                numerical_cols = data.select_dtypes(include=[np.number]).columns.tolist()
                # Filter columns that have valid data
                valid_numerical_cols = []
                for col in numerical_cols:
                    clean_series = pd.to_numeric(data[col], errors='coerce')
                    clean_data = clean_series.dropna()
                    if len(clean_data) > 0:
                        valid_numerical_cols.append(col)
                
                if valid_numerical_cols:
                    st.subheader("Numerical Data Distributions")
                    selected_num_col = st.selectbox("Select numerical column", valid_numerical_cols)
                    if selected_num_col:
                        fig = visualizer.create_distribution_plot(selected_num_col)
                        st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("No valid numerical data available for distribution plots.")
            
            with viz_tab2:
                st.subheader("Create Custom Metrics")
                
                # Custom metrics creation interface
                col1, col2 = st.columns([1, 1])
                
                with col1:
                    st.markdown("**Chart Configuration**")
                    
                    # Chart type selection
                    chart_type = st.selectbox(
                        "Chart Type",
                        ["Bar Chart", "Line Chart", "Scatter Plot", "Pie Chart", "Area Chart"],
                        help="Choose the type of visualization"
                    )
                    
                    # Metric name
                    metric_name = st.text_input(
                        "Metric Name",
                        placeholder="e.g., Sales Performance",
                        help="Give your custom metric a descriptive name"
                    )
                
                with col2:
                    st.markdown("**Data Selection**")
                    
                    # Get all available columns
                    all_cols = data.columns.tolist()
                    
                    if chart_type in ["Bar Chart", "Line Chart", "Area Chart"]:
                        x_column = st.selectbox("X-axis (Categories)", all_cols, help="Choose categorical data for X-axis")
                        y_column = st.selectbox("Y-axis (Values)", valid_numerical_cols if valid_numerical_cols else all_cols, help="Choose numerical data for Y-axis")
                        color_column = st.selectbox("Color Grouping (Optional)", ["None"] + all_cols, help="Optional: Group data by color")
                        
                    elif chart_type == "Scatter Plot":
                        x_column = st.selectbox("X-axis", valid_numerical_cols if valid_numerical_cols else all_cols, help="Choose numerical data for X-axis")
                        y_column = st.selectbox("Y-axis", valid_numerical_cols if valid_numerical_cols else all_cols, help="Choose numerical data for Y-axis")
                        color_column = st.selectbox("Color Grouping (Optional)", ["None"] + all_cols, help="Optional: Group data by color")
                        
                    elif chart_type == "Pie Chart":
                        x_column = st.selectbox("Categories", all_cols, help="Choose categorical data")
                        y_column = st.selectbox("Values", valid_numerical_cols if valid_numerical_cols else all_cols, help="Choose numerical data for pie sizes")
                        color_column = None
                
                # Advanced options
                with st.expander("Advanced Options"):
                    col3, col4 = st.columns(2)
                    with col