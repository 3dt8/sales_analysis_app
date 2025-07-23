import streamlit as st
import pandas as pd
import yaml
from src.data_processing import load_and_validate_data
from src.visualizations import (
    plot_overview, plot_product_analysis, plot_customer_analysis, plot_tdv_analysis,
    get_summary_info
)
from src.utils import log_error
import io

# Load configuration
try:
    with open("config.yaml", "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
    if not config or 'app' not in config or 'title' not in config['app']:
        raise KeyError
except (FileNotFoundError, KeyError):
    log_error("Using default configuration due to config.yaml error")
    config = {
        "app": {"title": "Phân Tích Doanh Số"},
        "colors": {"prev_year": "#F97316", "curr_year": "#3B82F6"}  # Tailwind CSS colors
    }

# Set Streamlit page configuration with wide layout and custom theme
st.set_page_config(
    layout="wide",
    page_title=config["app"]["title"],
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# Custom CSS for UI improvements
st.markdown("""
    <style>
    .stApp {
        font-family: 'Inter', sans-serif;
        background-color: #F9FAFB;
    }
    .sidebar .sidebar-content {
        background-color: #FFFFFF;
        padding: 1.5rem;
        border-radius: 0.5rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    h1, h2, h3 {
        color: #1F2937;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 1rem;
        background-color: #FFFFFF;
        padding: 0.5rem;
        border-radius: 0.5rem;
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 1.1rem;
        padding: 0.75rem 1.5rem;
        border-radius: 0.5rem;
        color: #4B5563;
    }
    .stTabs [data-baseweb="tab"].stTabsActive {
        background-color: #3B82F6;
        color: #FFFFFF;
    }
    .stButton>button {
        background-color: #3B82F6;
        color: white;
        border-radius: 0.5rem;
        padding: 0.5rem 1rem;
        font-weight: 500;
    }
    .stButton>button:hover {
        background-color: #2563EB;
    }
    </style>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("<h2 class='text-xl font-semibold mb-4'>Bộ Lọc</h2>", unsafe_allow_html=True)
    
    # File uploaders
    prev_year_file = st.file_uploader("File Năm Trước", type=['xlsx'], key="prev_file")
    curr_year_file = st.file_uploader("File Năm Nay", type=['xlsx'], key="curr_file")

    if prev_year_file and curr_year_file:
        # Load and validate data with caching
        df_prev = load_and_validate_data(prev_year_file, "Năm Trước")
        df_curr = load_and_validate_data(curr_year_file, "Năm Nay")
        
        if df_prev is None or df_curr is None:
            st.error("Không thể tải dữ liệu. Vui lòng kiểm tra file.")
            st.stop()

        # Filters
        months = list(range(1, 13))
        selected_months = st.multiselect("Tháng", months, default=months, key="month_select")
        
        # Customer filter
        customer_data = pd.concat([
            df_prev[['Customer', 'Name']],
            df_curr[['Customer', 'Name']]
        ]).drop_duplicates(subset=['Customer']).set_index('Customer')
        customer_options = [
            f"{cust} - {row['Name'] if pd.notna(row['Name']) else 'Không có tên'}"
            for cust, row in customer_data.iterrows()
        ]
        selected_customer_display = st.multiselect(
            "Khách Hàng", customer_options, default=[], key="customer_select"
        )
        selected_customers = [
            cust for cust, row in customer_data.iterrows()
            if f"{cust} - {row['Name'] if pd.notna(row['Name']) else 'Không có tên'}" in selected_customer_display
        ]

        # Material filter
        material_data = pd.concat([
            df_prev[['Material', 'Item Description']],
            df_curr[['Material', 'Item Description']]
        ]).drop_duplicates(subset=['Material']).set_index('Material')
        material_options = [
            f"{mat} - {row['Item Description'] if pd.notna(row['Item Description']) else 'Không có mô tả'}"
            for mat, row in material_data.iterrows()
        ]
        selected_material_display = st.multiselect(
            "Sản Phẩm", material_options, default=[], key="material_select"
        )
        selected_materials = [
            mat for mat, row in material_data.iterrows()
            if f"{mat} - {row['Item Description'] if pd.notna(row['Item Description']) else 'Không có mô tả'}" in selected_material_display
        ]

        # TDV filter
        tdv_options = sorted(pd.concat([
            df_prev['Tên TDV'], df_curr['Tên TDV']
        ]).drop_duplicates().dropna().tolist())
        selected_tdvs = st.multiselect(
            "Tên TDV", tdv_options, default=tdv_options, key="tdv_select"
        )

        filters = {
            'months': selected_months or months,
            'customers': selected_customers,
            'materials': selected_materials,
            'tdvs': selected_tdvs
        }

        # Summary Info
        st.markdown("<h2 class='text-xl font-semibold mt-6 mb-4'>Thông Tin</h2>", unsafe_allow_html=True)
        summary = get_summary_info(df_prev, df_curr)
        if summary is not None:
            st.dataframe(
                summary.style.format({
                    'Năm Trước': '{:,.0f}',
                    'Năm Nay': '{:,.0f}',
                    'Thông tin': '{}'
                }).set_table_styles([
                    {'selector': 'th', 'props': [('background-color', '#E5E7EB'), ('font-weight', 'bold')]},
                    {'selector': 'td', 'props': [('border', '1px solid #E5E7EB')]}
                ]),
                use_container_width=True
            )
            # Export summary to Excel
            buffer = io.BytesIO()
            summary.to_excel(buffer, index=False, engine='openpyxl')
            buffer.seek(0)
            st.download_button(
                label="Tải Bảng Tóm Tắt (Excel)",
                data=buffer,
                file_name="summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.info("Vui lòng tải lên cả hai file dữ liệu.")
        st.stop()

# Main content
st.markdown(f"<h1 class='text-2xl font-bold mb-6'>{config['app']['title']}</h1>", unsafe_allow_html=True)

# Tabs
tabs = st.tabs(["Tổng Quan", "Sản Phẩm", "Khách Hàng", "TDV"])
with tabs[0]:
    plot_overview(df_prev, df_curr, config, filters)
with tabs[1]:
    plot_product_analysis(df_prev, df_curr, config, filters)
with tabs[2]:
    plot_customer_analysis(df_prev, df_curr, config, filters)
with tabs[3]:
    plot_tdv_analysis(df_prev, df_curr, config, filters)