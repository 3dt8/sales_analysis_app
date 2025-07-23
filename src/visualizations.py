import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from src.utils import log_error
import io

@st.cache_data
def calculate_rfm(df, latest_date):
    """Calculate RFM scores."""
    try:
        if df is None or df.empty or not all(col in df.columns for col in ['Customer', 'Billing Date', 'DS Ðã Trừ CK', 'Name']):
            log_error("Invalid input data for RFM")
            return None
        rfm = df.groupby(['Customer', 'Name']).agg({
            'Billing Date': lambda x: (latest_date - x.max()).days,
            'Customer': 'count',
            'DS Ðã Trừ CK': 'sum'
        }).rename(columns={'Billing Date': 'Recency', 'Customer': 'Frequency', 'DS Ðã Trừ CK': 'Monetary'})
        for col in ['Recency', 'Frequency', 'Monetary']:
            if rfm[col].nunique() <= 1:
                rfm[f'{col}_Score'] = 1
            else:
                rfm[f'{col}_Score'] = pd.qcut(rfm[col], q=4, labels=[4, 3, 2, 1] if col == 'Recency' else [1, 2, 3, 4], duplicates='drop')
        rfm['RFM_Segment'] = rfm['Recency_Score'].astype(str) + rfm['Frequency_Score'].astype(str) + rfm['Monetary_Score'].astype(str)
        return rfm.reset_index()
    except Exception as e:
        log_error(f"Error calculating RFM: {str(e)}")
        st.error(f"Error calculating RFM: {str(e)}")
        return None

@st.cache_data
def get_summary_info(df_prev, df_curr):
    """Create summary statistics."""
    try:
        if df_prev is None or df_curr is None or df_prev.empty or df_curr.empty:
            log_error("Invalid input data for summary")
            return None
        summary = {
            'Thông tin': ['Số dòng', 'Số tháng', 'Số KH', 'Tổng doanh số', 'Số sản phẩm', 'Số TDV'],
            'Năm Trước': [
                len(df_prev), len(df_prev['Billing Date'].dt.to_period('M').unique()),
                df_prev['Customer'].nunique(), df_prev['DS Ðã Trừ CK'].sum(),
                df_prev['Material'].nunique(), df_prev['Tên TDV'].nunique()
            ],
            'Năm Nay': [
                len(df_curr), len(df_curr['Billing Date'].dt.to_period('M').unique()),
                df_curr['Customer'].nunique(), df_curr['DS Ðã Trừ CK'].sum(),
                df_curr['Material'].nunique(), df_curr['Tên TDV'].nunique()
            ]
        }
        return pd.DataFrame(summary)
    except Exception as e:
        log_error(f"Error creating summary: {str(e)}")
        st.error(f"Error creating summary: {str(e)}")
        return None

def plot_overview(df_prev, df_curr, config, filters):
    """Plot Overview tab."""
    try:
        st.markdown("<h2 class='text-xl font-semibold mb-4'>Tổng Quan</h2>", unsafe_allow_html=True)
        df_prev_filtered = df_prev[df_prev['Billing Date'].dt.month.isin(filters['months'])].copy()
        df_curr_filtered = df_curr[df_curr['Billing Date'].dt.month.isin(filters['months'])].copy()
        if filters['customers']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Customer'].isin(filters['customers'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Customer'].isin(filters['customers'])]
        if filters['materials']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Material'].isin(filters['materials'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Material'].isin(filters['materials'])]
        if filters['tdvs']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Tên TDV'].isin(filters['tdvs'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Tên TDV'].isin(filters['tdvs'])]
        if df_prev_filtered.empty and df_curr_filtered.empty:
            st.warning("Không có dữ liệu sau khi lọc. Vui lòng kiểm tra bộ lọc.")
            return

        # Key metrics
        total_prev = df_prev_filtered['DS Ðã Trừ CK'].sum() if not df_prev_filtered.empty else 0
        total_curr = df_curr_filtered['DS Ðã Trừ CK'].sum() if not df_curr_filtered.empty else 0
        growth = (total_curr - total_prev) / total_prev * 100 if total_prev else 0
        num_orders_prev = df_prev_filtered.groupby(['Customer', df_prev_filtered['Billing Date'].dt.date]).ngroups if not df_prev_filtered.empty else 0
        num_orders_curr = df_curr_filtered.groupby(['Customer', df_curr_filtered['Billing Date'].dt.date]).ngroups if not df_curr_filtered.empty else 0

        col1, col2 = st.columns(2)
        with col1:
            st.metric("Doanh thu cùng kỳ", f"{total_curr:,.0f} VND", f"{growth:+.2f}%")
        with col2:
            st.metric("Số đơn hàng", num_orders_curr, num_orders_curr - num_orders_prev)

        # Customer sales table
        customer_prev = df_prev_filtered.groupby('Customer').agg({'DS Ðã Trừ CK': 'sum', 'Name': 'first'}).reset_index() if not df_prev_filtered.empty else pd.DataFrame(columns=['Customer', 'Name', 'DS Ðã Trừ CK'])
        customer_curr = df_curr_filtered.groupby('Customer').agg({'DS Ðã Trừ CK': 'sum', 'Name': 'first'}).reset_index() if not df_curr_filtered.empty else pd.DataFrame(columns=['Customer', 'Name', 'DS Ðã Trừ CK'])
        customer_merged = customer_prev.merge(customer_curr, on='Customer', how='outer', suffixes=('_prev', '_curr')).fillna(0)
        customer_merged['Tăng trưởng (%)'] = ((customer_merged['DS Ðã Trừ CK_curr'] - customer_merged['DS Ðã Trừ CK_prev']) / customer_merged['DS Ðã Trừ CK_prev'] * 100).replace([float('inf'), -float('inf')], 0)
        total_row = pd.DataFrame({
            'Customer': [''], 'Name_prev': [''], 'DS Ðã Trừ CK_prev': [customer_merged['DS Ðã Trừ CK_prev'].sum()],
            'Name_curr': [''], 'DS Ðã Trừ CK_curr': [customer_merged['DS Ðã Trừ CK_curr'].sum()],
            'Tăng trưởng (%)': [(customer_merged['DS Ðã Trừ CK_curr'].sum() - customer_merged['DS Ðã Trừ CK_prev'].sum()) / customer_merged['DS Ðã Trừ CK_prev'].sum() * 100 if customer_merged['DS Ðã Trừ CK_prev'].sum() else 0]
        })
        customer_table = pd.concat([customer_merged[['Customer', 'Name_prev', 'DS Ðã Trừ CK_prev', 'Name_curr', 'DS Ðã Trừ CK_curr', 'Tăng trưởng (%)']], total_row]).fillna('')
        st.markdown("<h3 class='text-lg font-semibold mb-2'>Doanh số Khách hàng</h3>", unsafe_allow_html=True)
        st.dataframe(
            customer_table.style.format({
                'DS Ðã Trừ CK_prev': '{:,.0f}',
                'DS Ðã Trừ CK_curr': '{:,.0f}',
                'Tăng trưởng (%)': '{:.2f}%'
            }).set_table_styles([
                {'selector': 'th', 'props': [('background-color', '#E5E7EB'), ('font-weight', 'bold')]},
                {'selector': 'td', 'props': [('border', '1px solid #E5E7EB')]}
            ]),
            use_container_width=True
        )
        buffer = io.BytesIO()
        customer_table.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)
        st.download_button(
            label="Tải Doanh số Khách hàng (Excel)",
            data=buffer,
            file_name="customer_sales.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Product sales table
        product_prev = df_prev_filtered.groupby('Material').agg({'DS Ðã Trừ CK': 'sum', 'Item Description': 'first'}).reset_index() if not df_prev_filtered.empty else pd.DataFrame(columns=['Material', 'Item Description', 'DS Ðã Trừ CK'])
        product_curr = df_curr_filtered.groupby('Material').agg({'DS Ðã Trừ CK': 'sum', 'Item Description': 'first'}).reset_index() if not df_curr_filtered.empty else pd.DataFrame(columns=['Material', 'Item Description', 'DS Ðã Trừ CK'])
        product_merged = product_prev.merge(product_curr, on='Material', how='outer', suffixes=('_prev', '_curr')).fillna(0)
        product_merged['Tăng trưởng (%)'] = ((product_merged['DS Ðã Trừ CK_curr'] - product_merged['DS Ðã Trừ CK_prev']) / product_merged['DS Ðã Trừ CK_prev'] * 100).replace([float('inf'), -float('inf')], 0)
        total_row = pd.DataFrame({
            'Material': [''], 'Item Description_prev': [''], 'DS Ðã Trừ CK_prev': [product_merged['DS Ðã Trừ CK_prev'].sum()],
            'Item Description_curr': [''], 'DS Ðã Trừ CK_curr': [product_merged['DS Ðã Trừ CK_curr'].sum()],
            'Tăng trưởng (%)': [(product_merged['DS Ðã Trừ CK_curr'].sum() - product_merged['DS Ðã Trừ CK_prev'].sum()) / product_merged['DS Ðã Trừ CK_prev'].sum() * 100 if product_merged['DS Ðã Trừ CK_prev'].sum() else 0]
        })
        product_table = pd.concat([product_merged[['Material', 'Item Description_prev', 'DS Ðã Trừ CK_prev', 'Item Description_curr', 'DS Ðã Trừ CK_curr', 'Tăng trưởng (%)']], total_row]).fillna('')
        st.markdown("<h3 class='text-lg font-semibold mb-2'>Doanh số Sản phẩm</h3>", unsafe_allow_html=True)
        st.dataframe(
            product_table.style.format({
                'DS Ðã Trừ CK_prev': '{:,.0f}',
                'DS Ðã Trừ CK_curr': '{:,.0f}',
                'Tăng trưởng (%)': '{:.2f}%'
            }).set_table_styles([
                {'selector': 'th', 'props': [('background-color', '#E5E7EB'), ('font-weight', 'bold')]},
                {'selector': 'td', 'props': [('border', '1px solid #E5E7EB')]}
            ]),
            use_container_width=True
        )
        buffer = io.BytesIO()
        product_table.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)
        st.download_button(
            label="Tải Doanh số Sản phẩm (Excel)",
            data=buffer,
            file_name="product_sales.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Sales trend chart
        trend_prev = df_prev_filtered.groupby(df_prev_filtered['Billing Date'].dt.month)['DS Ðã Trừ CK'].sum().reset_index(name='DS Ðã Trừ CK') if not df_prev_filtered.empty else pd.DataFrame({'Billing Date': [], 'DS Ðã Trừ CK': []})
        trend_curr = df_curr_filtered.groupby(df_curr_filtered['Billing Date'].dt.month)['DS Ðã Trừ CK'].sum().reset_index(name='DS Ðã Trừ CK') if not df_curr_filtered.empty else pd.DataFrame({'Billing Date': [], 'DS Ðã Trừ CK': []})
        fig_sales = go.Figure()
        if not trend_prev.empty:
            fig_sales.add_trace(go.Bar(x=trend_prev['Billing Date'], y=trend_prev['DS Ðã Trừ CK'], name='Năm Trước', marker_color=config["colors"]["prev_year"]))
        if not trend_curr.empty:
            fig_sales.add_trace(go.Bar(x=trend_curr['Billing Date'], y=trend_curr['DS Ðã Trừ CK'], name='Năm Nay', marker_color=config["colors"]["curr_year"]))
        fig_sales.update_layout(barmode='group', xaxis_title="Tháng", yaxis_title="Doanh số (VND)", height=400, margin=dict(t=50))
        st.markdown("<h3 class='text-lg font-semibold mb-2'>Biểu đồ Doanh số</h3>", unsafe_allow_html=True)
        st.plotly_chart(fig_sales, use_container_width=True)
    except Exception as e:
        log_error(f"Error in plot_overview: {str(e)}")
        st.error(f"Lỗi tab Tổng Quan: {str(e)}")

def plot_product_analysis(df_prev, df_curr, config, filters):
    """Plot Product tab."""
    try:
        st.markdown("<h2 class='text-xl font-semibold mb-4'>Sản Phẩm</h2>", unsafe_allow_html=True)
        df_prev_filtered = df_prev[df_prev['Billing Date'].dt.month.isin(filters['months'])].copy()
        df_curr_filtered = df_curr[df_curr['Billing Date'].dt.month.isin(filters['months'])].copy()
        if filters['customers']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Customer'].isin(filters['customers'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Customer'].isin(filters['customers'])]
        if filters['materials']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Material'].isin(filters['materials'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Material'].isin(filters['materials'])]
        if filters['tdvs']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Tên TDV'].isin(filters['tdvs'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Tên TDV'].isin(filters['tdvs'])]
        if df_prev_filtered.empty and df_curr_filtered.empty:
            st.warning("Không có dữ liệu sau khi lọc. Vui lòng kiểm tra bộ lọc.")
            return

        # Product sales table
        product_prev = df_prev_filtered.groupby('Material').agg({'DS Ðã Trừ CK': 'sum', 'Item Description': 'first'}).reset_index() if not df_prev_filtered.empty else pd.DataFrame(columns=['Material', 'Item Description', 'DS Ðã Trừ CK'])
        product_curr = df_curr_filtered.groupby('Material').agg({'DS Ðã Trừ CK': 'sum', 'Item Description': 'first'}).reset_index() if not df_curr_filtered.empty else pd.DataFrame(columns=['Material', 'Item Description', 'DS Ðã Trừ CK'])
        product_merged = product_prev.merge(product_curr, on='Material', how='outer', suffixes=('_prev', '_curr')).fillna(0)
        product_merged['Tăng trưởng (%)'] = ((product_merged['DS Ðã Trừ CK_curr'] - product_merged['DS Ðã Trừ CK_prev']) / product_merged['DS Ðã Trừ CK_prev'] * 100).replace([float('inf'), -float('inf')], 0)
        total_row = pd.DataFrame({
            'Material': [''], 'Item Description_prev': [''], 'DS Ðã Trừ CK_prev': [product_merged['DS Ðã Trừ CK_prev'].sum()],
            'Item Description_curr': [''], 'DS Ðã Trừ CK_curr': [product_merged['DS Ðã Trừ CK_curr'].sum()],
            'Tăng trưởng (%)': [(product_merged['DS Ðã Trừ CK_curr'].sum() - product_merged['DS Ðã Trừ CK_prev'].sum()) / product_merged['DS Ðã Trừ CK_prev'].sum() * 100 if product_merged['DS Ðã Trừ CK_prev'].sum() else 0]
        })
        product_table = pd.concat([product_merged[['Material', 'Item Description_prev', 'DS Ðã Trừ CK_prev', 'Item Description_curr', 'DS Ðã Trừ CK_curr', 'Tăng trưởng (%)']], total_row]).fillna('')
        st.markdown("<h3 class='text-lg font-semibold mb-2'>Doanh số Sản phẩm</h3>", unsafe_allow_html=True)
        st.dataframe(
            product_table.style.format({
                'DS Ðã Trừ CK_prev': '{:,.0f}',
                'DS Ðã Trừ CK_curr': '{:,.0f}',
                'Tăng trưởng (%)': '{:.2f}%'
            }).set_table_styles([
                {'selector': 'th', 'props': [('background-color', '#E5E7EB'), ('font-weight', 'bold')]},
                {'selector': 'td', 'props': [('border', '1px solid #E5E7EB')]}
            ]),
            use_container_width=True
        )
        buffer = io.BytesIO()
        product_table.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)
        st.download_button(
            label="Tải Doanh số Sản phẩm (Excel)",
            data=buffer,
            file_name="product_sales_product_tab.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Product sales chart
        if filters['materials']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Material'].isin(filters['materials'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Material'].isin(filters['materials'])]
            if not df_prev_filtered.empty or not df_curr_filtered.empty:
                trend_prev = df_prev_filtered.groupby([df_prev_filtered['Billing Date'].dt.month, 'Material'])['DS Ðã Trừ CK'].sum().reset_index() if not df_prev_filtered.empty else pd.DataFrame({'Billing Date': [], 'Material': [], 'DS Ðã Trừ CK': []})
                trend_curr = df_curr_filtered.groupby([df_curr_filtered['Billing Date'].dt.month, 'Material'])['DS Ðã Trừ CK'].sum().reset_index() if not df_curr_filtered.empty else pd.DataFrame({'Billing Date': [], 'Material': [], 'DS Ðã Trừ CK': []})
                fig = go.Figure()
                for mat in filters['materials']:
                    desc = df_curr[df_curr['Material'] == mat]['Item Description'].iloc[0] if not df_curr[df_curr['Material'] == mat].empty else mat
                    prev_data = trend_prev[trend_prev['Material'] == mat] if not trend_prev.empty else pd.DataFrame({'Billing Date': [], 'DS Ðã Trừ CK': []})
                    curr_data = trend_curr[trend_curr['Material'] == mat] if not trend_curr.empty else pd.DataFrame({'Billing Date': [], 'DS Ðã Trừ CK': []})
                    if not prev_data.empty:
                        fig.add_trace(go.Bar(x=prev_data['Billing Date'], y=prev_data['DS Ðã Trừ CK'], name=f"{desc} (Năm Trước)", marker_color=config["colors"]["prev_year"]))
                    if not curr_data.empty:
                        fig.add_trace(go.Bar(x=curr_data['Billing Date'], y=curr_data['DS Ðã Trừ CK'], name=f"{desc} (Năm Nay)", marker_color=config["colors"]["curr_year"]))
                fig.update_layout(barmode='group', xaxis_title="Tháng", yaxis_title="Doanh số (VND)", height=400, margin=dict(t=50))
                st.markdown("<h3 class='text-lg font-semibold mb-2'>Biểu đồ Doanh số Sản phẩm</h3>", unsafe_allow_html=True)
                st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        log_error(f"Error in plot_product_analysis: {str(e)}")
        st.error(f"Lỗi tab Sản Phẩm: {str(e)}")

def plot_customer_analysis(df_prev, df_curr, config, filters):
    """Plot Customer tab."""
    try:
        st.markdown("<h2 class='text-xl font-semibold mb-4'>Khách Hàng</h2>", unsafe_allow_html=True)
        df_prev_filtered = df_prev[df_prev['Billing Date'].dt.month.isin(filters['months'])].copy()
        df_curr_filtered = df_curr[df_curr['Billing Date'].dt.month.isin(filters['months'])].copy()
        if filters['customers']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Customer'].isin(filters['customers'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Customer'].isin(filters['customers'])]
        if filters['materials']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Material'].isin(filters['materials'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Material'].isin(filters['materials'])]
        if filters['tdvs']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Tên TDV'].isin(filters['tdvs'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Tên TDV'].isin(filters['tdvs'])]
        if df_prev_filtered.empty and df_curr_filtered.empty:
            st.warning("Không có dữ liệu sau khi lọc. Vui lòng kiểm tra bộ lọc.")
            return

        # Customer sales table
        customer_prev = df_prev_filtered.groupby('Customer').agg({'DS Ðã Trừ CK': 'sum', 'Name': 'first'}).reset_index() if not df_prev_filtered.empty else pd.DataFrame(columns=['Customer', 'Name', 'DS Ðã Trừ CK'])
        customer_curr = df_curr_filtered.groupby('Customer').agg({'DS Ðã Trừ CK': 'sum', 'Name': 'first'}).reset_index() if not df_curr_filtered.empty else pd.DataFrame(columns=['Customer', 'Name', 'DS Ðã Trừ CK'])
        customer_merged = customer_prev.merge(customer_curr, on='Customer', how='outer', suffixes=('_prev', '_curr')).fillna(0)
        customer_merged['Tăng trưởng (%)'] = ((customer_merged['DS Ðã Trừ CK_curr'] - customer_merged['DS Ðã Trừ CK_prev']) / customer_merged['DS Ðã Trừ CK_prev'] * 100).replace([float('inf'), -float('inf')], 0)
        total_row = pd.DataFrame({
            'Customer': [''], 'Name_prev': [''], 'DS Ðã Trừ CK_prev': [customer_merged['DS Ðã Trừ CK_prev'].sum()],
            'Name_curr': [''], 'DS Ðã Trừ CK_curr': [customer_merged['DS Ðã Trừ CK_curr'].sum()],
            'Tăng trưởng (%)': [(customer_merged['DS Ðã Trừ CK_curr'].sum() - customer_merged['DS Ðã Trừ CK_prev'].sum()) / customer_merged['DS Ðã Trừ CK_prev'].sum() * 100 if customer_merged['DS Ðã Trừ CK_prev'].sum() else 0]
        })
        customer_table = pd.concat([customer_merged[['Customer', 'Name_prev', 'DS Ðã Trừ CK_prev', 'Name_curr', 'DS Ðã Trừ CK_curr', 'Tăng trưởng (%)']], total_row]).fillna('')
        st.markdown("<h3 class='text-lg font-semibold mb-2'>Doanh số Khách hàng</h3>", unsafe_allow_html=True)
        st.dataframe(
            customer_table.style.format({
                'DS Ðã Trừ CK_prev': '{:,.0f}',
                'DS Ðã Trừ CK_curr': '{:,.0f}',
                'Tăng trưởng (%)': '{:.2f}%'
            }).set_table_styles([
                {'selector': 'th', 'props': [('background-color', '#E5E7EB'), ('font-weight', 'bold')]},
                {'selector': 'td', 'props': [('border', '1px solid #E5E7EB')]}
            ]),
            use_container_width=True
        )
        buffer = io.BytesIO()
        customer_table.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)
        st.download_button(
            label="Tải Doanh số Khách hàng (Excel)",
            data=buffer,
            file_name="customer_sales_customer_tab.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # RFM Analysis
        if filters['customers']:
            latest_date = max(df_curr['Billing Date'].max(), df_prev['Billing Date'].max()) if not df_curr.empty and not df_prev.empty else pd.Timestamp.now()
            rfm_curr = calculate_rfm(df_curr_filtered, latest_date) if not df_curr_filtered.empty else None
            if rfm_curr is not None:
                rfm_curr = rfm_curr.groupby('Customer').agg({
                    'Recency': 'min', 'Frequency': 'sum', 'Monetary': 'sum',
                    'Name': 'first', 'RFM_Segment': 'first'
                }).reset_index()
                total_row = pd.DataFrame({
                    'Customer': [''], 'Name': [''], 'Recency': [''],
                    'Frequency': [rfm_curr['Frequency'].sum()],
                    'Monetary': [rfm_curr['Monetary'].sum()], 'RFM_Segment': ['']
                })
                rfm_table = pd.concat([rfm_curr[['Customer', 'Name', 'Recency', 'Frequency', 'Monetary', 'RFM_Segment']], total_row]).fillna('')
                st.markdown("<h3 class='text-lg font-semibold mb-2'>Phân nhóm RFM</h3>", unsafe_allow_html=True)
                st.dataframe(
                    rfm_table.style.format({
                        'Frequency': '{:,.0f}',
                        'Monetary': '{:,.0f}'
                    }).set_table_styles([
                        {'selector': 'th', 'props': [('background-color', '#E5E7EB'), ('font-weight', 'bold')]},
                        {'selector': 'td', 'props': [('border', '1px solid #E5E7EB')]}
                    ]),
                    use_container_width=True
                )
                buffer = io.BytesIO()
                rfm_table.to_excel(buffer, index=False, engine='openpyxl')
                buffer.seek(0)
                st.download_button(
                    label="Tải Phân nhóm RFM (Excel)",
                    data=buffer,
                    file_name="rfm_analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Vui lòng chọn ít nhất một khách hàng để tính RFM.")

        # Customer sales chart
        if filters['customers']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Customer'].isin(filters['customers'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Customer'].isin(filters['customers'])]
            if not df_prev_filtered.empty or not df_curr_filtered.empty:
                trend_prev = df_prev_filtered.groupby([df_prev_filtered['Billing Date'].dt.month, 'Customer'])['DS Ðã Trừ CK'].sum().reset_index() if not df_prev_filtered.empty else pd.DataFrame({'Billing Date': [], 'Customer': [], 'DS Ðã Trừ CK': []})
                trend_curr = df_curr_filtered.groupby([df_curr_filtered['Billing Date'].dt.month, 'Customer'])['DS Ðã Trừ CK'].sum().reset_index() if not df_curr_filtered.empty else pd.DataFrame({'Billing Date': [], 'Customer': [], 'DS Ðã Trừ CK': []})
                fig = go.Figure()
                for cust in filters['customers']:
                    name = df_curr[df_curr['Customer'] == cust]['Name'].iloc[0] if not df_curr[df_curr['Customer'] == cust].empty else cust
                    prev_data = trend_prev[trend_prev['Customer'] == cust] if not trend_prev.empty else pd.DataFrame({'Billing Date': [], 'DS Ðã Trừ CK': []})
                    curr_data = trend_curr[trend_curr['Customer'] == cust] if not trend_curr.empty else pd.DataFrame({'Billing Date': [], 'DS Ðã Trừ CK': []})
                    if not prev_data.empty:
                        fig.add_trace(go.Bar(x=prev_data['Billing Date'], y=prev_data['DS Ðã Trừ CK'], name=f"{name} (Năm Trước)", marker_color=config["colors"]["prev_year"]))
                    if not curr_data.empty:
                        fig.add_trace(go.Bar(x=curr_data['Billing Date'], y=curr_data['DS Ðã Trừ CK'], name=f"{name} (Năm Nay)", marker_color=config["colors"]["curr_year"]))
                fig.update_layout(barmode='group', xaxis_title="Tháng", yaxis_title="Doanh số (VND)", height=400, margin=dict(t=50))
                st.markdown("<h3 class='text-lg font-semibold mb-2'>Biểu đồ Doanh số Khách hàng</h3>", unsafe_allow_html=True)
                st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        log_error(f"Error in plot_customer_analysis: {str(e)}")
        st.error(f"Lỗi tab Khách Hàng: {str(e)}")

def plot_tdv_analysis(df_prev, df_curr, config, filters):
    """Plot TDV tab."""
    try:
        st.markdown("<h2 class='text-xl font-semibold mb-4'>TDV</h2>", unsafe_allow_html=True)
        df_prev_filtered = df_prev[df_prev['Billing Date'].dt.month.isin(filters['months'])].copy()
        df_curr_filtered = df_curr[df_curr['Billing Date'].dt.month.isin(filters['months'])].copy()
        if filters['tdvs']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Tên TDV'].isin(filters['tdvs'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Tên TDV'].isin(filters['tdvs'])]
        if filters['materials']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Material'].isin(filters['materials'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Material'].isin(filters['materials'])]
        if filters['customers']:
            df_prev_filtered = df_prev_filtered[df_prev_filtered['Customer'].isin(filters['customers'])]
            df_curr_filtered = df_curr_filtered[df_curr_filtered['Customer'].isin(filters['customers'])]
        if df_prev_filtered.empty and df_curr_filtered.empty:
            st.warning("Không có dữ liệu sau khi lọc. Vui lòng kiểm tra bộ lọc.")
            return

        # Pie charts for TDV sales distribution
        total_prev = df_prev_filtered['DS Ðã Trừ CK'].sum() if not df_prev_filtered.empty else 0
        total_curr = df_curr_filtered['DS Ðã Trừ CK'].sum() if not df_curr_filtered.empty else 0
        tdv_prev = df_prev_filtered.groupby('Tên TDV')['DS Ðã Trừ CK'].sum().reset_index() if not df_prev_filtered.empty else pd.DataFrame(columns=['Tên TDV', 'DS Ðã Trừ CK'])
        tdv_curr = df_curr_filtered.groupby('Tên TDV')['DS Ðã Trừ CK'].sum().reset_index() if not df_curr_filtered.empty else pd.DataFrame(columns=['Tên TDV', 'DS Ðã Trừ CK'])
        fig_pie_prev = go.Figure(data=[go.Pie(labels=tdv_prev['Tên TDV'], values=tdv_prev['DS Ðã Trừ CK'], name='Năm Trước')]) if not tdv_prev.empty else go.Figure()
        fig_pie_curr = go.Figure(data=[go.Pie(labels=tdv_curr['Tên TDV'], values=tdv_curr['DS Ðã Trừ CK'], name='Năm Nay')]) if not tdv_curr.empty else go.Figure()
        st.markdown("<h3 class='text-lg font-semibold mb-2'>Tỷ trọng Doanh số TDV</h3>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(fig_pie_prev, use_container_width=True)
        with col2:
            st.plotly_chart(fig_pie_curr, use_container_width=True)

        # TDV sales by month
        trend_prev = df_prev_filtered.groupby([df_prev_filtered['Billing Date'].dt.month, 'Tên TDV'])['DS Ðã Trừ CK'].sum().reset_index() if not df_prev_filtered.empty else pd.DataFrame({'Billing Date': [], 'Tên TDV': [], 'DS Ðã Trừ CK': []})
        trend_curr = df_curr_filtered.groupby([df_curr_filtered['Billing Date'].dt.month, 'Tên TDV'])['DS Ðã Trừ CK'].sum().reset_index() if not df_curr_filtered.empty else pd.DataFrame({'Billing Date': [], 'Tên TDV': [], 'DS Ðã Trừ CK': []})
        fig = go.Figure()
        for tdv in filters.get('tdvs', df_curr['Tên TDV'].unique()) if not df_curr.empty else []:
            prev_data = trend_prev[trend_prev['Tên TDV'] == tdv] if not trend_prev.empty else pd.DataFrame({'Billing Date': [], 'DS Ðã Trừ CK': []})
            curr_data = trend_curr[trend_curr['Tên TDV'] == tdv] if not trend_curr.empty else pd.DataFrame({'Billing Date': [], 'DS Ðã Trừ CK': []})
            if not prev_data.empty:
                fig.add_trace(go.Bar(x=prev_data['Billing Date'], y=prev_data['DS Ðã Trừ CK'], name=f"{tdv} (Năm Trước)", marker_color=config["colors"]["prev_year"]))
            if not curr_data.empty:
                fig.add_trace(go.Bar(x=curr_data['Billing Date'], y=curr_data['DS Ðã Trừ CK'], name=f"{tdv} (Năm Nay)", marker_color=config["colors"]["curr_year"]))
        fig.update_layout(barmode='group', xaxis_title="Tháng", yaxis_title="Doanh số (VND)", height=400, margin=dict(t=50))
        st.markdown("<h3 class='text-lg font-semibold mb-2'>Doanh số TDV theo tháng</h3>", unsafe_allow_html=True)
        st.plotly_chart(fig, use_container_width=True)

        # Customer sales by TDV
        if filters['tdvs']:
            for tdv in filters['tdvs']:
                st.markdown(f"<h3 class='text-lg font-semibold mb-2'>TDV: {tdv}</h3>", unsafe_allow_html=True)
                customer_prev = df_prev_filtered[df_prev_filtered['Tên TDV'] == tdv].groupby('Customer').agg({'DS Ðã Trừ CK': 'sum', 'Name': 'first'}).reset_index() if not df_prev_filtered.empty else pd.DataFrame(columns=['Customer', 'Name', 'DS Ðã Trừ CK'])
                customer_curr = df_curr_filtered[df_curr_filtered['Tên TDV'] == tdv].groupby('Customer').agg({'DS Ðã Trừ CK': 'sum', 'Name': 'first'}).reset_index() if not df_curr_filtered.empty else pd.DataFrame(columns=['Customer', 'Name', 'DS Ðã Trừ CK'])
                merged_customer = customer_prev.merge(customer_curr, on='Customer', how='outer', suffixes=('_prev', '_curr')).fillna(0)
                merged_customer['Tăng trưởng (%)'] = ((merged_customer['DS Ðã Trừ CK_curr'] - merged_customer['DS Ðã Trừ CK_prev']) / merged_customer['DS Ðã Trừ CK_prev'] * 100).replace([float('inf'), -float('inf')], 0)
                total_row = pd.DataFrame({
                    'Customer': [''], 'Name_prev': [''], 'DS Ðã Trừ CK_prev': [merged_customer['DS Ðã Trừ CK_prev'].sum()],
                    'Name_curr': [''], 'DS Ðã Trừ CK_curr': [merged_customer['DS Ðã Trừ CK_curr'].sum()],
                    'Tăng trưởng (%)': [(merged_customer['DS Ðã Trừ CK_curr'].sum() - merged_customer['DS Ðã Trừ CK_prev'].sum()) / merged_customer['DS Ðã Trừ CK_prev'].sum() * 100 if merged_customer['DS Ðã Trừ CK_prev'].sum() else 0]
                })
                customer_table = pd.concat([merged_customer[['Customer', 'Name_prev', 'DS Ðã Trừ CK_prev', 'Name_curr', 'DS Ðã Trừ CK_curr', 'Tăng trưởng (%)']], total_row]).fillna('')
                st.dataframe(
                    customer_table.style.format({
                        'DS Ðã Trừ CK_prev': '{:,.0f}',
                        'DS Ðã Trừ CK_curr': '{:,.0f}',
                        'Tăng trưởng (%)': '{:.2f}%'
                    }).set_table_styles([
                        {'selector': 'th', 'props': [('background-color', '#E5E7EB'), ('font-weight', 'bold')]},
                        {'selector': 'td', 'props': [('border', '1px solid #E5E7EB')]}
                    ]),
                    use_container_width=True
                )
                buffer = io.BytesIO()
                customer_table.to_excel(buffer, index=False, engine='openpyxl')
                buffer.seek(0)
                st.download_button(
                    label=f"Tải Doanh số Khách hàng TDV {tdv} (Excel)",
                    data=buffer,
                    file_name=f"customer_sales_tdv_{tdv}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        log_error(f"Error in plot_tdv_analysis: {str(e)}")
        st.error(f"Lỗi tab TDV: {str(e)}")