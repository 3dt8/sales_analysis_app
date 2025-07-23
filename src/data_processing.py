import pandas as pd
import streamlit as st
from src.utils import log_error

@st.cache_data
def load_and_validate_data(file, year_label):
    """
    Tải và kiểm tra dữ liệu từ file Excel.
    
    Args:
        file: File Excel được tải lên
        year_label: Tên năm (Năm Trước/Năm Nay)
    
    Returns:
        pandas.DataFrame: DataFrame đã xử lý hoặc None nếu có lỗi
    """
    required_columns = [
        "Billing Date", "Customer", "Name", "Material", "Item Description",
        "Số Lượng", "Đơn Giá", "DS Ðã Trừ CK", "Program", "Product Hierarchy", "Tên TDV"
    ]
    
    try:
        df = pd.read_excel(file)
        
        # Kiểm tra các cột bắt buộc
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            log_error(f"Thiếu cột trong file {year_label}: {missing_cols}")
            st.error(f"File {year_label} thiếu cột: {missing_cols}")
            return None
        
        # Xử lý định dạng ngày (dd/mm/yyyy)
        df["Billing Date"] = pd.to_datetime(df["Billing Date"], dayfirst=True, errors="coerce")
        invalid_dates = df[df["Billing Date"].isna()]
        if not invalid_dates.empty:
            log_error(f"File {year_label} có {len(invalid_dates)} dòng ngày không hợp lệ: {invalid_dates.index.tolist()}")
            st.warning(f"File {year_label} có {len(invalid_dates)} dòng ngày không hợp lệ đã bị loại.")
        df = df.dropna(subset=["Billing Date"])
        
        # Làm sạch dữ liệu (chuyển sang int rồi sang str để loại bỏ thập phân)
        df["Customer"] = df["Customer"].astype(str).str.replace(r"\.0+$", "", regex=True)
        df["Material"] = df["Material"].astype(str).str.replace(r"\.0+$", "", regex=True)
        
        # Đảm bảo kiểu số
        numeric_cols = ["Số Lượng", "Đơn Giá", "DS Ðã Trừ CK"]
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        invalid_numeric = df[df[numeric_cols].isna().any(axis=1)]
        if not invalid_numeric.empty:
            log_error(f"File {year_label} có {len(invalid_numeric)} dòng số không hợp lệ: {invalid_numeric.index.tolist()}")
            st.warning(f"File {year_label} có {len(invalid_numeric)} dòng số không hợp lệ đã bị loại.")
        df = df.dropna(subset=numeric_cols)
        
        # Định dạng số
        df["DS Ðã Trừ CK"] = df["DS Ðã Trừ CK"].round(2)
        
        return df
    except Exception as e:
        log_error(f"Lỗi xử lý file {year_label}: {str(e)}")
        st.error(f"Lỗi xử lý file {year_label}: {str(e)}")
        return None