import streamlit as st
import pandas as pd
import io
import re

# Cấu hình giao diện
st.set_page_config(page_title="Tool Chuẩn Hóa BHXH (CT07)", page_icon="🏥", layout="wide")

st.title("🏥 Tool Chuẩn Hóa Giấy Chứng Nhận Hưởng BHXH (CT07)")
st.write("Tải lên file Excel để tự động cập nhật `MAU_SO`, `LOAI_GIAYTO`, trích xuất CCCD và lọc dữ liệu.")

# Hàm chuyển đổi định dạng ngày sang YYYYMMDD (ví dụ: 21/02/2021 -> 20210221)
def format_to_yyyymmdd(date_str):
    if pd.isna(date_str) or not date_str:
        return ""
    
    date_str = str(date_str).strip()
    
    # Xử lý trường hợp số bị dính đuôi .0 (ví dụ: 20260213.0)
    if date_str.endswith('.0'):
        date_str = date_str[:-2]
        
    # Nếu đã là 8 chữ số dạng YYYYMMDD
    if len(date_str) == 8 and date_str.isdigit():
        return date_str
    
    # Tìm dạng DD/MM/YYYY hoặc DD-MM-YYYY
    m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', date_str)
    if m:
        day, month, year = m.groups()
        return f"{year}{int(month):02d}{int(day):02d}"
    
    # Tìm dạng YYYY-MM-DD
    m2 = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', date_str)
    if m2:
        year, month, day = m2.groups()
        return f"{year}{int(month):02d}{int(day):02d}"
        
    return date_str

# Hàm trích xuất 12 số CCCD và Ngày cấp từ chuỗi văn bản
def extract_cccd_and_date(text):
    if pd.isna(text) or not isinstance(text, str):
        return None, None
    
    # Tìm dãy 12 số liên tiếp (CCCD)
    cccd_match = re.search(r'\b\d{12}\b', text)
    cccd = cccd_match.group(0) if cccd_match else None
    
    # Tìm ngày cấp dạng DD/MM/YYYY hoặc DD-MM-YYYY
    date_match = re.search(r'\b(\d{1,2}[/-]\d{1,2}[/-]\d{4})\b', text)
    date_val = format_to_yyyymmdd(date_match.group(0)) if date_match else None
    
    return cccd, date_val

# Khu vực tải file
file_excel = st.file_uploader("📂 Kéo thả hoặc chọn file Excel (CT07) cần xử lý", type=["xlsx"])

if file_excel:
    if st.button("🚀 Tiến Hành Chuẩn Hóa Dữ Liệu", type="primary"):
        with st.spinner("Đang đọc và xử lý dữ liệu..."):
            # Đọc dữ liệu từ sheet đầu tiên
            df = pd.read_excel(file_excel, sheet_name=0)

            # Ép kiểu dữ liệu dạng chuỗi cho các cột làm việc
            for col in ['SO_CCCD', 'NGAYCAP_CCCD', 'HO_TEN_CHA', 'HO_TEN_ME']:
                if col in df.columns:
                    df[col] = df[col].astype(str).replace({'nan': '', 'None': ''}).str.strip()

            # YÊU CẦU 1: Xóa dòng nếu KHÔNG CÓ CẢ họ tên Cha VÀ Mẹ
            total_before = len(df)
            has_cha = df['HO_TEN_CHA'].fillna('').str.strip() != ''
            has_me = df['HO_TEN_ME'].fillna('').str.strip() != ''
            
            # Chỉ giữ lại dòng có Cha HOẶC có Mẹ
            df = df[has_cha | has_me].copy()
            deleted_rows_count = total_before - len(df)

            # Quy tắc Gán mặc định: MAU_SO = 'CT07' và LOAI_GIAYTO = 1
            df['MAU_SO'] = 'CT07'
            df['LOAI_GIAYTO'] = 1

            # YÊU CẦU 2 & 3: Xử lý CCCD, Trích xuất & Định dạng NGAYCAP_CCCD thành YYYYMMDD
            count_fixed_cccd = 0
            for idx in df.index:
                # 1. Định dạng lại cột NGAYCAP_CCCD hiện có về dạng YYYYMMDD
                if 'NGAYCAP_CCCD' in df.columns and pd.notna(df.at[idx, 'NGAYCAP_CCCD']):
                    df.at[idx, 'NGAYCAP_CCCD'] = format_to_yyyymmdd(df.at[idx, 'NGAYCAP_CCCD'])

                # 2. Xử lý CCCD trống
                cccd_val = str(df.at[idx, 'SO_CCCD']).strip()
                if not cccd_val or cccd_val.lower() in ['', 'nan']:
                    text_cha = df.at[idx, 'HO_TEN_CHA']
                    text_me = df.at[idx, 'HO_TEN_ME']
                    
                    cccd_ext, date_ext = extract_cccd_and_date(text_cha)
                    if not cccd_ext:
                        cccd_ext, date_ext = extract_cccd_and_date(text_me)
                    
                    if cccd_ext:
                        df.at[idx, 'SO_CCCD'] = cccd_ext
                        if date_ext and ('NGAYCAP_CCCD' in df.columns):
                            df.at[idx, 'NGAYCAP_CCCD'] = date_ext
                        count_fixed_cccd += 1

            # Đánh lại STT sau khi xóa dòng
            df['STT'] = range(1, len(df) + 1)

            # Thông báo kết quả
            st.success("✅ Đã xử lý hoàn tất!")
            col_m1, col_m2, col_m3 = st.columns(3)
            col_m1.metric("🗑️ Dòng bị xóa (Thiếu Cha&Mẹ)", f"{deleted_rows_count} dòng")
            col_m2.metric("✨ Dòng được bổ sung CCCD", f"{count_fixed_cccd} dòng")
            col_m3.metric("📊 Dòng còn lại", f"{len(df)} dòng")

            # Tạo file kết quả Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Dulieu_DaChuanHoa', index=False)
            output.seek(0)

            # Nút tải file
            st.download_button(
                label="📥 Tải Về File Kết Quả (Excel)",
                data=output,
                file_name="GiayChungNhan_BHXH_DaChuanHoa.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            # Xem trước
            st.subheader("📋 Xem trước 10 dòng dữ liệu đầu tiên:")
            st.dataframe(df.head(10), use_container_width=True)
