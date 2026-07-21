import streamlit as st
import pandas as pd
import io
import re
import random

# Cấu hình giao diện
st.set_page_config(page_title="Tool Chuẩn Hóa BHXH (CT07)", page_icon="🏥", layout="wide")

st.title("🏥 Tool Chuẩn Hóa Giấy Chứng Nhận Hưởng BHXH (CT07)")
st.write("Tải lên file Excel để tự động chuẩn hóa `MAU_SO`, `LOAI_GIAYTO`, `SO_CCCD` (bảo toàn số 0 ở đầu), `NGAYCAP_CCCD` và lọc bản ghi.")

# Hàm làm sạch và bảo toàn số 0 đầu dãy CCCD (đủ 12 chữ số)
def format_cccd(cccd_str):
    if pd.isna(cccd_str) or not cccd_str:
        return ""
    
    val = str(cccd_str).strip()
    
    # Loại bỏ đuôi .0 nếu bị dính do kiểu float (vd: 96077005028.0)
    if val.endswith('.0'):
        val = val[:-2]
        
    val = re.sub(r'\D', '', val) # Chỉ giữ lại các chữ số
    
    # Nếu là dãy số từ 1 đến 12 chữ số, bù đủ số 0 ở đầu cho đủ 12 chữ số
    if 0 < len(val) <= 12:
        return val.zfill(12)
        
    return val

# Hàm chuyển đổi định dạng ngày sang YYYYMMDD
def format_to_yyyymmdd(date_str):
    if pd.isna(date_str) or not date_str:
        return ""
    
    date_str = str(date_str).strip()
    
    if date_str.endswith('.0'):
        date_str = date_str[:-2]
        
    if len(date_str) == 8 and date_str.isdigit():
        return date_str
    
    # Dạng DD/MM/YYYY hoặc DD-MM-YYYY
    m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', date_str)
    if m:
        day, month, year = m.groups()
        return f"{year}{int(month):02d}{int(day):02d}"
    
    # Dạng YYYY-MM-DD
    m2 = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', date_str)
    if m2:
        year, month, day = m2.groups()
        return f"{year}{int(month):02d}{int(day):02d}"
        
    return date_str

# Hàm trích xuất Năm sinh và chuỗi YYYYMMDD từ ngày sinh
def parse_birth_date(date_str):
    if pd.isna(date_str) or not date_str:
        return None, None
    date_str = str(date_str).strip()
    
    # Dạng DD/MM/YYYY
    m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', date_str)
    if m:
        day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return year, f"{year}{month:02d}{day:02d}"
        
    # Dạng YYYY-MM-DD
    m2 = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', date_str)
    if m2:
        year, month, day = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        return year, f"{year}{month:02d}{day:02d}"
        
    return None, None

# Hàm trích xuất 12 số CCCD và Ngày cấp từ chuỗi văn bản Bố/Mẹ
def extract_cccd_and_date(text):
    if pd.isna(text) or not isinstance(text, str):
        return None, None
    
    cccd_match = re.search(r'\b\d{12}\b', text)
    cccd = cccd_match.group(0) if cccd_match else None
    
    date_match = re.search(r'\b(\d{1,2}[/-]\d{1,2}[/-]\d{4})\b', text)
    date_val = format_to_yyyymmdd(date_match.group(0)) if date_match else None
    
    return cccd, date_val

# Khu vực tải file
file_excel = st.file_uploader("📂 Kéo thả hoặc chọn file Excel (CT07) cần xử lý", type=["xlsx"])

if file_excel:
    if st.button("🚀 Tiến Hành Chuẩn Hóa Dữ Liệu", type="primary"):
        with st.spinner("Đang đọc và xử lý dữ liệu..."):
            # Đọc toàn bộ dữ liệu dưới dạng Chuỗi (dtype=str) để giữ nguyên số 0 ở đầu các cột số
            df = pd.read_excel(file_excel, sheet_name=0, dtype=str)

            # Ép kiểu dữ liệu dạng chuỗi cho các cột cần làm việc
            check_cols = ['SO_CCCD', 'NGAYCAP_CCCD', 'HO_TEN_CHA', 'HO_TEN_ME', 'NGAY_SINH']
            for col in check_cols:
                if col in df.columns:
                    df[col] = df[col].fillna('').astype(str).str.strip()

            # Gán giá trị mặc định cho MAU_SO và LOAI_GIAYTO
            df['MAU_SO'] = 'CT07'
            df['LOAI_GIAYTO'] = '1'

            rows_to_keep = []
            count_fixed_cccd = 0
            count_fixed_ngaycap = 0

            current_year = 2026  # Năm hiện tại

            for idx in df.index:
                # 1. Định dạng lại NGAYCAP_CCCD hiện có nếu có sẵn
                if 'NGAYCAP_CCCD' in df.columns and df.at[idx, 'NGAYCAP_CCCD']:
                    df.at[idx, 'NGAYCAP_CCCD'] = format_to_yyyymmdd(df.at[idx, 'NGAYCAP_CCCD'])

                # 2. Xử lý & bảo toàn số 0 ở đầu cho SO_CCCD hiện có
                cccd_raw = df.at[idx, 'SO_CCCD'] if 'SO_CCCD' in df.columns else ''
                cccd_val = format_cccd(cccd_raw)
                
                if cccd_val:
                    df.at[idx, 'SO_CCCD'] = cccd_val
                    has_cccd = True
                else:
                    has_cccd = False

                # Nếu chưa có SO_CCCD, trích xuất từ Bố hoặc Mẹ
                if not has_cccd:
                    text_cha = df.at[idx, 'HO_TEN_CHA'] if 'HO_TEN_CHA' in df.columns else ''
                    text_me = df.at[idx, 'HO_TEN_ME'] if 'HO_TEN_ME' in df.columns else ''
                    
                    cccd_ext, date_ext = extract_cccd_and_date(text_cha)
                    if not cccd_ext:
                        cccd_ext, date_ext = extract_cccd_and_date(text_me)
                    
                    if cccd_ext:
                        df.at[idx, 'SO_CCCD'] = format_cccd(cccd_ext)
                        if date_ext and ('NGAYCAP_CCCD' in df.columns):
                            df.at[idx, 'NGAYCAP_CCCD'] = date_ext
                        count_fixed_cccd += 1
                        has_cccd = True

                # 3. Bổ sung NGAYCAP_CCCD nếu đã có SO_CCCD nhưng vẫn thiếu NGAYCAP_CCCD
                if has_cccd and 'NGAYCAP_CCCD' in df.columns:
                    ngaycap_val = str(df.at[idx, 'NGAYCAP_CCCD']).strip()
                    
                    if not ngaycap_val or ngaycap_val.lower() in ['', 'nan']:
                        birth_str = df.at[idx, 'NGAY_SINH'] if 'NGAY_SINH' in df.columns else ''
                        birth_year, birth_formatted = parse_birth_date(birth_str)
                        
                        if birth_year:
                            age = current_year - birth_year
                            if age > 16:
                                # > 16 tuổi: Năm 2022 tháng 02, ngày ngẫu nhiên (01 - 28)
                                rand_day = random.randint(1, 28)
                                df.at[idx, 'NGAYCAP_CCCD'] = f"202202{rand_day:02d}"
                            else:
                                # <= 16 tuổi: Lấy theo ngày sinh YYYYMMDD
                                df.at[idx, 'NGAYCAP_CCCD'] = birth_formatted
                            count_fixed_ngaycap += 1

                # Chỉ giữ lại dòng hợp lệ (có SO_CCCD)
                if has_cccd:
                    rows_to_keep.append(idx)

            # Lọc lại DataFrame
            total_before = len(df)
            df = df.loc[rows_to_keep].copy()
            deleted_rows_count = total_before - len(df)

            # Đánh lại STT liên tục
            df['STT'] = [str(i) for i in range(1, len(df) + 1)]

            # Thông báo kết quả
            st.success("✅ Đã xử lý hoàn tất! Đã bảo toàn đầy đủ số 0 ở đầu dãy CCCD.")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("✨ Bổ sung CCCD từ Bố/Mẹ", f"{count_fixed_cccd} dòng")
            c2.metric("📅 Bổ sung Ngày Cấp theo Tuổi", f"{count_fixed_ngaycap} dòng")
            c3.metric("🗑️ Dòng bị xóa (Không CCCD)", f"{deleted_rows_count} dòng")
            c4.metric("📊 Dòng hợp lệ", f"{len(df)} dòng")

            # Xuất file Excel
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
