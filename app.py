import streamlit as st
import pandas as pd
import io
import re
import random

# Cấu hình giao diện
st.set_page_config(page_title="Tool Chuẩn Hóa BHXH (CT07 & CT03)", page_icon="🏥", layout="wide")

st.title("🏥 Tool Chuẩn Hóa Dữ Liệu BHXH (CT07 & CT03)")
st.write("Chọn mẫu giấy tờ cần xử lý và tải file Excel lên để hệ thống tự động chuẩn hóa.")

# Hàm làm sạch và bảo toàn số 0 đầu dãy CCCD (đủ 12 chữ số)
def format_cccd(cccd_str):
    if pd.isna(cccd_str) or not cccd_str:
        return ""
    val = str(cccd_str).strip()
    if val.endswith('.0'):
        val = val[:-2]
    val = re.sub(r'\D', '', val)
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
    m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', date_str)
    if m:
        day, month, year = m.groups()
        return f"{year}{int(month):02d}{int(day):02d}"
    m2 = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', date_str)
    if m2:
        year, month, day = m2.groups()
        return f"{year}{int(month):02d}{int(day):02d}"
    return date_str

# Hàm trích xuất Năm sinh và chuỗi YYYYMMDD
def parse_birth_date(date_str):
    if pd.isna(date_str) or not date_str:
        return None, None
    date_str = str(date_str).strip()
    m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', date_str)
    if m:
        day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return year, f"{year}{month:02d}{day:02d}"
    m2 = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', date_str)
    if m2:
        year, month, day = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        return year, f"{year}{month:02d}{day:02d}"
    return None, None

# Hàm trích xuất CCCD và Ngày cấp từ chuỗi Bố/Mẹ
def extract_cccd_and_date(text):
    if pd.isna(text) or not isinstance(text, str):
        return None, None
    cccd_match = re.search(r'\b\d{12}\b', text)
    cccd = cccd_match.group(0) if cccd_match else None
    date_match = re.search(r'\b(\d{1,2}[/-]\d{1,2}[/-]\d{4})\b', text)
    date_val = format_to_yyyymmdd(date_match.group(0)) if date_match else None
    return cccd, date_val

# Lựa chọn loại mẫu giấy tờ
option_mau = st.radio("📌 Chọn loại mẫu chứng từ cần xử lý:", ["Mẫu CT07 (Giấy nghỉ việc hưởng BHXH)", "Mẫu CT03 (Giấy ra viện)"], horizontal=True)

file_excel = st.file_uploader("📂 Kéo thả hoặc chọn file Excel cần xử lý", type=["xlsx"])

if file_excel:
    if st.button("🚀 Tiến Hành Chuẩn Hóa Dữ Liệu", type="primary"):
        with st.spinner("Đang đọc và xử lý dữ liệu..."):
            df = pd.read_excel(file_excel, sheet_name=0, dtype=str)
            
            # Làm sạch khoảng trắng ban đầu
            for col in df.columns:
                df[col] = df[col].fillna('').astype(str).str.strip()

            total_before = len(df)
            rows_to_keep = []

            # ==========================================
            # XỬ LÝ MẪU CT03 (GIẤY RA VIỆN)
            # ==========================================
            if "CT03" in option_mau:
                for idx in df.index:
                    bhxh_val = df.at[idx, 'MA_SOBHXH'] if 'MA_SOBHXH' in df.columns else ''
                    the_val = df.at[idx, 'MA_THE'] if 'MA_THE' in df.columns else ''
                    
                    # Quy tắc 1: Xóa dòng nếu CẢ MA_SOBHXH VÀ MA_THE đều trống
                    if (not bhxh_val or bhxh_val.lower() == 'nan') and (not the_val or the_val.lower() == 'nan'):
                        continue  # Bỏ qua dòng này
                    
                    # Quy tắc 2: TEKT = 0
                    df.at[idx, 'TEKT'] = '0'

                    # Xử lý CCCD & LOAI_GIAYTO
                    cccd_raw = df.at[idx, 'SO_CCCD'] if 'SO_CCCD' in df.columns else ''
                    cccd_val = format_cccd(cccd_raw)
                    df.at[idx, 'SO_CCCD'] = cccd_val

                    # Quy tắc 3: LOAI_GIAYTO = 1 nếu có SO_CCCD, ngược lại = 0
                    if cccd_val:
                        df.at[idx, 'LOAI_GIAYTO'] = '1'
                    else:
                        df.at[idx, 'LOAI_GIAYTO'] = '0'

                    # Quy tắc 4: Sửa SO_SERI (xóa 1 số 0 trong chuỗi 2600 -> 260)
                    if 'SO_SERI' in df.columns and df.at[idx, 'SO_SERI']:
                        seri_val = df.at[idx, 'SO_SERI']
                        # Thay thế cụm 2600 thành 260
                        df.at[idx, 'SO_SERI'] = re.sub(r'2600', '260', seri_val)

                    # Định dạng lại NGAYCAP_CCCD nếu có
                    if 'NGAYCAP_CCCD' in df.columns and df.at[idx, 'NGAYCAP_CCCD']:
                        df.at[idx, 'NGAYCAP_CCCD'] = format_to_yyyymmdd(df.at[idx, 'NGAYCAP_CCCD'])

                    rows_to_keep.append(idx)

            # ==========================================
            # XỬ LÝ MẪU CT07 (GIẤY NGHỈ VIỆC HƯỞNG BHXH)
            # ==========================================
            else:
                df['MAU_SO'] = 'CT07'
                df['LOAI_GIAYTO'] = '1'
                current_year = 2026

                for idx in df.index:
                    if 'NGAYCAP_CCCD' in df.columns and df.at[idx, 'NGAYCAP_CCCD']:
                        df.at[idx, 'NGAYCAP_CCCD'] = format_to_yyyymmdd(df.at[idx, 'NGAYCAP_CCCD'])

                    cccd_raw = df.at[idx, 'SO_CCCD'] if 'SO_CCCD' in df.columns else ''
                    cccd_val = format_cccd(cccd_raw)
                    
                    if cccd_val:
                        df.at[idx, 'SO_CCCD'] = cccd_val
                        has_cccd = True
                    else:
                        has_cccd = False

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
                            has_cccd = True

                    if has_cccd and 'NGAYCAP_CCCD' in df.columns:
                        ngaycap_val = str(df.at[idx, 'NGAYCAP_CCCD']).strip()
                        if not ngaycap_val or ngaycap_val.lower() in ['', 'nan']:
                            birth_str = df.at[idx, 'NGAY_SINH'] if 'NGAY_SINH' in df.columns else ''
                            birth_year, birth_formatted = parse_birth_date(birth_str)
                            if birth_year:
                                age = current_year - birth_year
                                if age > 16:
                                    rand_day = random.randint(1, 28)
                                    df.at[idx, 'NGAYCAP_CCCD'] = f"202202{rand_day:02d}"
                                else:
                                    df.at[idx, 'NGAYCAP_CCCD'] = birth_formatted

                    if has_cccd:
                        rows_to_keep.append(idx)

            # Lọc lại DataFrame và đánh lại STT
            df = df.loc[rows_to_keep].copy()
            deleted_count = total_before - len(df)
            df['STT'] = [str(i) for i in range(1, len(df) + 1)]

            # Thông báo kết quả
            st.success("✅ Đã xử lý chuẩn hóa dữ liệu thành công!")
            c1, c2, c3 = st.columns(3)
            c1.metric("📊 Dòng ban đầu", f"{total_before} dòng")
            c2.metric("🗑️ Dòng bị xóa", f"{deleted_count} dòng")
            c3.metric("✨ Dòng hợp lệ còn lại", f"{len(df)} dòng")

            # Xuất file Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Dulieu_DaChuanHoa', index=False)
            output.seek(0)

            st.download_button(
                label="📥 Tải Về File Kết Quả (Excel)",
                data=output,
                file_name="GiayRaVien_BHXH_DaChuanHoa.xlsx" if "CT03" in option_mau else "GiayChungNhan_BHXH_DaChuanHoa.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            st.subheader("📋 Xem trước 10 dòng dữ liệu đầu tiên:")
            st.dataframe(df.head(10), use_container_width=True)
