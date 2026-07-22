import streamlit as st
import pandas as pd
import io
import re
import random

# Cấu hình giao diện
st.set_page_config(page_title="Tool Chuẩn Hóa BHXH (CT07 & CT03)", page_icon="🏥", layout="wide")

st.title("🏥 Tool Chuẩn Hóa Dữ Liệu BHXH (CT07 & CT03)")
st.write("Chọn mẫu giấy tờ cần xử lý và tải file Excel lên để hệ thống tự động chuẩn hóa.")

# Hàm làm sạch và kiểm tra CCCD đủ 12 chữ số
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

# Hàm lấy chuỗi Ngày/Tháng/Năm từ cột NGAY_CT để tạo tên file
def get_clean_date_str(df):
    if 'NGAY_CT' in df.columns:
        first_valid_date = df['NGAY_CT'].dropna().astype(str).str.strip()
        first_valid_date = first_valid_date[first_valid_date != ''].first_valid_index()
        if first_valid_date is not None:
            raw_date = df.at[first_valid_date, 'NGAY_CT']
            clean_str = re.sub(r'\D', '', str(raw_date))
            if len(clean_str) >= 8:
                return clean_str[:8]
    return ""

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
            modified_rows_log = []  # Nhật ký các dòng có chỉnh sửa / cảnh báo
            deleted_rows_log = []   # Nhật ký các dòng bị xóa
            current_year = 2026

            # ==========================================
            # XỬ LÝ MẪU CT03 (GIẤY RA VIỆN)
            # ==========================================
            if "CT03" in option_mau:
                for idx in df.index:
                    changes = []
                    bhxh_val = df.at[idx, 'MA_SOBHXH'] if 'MA_SOBHXH' in df.columns else ''
                    the_val = df.at[idx, 'MA_THE'] if 'MA_THE' in df.columns else ''
                    
                    # ❌ XÓA DÒNG: Nếu CẢ MA_SOBHXH VÀ MA_THE đều trống
                    if (not bhxh_val or bhxh_val.lower() == 'nan') and (not the_val or the_val.lower() == 'nan'):
                        deleted_rows_log.append({
                            'STT Gốc': df.at[idx, 'STT'] if 'STT' in df.columns else str(idx + 1),
                            'Họ và Tên': df.at[idx, 'HO_TEN'] if 'HO_TEN' in df.columns else '',
                            'Mã BHXH': bhxh_val,
                            'Mã Thẻ': the_val,
                            'Lý do xóa': 'Trống cả Mã số BHXH lẫn Mã thẻ BHYT'
                        })
                        continue
                    
                    if 'TEKT' in df.columns and df.at[idx, 'TEKT'] != '0':
                        df.at[idx, 'TEKT'] = '0'
                        changes.append("Gán TEKT = 0")

                    cccd_raw = df.at[idx, 'SO_CCCD'] if 'SO_CCCD' in df.columns else ''
                    cccd_val = format_cccd(cccd_raw)
                    if df.at[idx, 'SO_CCCD'] != cccd_val:
                        df.at[idx, 'SO_CCCD'] = cccd_val
                        changes.append(f"Chuẩn hóa CCCD: {cccd_val}")

                    new_lg = '1' if cccd_val else '0'
                    if 'LOAI_GIAYTO' in df.columns and df.at[idx, 'LOAI_GIAYTO'] != new_lg:
                        df.at[idx, 'LOAI_GIAYTO'] = new_lg
                        changes.append(f"Gán LOAI_GIAYTO = {new_lg}")

                    if 'SO_SERI' in df.columns and df.at[idx, 'SO_SERI']:
                        old_seri = df.at[idx, 'SO_SERI']
                        new_seri = re.sub(r'2600', '260', old_seri)
                        if old_seri != new_seri:
                            df.at[idx, 'SO_SERI'] = new_seri
                            changes.append(f"Sửa Seri: {old_seri} -> {new_seri}")

                    if 'NGAYCAP_CCCD' in df.columns and df.at[idx, 'NGAYCAP_CCCD']:
                        old_nc = df.at[idx, 'NGAYCAP_CCCD']
                        new_nc = format_to_yyyymmdd(old_nc)
                        if old_nc != new_nc:
                            df.at[idx, 'NGAYCAP_CCCD'] = new_nc
                            changes.append(f"Sửa Ngày cấp: {old_nc} -> {new_nc}")

                    # ⚠️ CẢNH BÁO MỚI: Bệnh nhân dưới 7 tuổi nhưng trống ô HO_TEN_CHA và HO_TEN_ME
                    birth_str = df.at[idx, 'NGAY_SINH'] if 'NGAY_SINH' in df.columns else ''
                    birth_year, _ = parse_birth_date(birth_str)
                    
                    if birth_year:
                        age = current_year - birth_year
                        if age < 7:
                            text_cha = df.at[idx, 'HO_TEN_CHA'] if 'HO_TEN_CHA' in df.columns else ''
                            text_me = df.at[idx, 'HO_TEN_ME'] if 'HO_TEN_ME' in df.columns else ''
                            if (not text_cha or text_cha.lower() == 'nan') and (not text_me or text_me.lower() == 'nan'):
                                changes.append(f"⚠️ CẢNH BÁO: Bệnh nhân {age} tuổi (<7t) nhưng trống cả Họ tên Cha lẫn Mẹ")

                    rows_to_keep.append(idx)

                    if changes:
                        modified_rows_log.append({
                            'STT': df.at[idx, 'STT'] if 'STT' in df.columns else str(idx + 1),
                            'Họ và Tên': df.at[idx, 'HO_TEN'] if 'HO_TEN' in df.columns else '',
                            'Mã BHXH': df.at[idx, 'MA_SOBHXH'] if 'MA_SOBHXH' in df.columns else '',
                            'Trạng thái / Nội dung': " | ".join(changes)
                        })

            # ==========================================
            # XỬ LÝ MẪU CT07 (GIẤY NGHỈ VIỆC HƯỞNG BHXH)
            # ==========================================
            else:
                df['MAU_SO'] = 'CT07'
                df['LOAI_GIAYTO'] = '1'

                for idx in df.index:
                    changes = []

                    # 1. Chuẩn hóa Ngày cấp nếu có sẵn
                    if 'NGAYCAP_CCCD' in df.columns and df.at[idx, 'NGAYCAP_CCCD']:
                        old_date = df.at[idx, 'NGAYCAP_CCCD']
                        new_date = format_to_yyyymmdd(old_date)
                        if old_date != new_date:
                            df.at[idx, 'NGAYCAP_CCCD'] = new_date
                            changes.append(f"Định dạng Ngày cấp: {old_date} -> {new_date}")

                    # 2. Xử lý CCCD
                    cccd_raw = df.at[idx, 'SO_CCCD'] if 'SO_CCCD' in df.columns else ''
                    cccd_val = format_cccd(cccd_raw)
                    
                    if cccd_val:
                        if df.at[idx, 'SO_CCCD'] != cccd_val:
                            df.at[idx, 'SO_CCCD'] = cccd_val
                            changes.append(f"Bảo toàn số 0 CCCD: {cccd_val}")
                        has_cccd = True
                    else:
                        has_cccd = False

                    # Trích xuất CCCD từ Bố/Mẹ nếu bản thân chưa có
                    if not has_cccd:
                        text_cha = df.at[idx, 'HO_TEN_CHA'] if 'HO_TEN_CHA' in df.columns else ''
                        text_me = df.at[idx, 'HO_TEN_ME'] if 'HO_TEN_ME' in df.columns else ''
                        
                        cccd_ext, date_ext = extract_cccd_and_date(text_cha)
                        if not cccd_ext:
                            cccd_ext, date_ext = extract_cccd_and_date(text_me)
                        
                        if cccd_ext:
                            df.at[idx, 'SO_CCCD'] = format_cccd(cccd_ext)
                            changes.append(f"Trích xuất CCCD từ Bố/Mẹ: {df.at[idx, 'SO_CCCD']}")
                            if date_ext and ('NGAYCAP_CCCD' in df.columns):
                                df.at[idx, 'NGAYCAP_CCCD'] = date_ext
                                changes.append(f"Trích xuất Ngày cấp từ Bố/Mẹ: {date_ext}")
                            has_cccd = True

                    curr_cccd = str(df.at[idx, 'SO_CCCD']).strip()

                    # ❌ XÓA DÒNG: Nếu trống SO_CCCD và không lấy được từ Bố/Mẹ
                    if not curr_cccd or curr_cccd.lower() in ['', 'nan']:
                        deleted_rows_log.append({
                            'STT Gốc': df.at[idx, 'STT'] if 'STT' in df.columns else str(idx + 1),
                            'Họ và Tên': df.at[idx, 'HO_TEN'] if 'HO_TEN' in df.columns else '',
                            'Mã BHXH': df.at[idx, 'MA_SOBHXH'] if 'MA_SOBHXH' in df.columns else '',
                            'Lý do xóa': 'Trống số CCCD (và không trích xuất được từ Bố/Mẹ)'
                        })
                        continue

                    # ⚠️ CẢNH BÁO: Số CCCD không đủ 12 số
                    if len(curr_cccd) != 12:
                        changes.append(f"⚠️ CẢNH BÁO: Số CCCD không đủ 12 số (Hiện có {len(curr_cccd)} số: '{curr_cccd}') - Giữ nguyên không chỉnh sửa")

                    # Bổ sung Ngày cấp nếu bị thiếu
                    elif 'NGAYCAP_CCCD' in df.columns:
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
                                changes.append(f"Tự động điền Ngày cấp (Tuổi {age}): {df.at[idx, 'NGAYCAP_CCCD']}")

                    rows_to_keep.append(idx)

                    if changes:
                        modified_rows_log.append({
                            'STT': df.at[idx, 'STT'] if 'STT' in df.columns else str(idx + 1),
                            'Họ và Tên': df.at[idx, 'HO_TEN'] if 'HO_TEN' in df.columns else '',
                            'Mã BHXH': df.at[idx, 'MA_SOBHXH'] if 'MA_SOBHXH' in df.columns else '',
                            'Số CCCD': df.at[idx, 'SO_CCCD'] if 'SO_CCCD' in df.columns else '',
                            'Trạng thái / Nội dung': " | ".join(changes)
                        })

            # Lọc lại DataFrame và đánh lại STT liên tục
            df = df.loc[rows_to_keep].copy()
            df['STT'] = [str(i) for i in range(1, len(df) + 1)]

            # Lấy chuỗi ngày tháng từ cột NGAY_CT để tạo tên file
            date_suffix = get_clean_date_str(df)
            if date_suffix:
                output_filename = f"GiayRaVien_BHXH_{date_suffix}.xlsx" if "CT03" in option_mau else f"GiayChungNhan_BHXH_{date_suffix}.xlsx"
            else:
                output_filename = "GiayRaVien_BHXH_DaChuanHoa.xlsx" if "CT03" in option_mau else "GiayChungNhan_BHXH_DaChuanHoa.xlsx"

            # Thông báo kết quả
            st.success("✅ Đã xử lý chuẩn hóa dữ liệu thành công!")
            c1, c2, c3 = st.columns(3)
            c1.metric("📊 Dòng ban đầu", f"{total_before} dòng")
            c2.metric("🗑️ Dòng bị xóa", f"{len(deleted_rows_log)} dòng")
            c3.metric("✨ Dòng hợp lệ xuất ra", f"{len(df)} dòng")

            # HIỂN THỊ CÁC TAB BẢNG NHẬT KÝ
            tab1, tab2 = st.tabs(["📝 Dòng có Chỉnh sửa / Cảnh báo", "🗑️ Danh sách Dòng bị Xóa"])
            
            with tab1:
                if modified_rows_log:
                    st.dataframe(pd.DataFrame(modified_rows_log), use_container_width=True)
                else:
                    st.info("Tất cả các dòng xuất ra đều hợp lệ và không có sự thay đổi dữ liệu!")

            with tab2:
                if deleted_rows_log:
                    st.dataframe(pd.DataFrame(deleted_rows_log), use_container_width=True)
                else:
                    st.info("Không có dòng nào bị xóa!")

            # Xuất file Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Dulieu_DaChuanHoa', index=False)
            output.seek(0)

            st.download_button(
                label=f"📥 Tải Về File Kết Quả ({output_filename})",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
