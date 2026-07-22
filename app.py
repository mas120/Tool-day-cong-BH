import streamlit as st
import pandas as pd
import io
import re
import random

# Cấu hình giao diện
st.set_page_config(page_title="Tool Chuẩn Hóa BHXH (CT07 & CT03)", page_icon="🏥", layout="wide")

st.title("🏥 Tool Chuẩn Hóa Dữ Liệu BHXH (CT07 & CT03)")
st.write("Chọn mẫu giấy tờ cần xử lý, tải file Excel lên và **chỉnh sửa trực tiếp các dòng cảnh báo** trước khi tải về.")

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
        # Reset lại session state khi bấm nút chạy file mới
        for key in ['df_clean', 'warn_indices', 'deleted_rows_log']:
            if key in st.session_state:
                del st.session_state[key]

        with st.spinner("Đang đọc và xử lý dữ liệu..."):
            df = pd.read_excel(file_excel, sheet_name=0, dtype=str)
            
            for col in df.columns:
                df[col] = df[col].fillna('').astype(str).str.strip()

            total_before = len(df)
            rows_to_keep = []
            warn_indices = []  # Chỉ lưu duy nhất vị trí các dòng BỊ CẢNH BÁO
            deleted_rows_log = []
            current_year = 2026

            # ==========================================
            # XỬ LÝ MẪU CT03 (GIẤY RA VIỆN)
            # ==========================================
            if "CT03" in option_mau:
                for idx in df.index:
                    is_warning = False
                    bhxh_val = df.at[idx, 'MA_SOBHXH'] if 'MA_SOBHXH' in df.columns else ''
                    the_val = df.at[idx, 'MA_THE'] if 'MA_THE' in df.columns else ''
                    
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

                    cccd_raw = df.at[idx, 'SO_CCCD'] if 'SO_CCCD' in df.columns else ''
                    cccd_val = format_cccd(cccd_raw)
                    df.at[idx, 'SO_CCCD'] = cccd_val

                    new_lg = '1' if cccd_val else '0'
                    if 'LOAI_GIAYTO' in df.columns:
                        df.at[idx, 'LOAI_GIAYTO'] = new_lg

                    if 'SO_SERI' in df.columns and df.at[idx, 'SO_SERI']:
                        df.at[idx, 'SO_SERI'] = re.sub(r'2600', '260', df.at[idx, 'SO_SERI'])

                    if 'NGAYCAP_CCCD' in df.columns and df.at[idx, 'NGAYCAP_CCCD']:
                        df.at[idx, 'NGAYCAP_CCCD'] = format_to_yyyymmdd(df.at[idx, 'NGAYCAP_CCCD'])

                    # 🎯 CHỈ BÁO CẢNH BÁO KHI: Trẻ dưới 7 tuổi bị trống cả Họ tên Cha lẫn Mẹ
                    birth_str = df.at[idx, 'NGAY_SINH'] if 'NGAY_SINH' in df.columns else ''
                    birth_year, _ = parse_birth_date(birth_str)
                    if birth_year:
                        age = current_year - birth_year
                        if age < 7:
                            text_cha = df.at[idx, 'HO_TEN_CHA'] if 'HO_TEN_CHA' in df.columns else ''
                            text_me = df.at[idx, 'HO_TEN_ME'] if 'HO_TEN_ME' in df.columns else ''
                            if (not text_cha or text_cha.lower() == 'nan') and (not text_me or text_me.lower() == 'nan'):
                                is_warning = True

                    rows_to_keep.append(idx)
                    if is_warning:
                        warn_indices.append(idx)

            # ==========================================
            # XỬ LÝ MẪU CT07 (GIẤY NGHỈ VIỆC HƯỞNG BHXH)
            # ==========================================
            else:
                df['MAU_SO'] = 'CT07'
                df['LOAI_GIAYTO'] = '1'

                for idx in df.index:
                    is_warning = False

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

                    curr_cccd = str(df.at[idx, 'SO_CCCD']).strip()

                    if not curr_cccd or curr_cccd.lower() in ['', 'nan']:
                        deleted_rows_log.append({
                            'STT Gốc': df.at[idx, 'STT'] if 'STT' in df.columns else str(idx + 1),
                            'Họ và Tên': df.at[idx, 'HO_TEN'] if 'HO_TEN' in df.columns else '',
                            'Mã BHXH': df.at[idx, 'MA_SOBHXH'] if 'MA_SOBHXH' in df.columns else '',
                            'Lý do xóa': 'Trống số CCCD (và không trích xuất được từ Bố/Mẹ)'
                        })
                        continue

                    # 🎯 CHỈ BÁO CẢNH BÁO KHI: CCCD không đủ 12 số
                    if len(curr_cccd) != 12:
                        is_warning = True

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

                    rows_to_keep.append(idx)
                    if is_warning:
                        warn_indices.append(idx)

            df_clean = df.loc[rows_to_keep].copy()
            df_clean['STT'] = [str(i) for i in range(1, len(df_clean) + 1)]

            # Lưu vào Session State
            st.session_state['df_clean'] = df_clean
            st.session_state['warn_indices'] = warn_indices
            st.session_state['deleted_rows_log'] = deleted_rows_log
            st.session_state['total_before'] = total_before
            st.session_state['option_mau'] = option_mau

# KHU VỰC CHỈNH SỬA
if 'df_clean' in st.session_state:
    df_clean = st.session_state['df_clean']
    warn_indices = st.session_state['warn_indices']
    deleted_rows_log = st.session_state['deleted_rows_log']
    total_before = st.session_state['total_before']
    option_mau = st.session_state['option_mau']

    st.success("✅ Đã xử lý chuẩn hóa dữ liệu thành công!")

    c1, c2, c3 = st.columns(3)
    c1.metric("📊 Dòng ban đầu", f"{total_before} dòng")
    c2.metric("🗑️ Dòng bị xóa", f"{len(deleted_rows_log)} dòng")
    c3.metric("⚠️ Dòng bị cảnh báo cần sửa", f"{len(warn_indices)} dòng")

    tab_edit, tab_del = st.tabs(["✏️ 1. Sửa Trực Tiếp Dòng Cảnh Báo", "🗑️ 2. Danh Sách Dòng Bị Xóa"])

    with tab_edit:
        # Lọc CHÍNH XÁC chỉ các dòng bị CẢNH BÁO
        valid_warn_indices = [i for i in warn_indices if i in df_clean.index]
        
        if valid_warn_indices:
            st.warning(f"⚠️ Phát hiện **{len(valid_warn_indices)} dòng bị cảnh báo** (Trẻ <7t thiếu thông tin Bố/Mẹ hoặc CCCD không đủ 12 số). Bạn nhấp đúp vào ô tương ứng để bổ sung/sửa lại trực tiếp!")
            
            # Chỉ lấy các dòng bị cảnh báo ra bảng sửa
            df_warn = df_clean.loc[valid_warn_indices].copy()
            
            # Bảng hiển thị sửa trực tiếp
            edited_warn = st.data_editor(df_warn, use_container_width=True, key="warn_editor")
            
            # Cập nhật kết quả vừa gõ sửa vào DataFrame chính
            df_clean.update(edited_warn)
        else:
            st.success("🎉 Tất cả dữ liệu đều đầy đủ thông tin, không có dòng nào bị cảnh báo!")

    with tab_del:
        if deleted_rows_log:
            st.dataframe(pd.DataFrame(deleted_rows_log), use_container_width=True)
        else:
            st.info("Không có dòng nào bị xóa!")

    # Tạo tên file
    date_suffix = get_clean_date_str(df_clean)
    if date_suffix:
        output_filename = f"GiayRaVien_BHXH_{date_suffix}.xlsx" if "CT03" in option_mau else f"GiayChungNhan_BHXH_{date_suffix}.xlsx"
    else:
        output_filename = "GiayRaVien_BHXH_DaChuanHoa.xlsx" if "CT03" in option_mau else "GiayChungNhan_BHXH_DaChuanHoa.xlsx"

    # Xuất file Excel đã đồng bộ
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_clean.to_excel(writer, sheet_name='Dulieu_DaChuanHoa', index=False)
    output.seek(0)

    st.markdown("---")
    st.download_button(
        label=f"📥 Tải Về File Kết Quả Đã Chỉnh Sửa ({output_filename})",
        data=output,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
