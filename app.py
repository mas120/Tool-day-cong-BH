import streamlit as st
import pandas as pd
import io
import re
import random
from datetime import datetime

st.set_page_config(page_title="Tool Chuẩn Hóa BHXH (CT07 & CT03)", page_icon="🏥", layout="wide")

st.title("🏥 Tool Chuẩn Hóa Dữ Liệu BHXH (CT07 & CT03)")
st.write("Chọn mẫu giấy tờ cần xử lý, tải file Excel lên và **chỉnh sửa trực tiếp các dòng cảnh báo** trước khi tải về.")

# Hàm làm sạch CCCD (Giữ nguyên gốc để kiểm tra tính hợp lệ)
def clean_cccd_raw(cccd_str):
    if pd.isna(cccd_str) or not cccd_str:
        return ""
    val = str(cccd_str).strip()
    if val.endswith('.0'):
        val = val[:-2]
    return val

# Hàm kiểm tra và format CCCD
def process_cccd(cccd_str):
    raw_val = clean_cccd_raw(cccd_str)
    if not raw_val:
        return "", False  # Rống, Không hợp lệ
    
    # Nếu chứa chữ cái (như AI2849873) -> Không hợp lệ
    if re.search(r'\D', raw_val):
        return raw_val, False
        
    # Nếu chỉ toàn số
    if len(raw_val) == 12:
        return raw_val, True
    elif 0 < len(raw_val) < 12:
        # Pad số 0 nhưng vẫn đánh dấu là cần kiểm tra nếu bản gốc quá ngắn/bị sai
        return raw_val.zfill(12), False
    else:
        return raw_val, False

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

def extract_cccd_and_date(text):
    if pd.isna(text) or not isinstance(text, str):
        return None, None
    cccd_match = re.search(r'\b\d{12}\b', text)
    cccd = cccd_match.group(0) if cccd_match else None
    date_match = re.search(r'\b(\d{1,2}[/-]\d{1,2}[/-]\d{4})\b', text)
    date_val = format_to_yyyymmdd(date_match.group(0)) if date_match else None
    return cccd, date_val

option_mau = st.radio("📌 Chọn loại mẫu chứng từ cần xử lý:", ["Mẫu CT07 (Giấy nghỉ việc hưởng BHXH)", "Mẫu CT03 (Giấy ra viện)"], horizontal=True)

file_excel = st.file_uploader("📂 Kéo thả hoặc chọn file Excel cần xử lý", type=["xlsx"])

if file_excel:
    if st.button("🚀 Tiến Hành Chuẩn Hóa Dữ Liệu", type="primary"):
        for key in ['df_clean', 'df_original_warn', 'warn_indices', 'deleted_rows_log', 'edit_history']:
            if key in st.session_state:
                del st.session_state[key]

        with st.spinner("Đang đọc và xử lý dữ liệu..."):
            df = pd.read_excel(file_excel, sheet_name=0, dtype=str)
            
            for col in df.columns:
                df[col] = df[col].fillna('').astype(str).str.strip()

            total_before = len(df)
            rows_to_keep = []
            warn_indices = []
            deleted_rows_log = []
            current_year = datetime.now().year

            # ==========================================
            # XỬ LÝ MẪU CT03
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
                    cccd_val, is_valid_cccd = process_cccd(cccd_raw)
                    df.at[idx, 'SO_CCCD'] = cccd_val

                    if cccd_raw and not is_valid_cccd:
                        is_warning = True

                    new_lg = '1' if cccd_val else '0'
                    if 'LOAI_GIAYTO' in df.columns:
                        df.at[idx, 'LOAI_GIAYTO'] = new_lg

                    if 'SO_SERI' in df.columns and df.at[idx, 'SO_SERI']:
                        df.at[idx, 'SO_SERI'] = re.sub(r'2600', '260', df.at[idx, 'SO_SERI'])

                    if 'NGAYCAP_CCCD' in df.columns and df.at[idx, 'NGAYCAP_CCCD']:
                        df.at[idx, 'NGAYCAP_CCCD'] = format_to_yyyymmdd(df.at[idx, 'NGAYCAP_CCCD'])

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
            # XỬ LÝ MẪU CT07
            # ==========================================
            else:
                df['MAU_SO'] = 'CT07'
                df['LOAI_GIAYTO'] = '1'

                for idx in df.index:
                    is_warning = False

                    if 'NGAYCAP_CCCD' in df.columns and df.at[idx, 'NGAYCAP_CCCD']:
                        df.at[idx, 'NGAYCAP_CCCD'] = format_to_yyyymmdd(df.at[idx, 'NGAYCAP_CCCD'])

                    cccd_raw = df.at[idx, 'SO_CCCD'] if 'SO_CCCD' in df.columns else ''
                    cccd_val, is_valid_cccd = process_cccd(cccd_raw)
                    
                    if cccd_val:
                        df.at[idx, 'SO_CCCD'] = cccd_val
                        has_cccd = True
                    else:
                        has_cccd = False

                    if not has_cccd or not is_valid_cccd:
                        text_cha = df.at[idx, 'HO_TEN_CHA'] if 'HO_TEN_CHA' in df.columns else ''
                        text_me = df.at[idx, 'HO_TEN_ME'] if 'HO_TEN_ME' in df.columns else ''
                        
                        cccd_ext, date_ext = extract_cccd_and_date(text_cha)
                        if not cccd_ext:
                            cccd_ext, date_ext = extract_cccd_and_date(text_me)
                        
                        if cccd_ext:
                            df.at[idx, 'SO_CCCD'] = cccd_ext
                            is_valid_cccd = True
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

                    # Báo cảnh báo nếu chứa chữ cái hoặc không đúng chuẩn 12 số
                    if not is_valid_cccd or len(curr_cccd) != 12 or re.search(r'\D', curr_cccd):
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

            # Lưu lại trạng thái ban đầu của các dòng cảnh báo để so sánh chỉnh sửa
            df_original_warn = df_clean.loc[warn_indices].copy() if warn_indices else pd.DataFrame()

            st.session_state['df_clean'] = df_clean
            st.session_state['df_original_warn'] = df_original_warn
            st.session_state['warn_indices'] = warn_indices
            st.session_state['deleted_rows_log'] = deleted_rows_log
            st.session_state['total_before'] = total_before
            st.session_state['option_mau'] = option_mau

# KHU VỰC CHỈNH SỬA VÀ HIỂN THỊ
if 'df_clean' in st.session_state:
    df_clean = st.session_state['df_clean']
    df_original_warn = st.session_state['df_original_warn']
    warn_indices = st.session_state['warn_indices']
    deleted_rows_log = st.session_state['deleted_rows_log']
    total_before = st.session_state['total_before']
    option_mau = st.session_state['option_mau']

    st.success("✅ Đã xử lý chuẩn hóa dữ liệu thành công!")

    c1, c2, c3 = st.columns(3)
    c1.metric("📊 Dòng ban đầu", f"{total_before} dòng")
    c2.metric("🗑️ Dòng bị xóa", f"{len(deleted_rows_log)} dòng")
    c3.metric("⚠️ Dòng bị cảnh báo cần sửa", f"{len(warn_indices)} dòng")

    # Tạo 3 Tabs giao diện
    tab_edit, tab_del, tab_history = st.tabs([
        "✏️ 1. Sửa Trực Tiếp Dòng Cảnh Báo", 
        "🗑️ 2. Danh Sách Dòng Bị Xóa", 
        "📝 3. Nhật Ký Các Dòng Đã Chỉnh Sửa"
    ])

    with tab_edit:
        valid_warn_indices = [i for i in warn_indices if i in df_clean.index]
        
        if valid_warn_indices:
            st.warning(f"⚠️ Phát hiện **{len(valid_warn_indices)} dòng bị cảnh báo** (CCCD chứa ký tự/không đủ 12 số, hoặc trẻ <7t thiếu Bố/Mẹ). Nhấp đúp vào ô để sửa trực tiếp!")
            
            df_warn = df_clean.loc[valid_warn_indices].copy()
            edited_warn = st.data_editor(df_warn, use_container_width=True, key="warn_editor")
            
            # Cập nhật thay đổi vào df_clean
            df_clean.update(edited_warn)
            st.session_state['df_clean'] = df_clean
        else:
            st.success("🎉 Tất cả dữ liệu đều đầy đủ thông tin, không có dòng nào bị cảnh báo!")

    with tab_del:
        if deleted_rows_log:
            st.dataframe(pd.DataFrame(deleted_rows_log), use_container_width=True)
        else:
            st.info("Không có dòng nào bị xóa!")

    with tab_history:
        # Tự động ghi nhận và so sánh sự thay đổi giữa dữ liệu gốc & dữ liệu đã sửa
        edit_logs = []
        if not df_original_warn.empty:
            for idx in df_original_warn.index:
                row_orig = df_original_warn.loc[idx]
                row_curr = df_clean.loc[idx]
                
                changes = []
                for col in df_clean.columns:
                    val_orig = str(row_orig[col])
                    val_curr = str(row_curr[col])
                    if val_orig != val_curr:
                        changes.append(f"Cột **{col}**: '{val_orig}' ➔ '{val_curr}'")
                
                if changes:
                    edit_logs.append({
                        'STT': row_curr.get('STT', str(idx + 1)),
                        'Họ và Tên': row_curr.get('HO_TEN', ''),
                        'Mã BHXH': row_curr.get('MA_SOBHXH', ''),
                        'Nội dung chỉnh sửa': "; ".join(changes)
                    })

        if edit_logs:
            st.subheader("📋 Chi tiết các ô vừa được thao tác chỉnh sửa:")
            st.dataframe(pd.DataFrame(edit_logs), use_container_width=True)
        else:
            st.info("Chưa có thao tác chỉnh sửa nào trên bảng dòng cảnh báo.")

    # Tạo tên file & Tải về
    date_suffix = get_clean_date_str(df_clean)
    if date_suffix:
        output_filename = f"GiayRaVien_BHXH_{date_suffix}.xlsx" if "CT03" in option_mau else f"GiayChungNhan_BHXH_{date_suffix}.xlsx"
    else:
        output_filename = "GiayRaVien_BHXH_DaChuanHoa.xlsx" if "CT03" in option_mau else "GiayChungNhan_BHXH_DaChuanHoa.xlsx"

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
