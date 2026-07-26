import streamlit as st
import pandas as pd
import io
import re
import random
from datetime import datetime

# Cấu hình giao diện Streamlit
st.set_page_config(page_title="Tool Chuẩn Hóa BHXH (CT07 & CT03)", page_icon="🏥", layout="wide")

st.title("🏥 Tool Chuẩn Hóa Dữ Liệu BHXH (CT07 & CT03)")
st.write("Chọn mẫu giấy tờ cần xử lý, tải file Excel lên và **chỉnh sửa hoặc xóa các dòng cảnh báo** trước khi tải về.")

# Hàm làm sạch dữ liệu CCCD ban đầu
def clean_cccd_raw(cccd_str):
    if pd.isna(cccd_str) or not cccd_str:
        return ""
    val = str(cccd_str).strip()
    if val.endswith('.0'):
        val = val[:-2]
    return val

# Hàm kiểm tra và xử lý định dạng CCCD / CMND
def process_cccd(cccd_str, allow_cmnd_9_digits=False):
    raw_val = clean_cccd_raw(cccd_str)
    if not raw_val:
        return "", False, "Trống số CCCD/CMND"
    
    # Nếu chứa ký tự chữ cái (Ví dụ: AI2849873)
    if re.search(r'\D', raw_val):
        return raw_val, False, f"CCCD/CMND chứa ký tự chữ không hợp lệ ({raw_val})"
        
    # Trường hợp CCCD 12 số chuẩn
    if len(raw_val) == 12:
        return raw_val, True, ""
    
    # Trường hợp CMND 9 số cũ (Dành riêng cho CT03)
    elif len(raw_val) == 9 and allow_cmnd_9_digits:
        return raw_val, True, ""
        
    elif 0 < len(raw_val) < 12:
        if allow_cmnd_9_digits:
            return raw_val.zfill(12), False, f"Độ dài số ID là {len(raw_val)} chữ số (Yêu cầu 9 số CMND hoặc 12 số CCCD)"
        else:
            return raw_val.zfill(12), False, f"Độ dài CCCD là {len(raw_val)} chữ số (Yêu cầu đủ 12 số CCCD)"
    else:
        return raw_val, False, f"Độ dài CCCD không hợp lệ ({len(raw_val)} số)"

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

# Lấy chuỗi ngày chứng từ tạo tên file kết quả
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

# Trích xuất năm sinh và ngày sinh dạng YYYYMMDD
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

# Trích xuất CCCD và Ngày cấp từ văn bản Bố/Mẹ
def extract_cccd_and_date(text):
    if pd.isna(text) or not isinstance(text, str):
        return None, None
    cccd_match = re.search(r'\b\d{12}\b', text)
    cccd = cccd_match.group(0) if cccd_match else None
    date_match = re.search(r'\b(\d{1,2}[/-]\d{1,2}[/-]\d{4})\b', text)
    date_val = format_to_yyyymmdd(date_match.group(0)) if date_match else None
    return cccd, date_val

# Lựa chọn loại mẫu chứng từ
option_mau = st.radio("📌 Chọn loại mẫu chứng từ cần xử lý:", ["Mẫu CT07 (Giấy nghỉ việc hưởng BHXH)", "Mẫu CT03 (Giấy ra viện)"], horizontal=True)

file_excel = st.file_uploader("📂 Kéo thả hoặc chọn file Excel cần xử lý", type=["xlsx"])

if file_excel:
    if st.button("🚀 Tiến Hành Chuẩn Hóa Dữ Liệu", type="primary"):
        # Reset Session State
        for key in ['df_clean', 'warn_reasons', 'deleted_rows_log', 'auto_clean_logs']:
            if key in st.session_state:
                del st.session_state[key]

        with st.spinner("Đang đọc và xử lý dữ liệu..."):
            df = pd.read_excel(file_excel, sheet_name=0, dtype=str)
            
            for col in df.columns:
                df[col] = df[col].fillna('').astype(str).str.strip()

            total_before = len(df)
            rows_to_keep = []
            warn_reasons = {}  # Lưu trữ dạng: {index: "Lý do cảnh báo"}
            deleted_rows_log = []
            auto_clean_logs = []
            current_year = datetime.now().year

            def log_auto_fix(idx_val, col_name, old_val, new_val, reason):
                stt_str = str(df.at[idx_val, 'STT']) if 'STT' in df.columns else str(idx_val + 1)
                name_str = str(df.at[idx_val, 'HO_TEN']) if 'HO_TEN' in df.columns else ''
                bhxh_str = str(df.at[idx_val, 'MA_SOBHXH']) if 'MA_SOBHXH' in df.columns else ''
                auto_clean_logs.append({
                    'STT Gốc': stt_str,
                    'Họ và Tên': name_str,
                    'Mã BHXH': bhxh_str,
                    'Cột Xử Lý': col_name,
                    'Dữ Liệu Ban Đầu': old_val if old_val else '(Trống)',
                    'Dữ Liệu Tool Đã Sửa': new_val if new_val else '(Trống)',
                    'Hành Động / Lý Do Chuẩn Hóa': reason
                })

            # ==========================================
            # XỬ LÝ MẪU CT03 (GIẤY RA VIỆN)
            # ==========================================
            if "CT03" in option_mau:
                for idx in df.index:
                    reasons_list = []
                    bhxh_val = str(df.at[idx, 'MA_SOBHXH']) if 'MA_SOBHXH' in df.columns else ''
                    the_val = str(df.at[idx, 'MA_THE']) if 'MA_THE' in df.columns else ''
                    
                    if (not bhxh_val or bhxh_val.lower() == 'nan') and (not the_val or the_val.lower() == 'nan'):
                        deleted_rows_log.append({
                            'STT Gốc': str(df.at[idx, 'STT']) if 'STT' in df.columns else str(idx + 1),
                            'Họ và Tên': str(df.at[idx, 'HO_TEN']) if 'HO_TEN' in df.columns else '',
                            'Mã BHXH': bhxh_val,
                            'Mã Thẻ': the_val,
                            'Lý do xóa': 'Trống cả Mã số BHXH lẫn Mã thẻ BHYT'
                        })
                        continue
                    
                    if 'TEKT' in df.columns and str(df.at[idx, 'TEKT']) != '0':
                        old_tekt = str(df.at[idx, 'TEKT'])
                        df.at[idx, 'TEKT'] = '0'
                        log_auto_fix(idx, 'TEKT', old_tekt, '0', 'Mặc định gán TEKT = 0')

                    cccd_raw = str(df.at[idx, 'SO_CCCD']) if 'SO_CCCD' in df.columns else ''
                    cccd_val, is_valid_cccd, err_msg = process_cccd(cccd_raw, allow_cmnd_9_digits=True)
                    
                    if cccd_raw != cccd_val and not (len(cccd_raw) == 9 and cccd_raw.isdigit()):
                        log_auto_fix(idx, 'SO_CCCD', cccd_raw, cccd_val, 'Chuẩn hóa định dạng CCCD (Tự động thêm 0 cho đủ 12 số)')
                    df.at[idx, 'SO_CCCD'] = cccd_val

                    if cccd_raw and not is_valid_cccd:
                        reasons_list.append(err_msg)

                    new_lg = '1' if cccd_val else '0'
                    if 'LOAI_GIAYTO' in df.columns and str(df.at[idx, 'LOAI_GIAYTO']) != new_lg:
                        old_lg = str(df.at[idx, 'LOAI_GIAYTO'])
                        df.at[idx, 'LOAI_GIAYTO'] = new_lg
                        log_auto_fix(idx, 'LOAI_GIAYTO', old_lg, new_lg, f"Gán loại giấy tờ ({'1: Có CCCD/CMND' if new_lg=='1' else '0: Không CCCD/CMND'})")

                    if 'SO_SERI' in df.columns and df.at[idx, 'SO_SERI']:
                        old_seri = str(df.at[idx, 'SO_SERI'])
                        new_seri = re.sub(r'2600', '260', old_seri)
                        if old_seri != new_seri:
                            df.at[idx, 'SO_SERI'] = new_seri
                            log_auto_fix(idx, 'SO_SERI', old_seri, new_seri, 'Chuyển mã seri từ 2600 thành 260')

                    if 'NGAYCAP_CCCD' in df.columns and df.at[idx, 'NGAYCAP_CCCD']:
                        old_ngaycap = str(df.at[idx, 'NGAYCAP_CCCD'])
                        new_ngaycap = format_to_yyyymmdd(old_ngaycap)
                        if old_ngaycap != new_ngaycap:
                            df.at[idx, 'NGAYCAP_CCCD'] = new_ngaycap
                            log_auto_fix(idx, 'NGAYCAP_CCCD', old_ngaycap, new_ngaycap, 'Chuyển định dạng ngày sang YYYYMMDD')

                    birth_str = str(df.at[idx, 'NGAY_SINH']) if 'NGAY_SINH' in df.columns else ''
                    birth_year, _ = parse_birth_date(birth_str)
                    if birth_year:
                        age = current_year - birth_year
                        if age < 7:
                            text_cha = str(df.at[idx, 'HO_TEN_CHA']) if 'HO_TEN_CHA' in df.columns else ''
                            text_me = str(df.at[idx, 'HO_TEN_ME']) if 'HO_TEN_ME' in df.columns else ''
                            if (not text_cha or text_cha.lower() == 'nan') and (not text_me or text_me.lower() == 'nan'):
                                reasons_list.append("Trẻ < 7 tuổi bị trống thông tin cả Họ tên Cha lẫn Mẹ")

                    rows_to_keep.append(idx)
                    if reasons_list:
                        warn_reasons[idx] = " | ".join(reasons_list)

            # ==========================================
            # XỬ LÝ MẪU CT07 (GIẤY NGHỈ VIỆC HƯỞNG BHXH)
            # ==========================================
            else:
                df['MAU_SO'] = 'CT07'
                df['LOAI_GIAYTO'] = '1'

                for idx in df.index:
                    reasons_list = []

                    if 'NGAYCAP_CCCD' in df.columns and df.at[idx, 'NGAYCAP_CCCD']:
                        old_ngaycap = str(df.at[idx, 'NGAYCAP_CCCD'])
                        new_ngaycap = format_to_yyyymmdd(old_ngaycap)
                        if old_ngaycap != new_ngaycap:
                            df.at[idx, 'NGAYCAP_CCCD'] = new_ngaycap
                            log_auto_fix(idx, 'NGAYCAP_CCCD', old_ngaycap, new_ngaycap, 'Chuyển định dạng ngày sang YYYYMMDD')

                    cccd_raw = str(df.at[idx, 'SO_CCCD']) if 'SO_CCCD' in df.columns else ''
                    cccd_val, is_valid_cccd, err_msg = process_cccd(cccd_raw, allow_cmnd_9_digits=False)
                    
                    if cccd_val:
                        if cccd_raw != cccd_val:
                            log_auto_fix(idx, 'SO_CCCD', cccd_raw, cccd_val, 'Chuẩn hóa định dạng CCCD (Tự động thêm 0 cho đủ 12 số)')
                        df.at[idx, 'SO_CCCD'] = cccd_val
                        has_cccd = True
                    else:
                        has_cccd = False

                    if not has_cccd or not is_valid_cccd:
                        text_cha = str(df.at[idx, 'HO_TEN_CHA']) if 'HO_TEN_CHA' in df.columns else ''
                        text_me = str(df.at[idx, 'HO_TEN_ME']) if 'HO_TEN_ME' in df.columns else ''
                        
                        cccd_ext, date_ext = extract_cccd_and_date(text_cha)
                        source_ext = 'HO_TEN_CHA'
                        if not cccd_ext:
                            cccd_ext, date_ext = extract_cccd_and_date(text_me)
                            source_ext = 'HO_TEN_ME'
                        
                        if cccd_ext:
                            df.at[idx, 'SO_CCCD'] = cccd_ext
                            log_auto_fix(idx, 'SO_CCCD', cccd_raw, cccd_ext, f'Trích xuất tự động số CCCD từ cột {source_ext}')
                            is_valid_cccd = True
                            if date_ext and ('NGAYCAP_CCCD' in df.columns):
                                old_nc = str(df.at[idx, 'NGAYCAP_CCCD'])
                                df.at[idx, 'NGAYCAP_CCCD'] = date_ext
                                log_auto_fix(idx, 'NGAYCAP_CCCD', old_nc, date_ext, f'Trích xuất tự động Ngày cấp từ cột {source_ext}')
                            has_cccd = True

                    curr_cccd = str(df.at[idx, 'SO_CCCD']).strip()

                    if not curr_cccd or curr_cccd.lower() in ['', 'nan']:
                        deleted_rows_log.append({
                            'STT Gốc': str(df.at[idx, 'STT']) if 'STT' in df.columns else str(idx + 1),
                            'Họ và Tên': str(df.at[idx, 'HO_TEN']) if 'HO_TEN' in df.columns else '',
                            'Mã BHXH': str(df.at[idx, 'MA_SOBHXH']) if 'MA_SOBHXH' in df.columns else '',
                            'Lý do xóa': 'Trống số CCCD (và không trích xuất được từ Bố/Mẹ)'
                        })
                        continue

                    if not is_valid_cccd or len(curr_cccd) != 12 or re.search(r'\D', curr_cccd):
                        reasons_list.append(err_msg if err_msg else "CCCD không hợp lệ")

                    elif 'NGAYCAP_CCCD' in df.columns:
                        ngaycap_val = str(df.at[idx, 'NGAYCAP_CCCD']).strip()
                        if not ngaycap_val or ngaycap_val.lower() in ['', 'nan']:
                            birth_str = str(df.at[idx, 'NGAY_SINH']) if 'NGAY_SINH' in df.columns else ''
                            birth_year, birth_formatted = parse_birth_date(birth_str)
                            if birth_year:
                                age = current_year - birth_year
                                if age > 16:
                                    rand_day = random.randint(1, 28)
                                    auto_ngaycap = f"202202{rand_day:02d}"
                                    df.at[idx, 'NGAYCAP_CCCD'] = auto_ngaycap
                                    log_auto_fix(idx, 'NGAYCAP_CCCD', ngaycap_val, auto_ngaycap, 'Bệnh nhân > 16t trống ngày cấp -> Tự động sinh ngày cấp ngẫu nhiên trong 02/2022')
                                else:
                                    df.at[idx, 'NGAYCAP_CCCD'] = birth_formatted
                                    log_auto_fix(idx, 'NGAYCAP_CCCD', ngaycap_val, birth_formatted, 'Bệnh nhân <= 16t trống ngày cấp -> Mặc định gán Ngày cấp = Ngày sinh')

                    rows_to_keep.append(idx)
                    if reasons_list:
                        warn_reasons[idx] = " | ".join(reasons_list)

            df_clean = df.loc[rows_to_keep].copy()
            df_clean['STT'] = [str(i) for i in range(1, len(df_clean) + 1)]

            st.session_state['df_clean'] = df_clean
            st.session_state['warn_reasons'] = warn_reasons
            st.session_state['deleted_rows_log'] = deleted_rows_log
            st.session_state['auto_clean_logs'] = auto_clean_logs
            st.session_state['total_before'] = total_before
            st.session_state['option_mau'] = option_mau

# KHU VỰC CHỈNH SỬA VÀ HIỂN THỊ KẾT QUẢ
if 'df_clean' in st.session_state:
    df_clean = st.session_state['df_clean']
    warn_reasons = st.session_state.get('warn_reasons', {})
    deleted_rows_log = st.session_state['deleted_rows_log']
    auto_clean_logs = st.session_state['auto_clean_logs']
    total_before = st.session_state['total_before']
    option_mau = st.session_state['option_mau']

    st.success("✅ Đã xử lý chuẩn hóa dữ liệu thành công!")

    # Lọc danh sách index cảnh báo còn trong df_clean
    valid_warn_indices = [i for i in warn_reasons.keys() if i in df_clean.index]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📊 Dòng ban đầu", f"{total_before} dòng")
    c2.metric("🗑️ Dòng bị xóa", f"{len(deleted_rows_log)} dòng")
    c3.metric("⚡ Ô dữ liệu Tool tự sửa", f"{len(auto_clean_logs)} mục")
    c4.metric("⚠️ Dòng cảnh báo cần sửa thủ công", f"{len(valid_warn_indices)} dòng")

    tab_edit, tab_del, tab_auto_history = st.tabs([
        "✏️ 1. Sửa / Xóa Dòng Cảnh Báo Thủ Công", 
        "🗑️ 2. Danh Sách Dòng Bị Xóa", 
        "🤖 3. Lịch Sử Tool Tự Động Chuẩn Hóa"
    ])

    # --- TAB 1: SỬA / XÓA THỦ CÔNG ---
    with tab_edit:
        if valid_warn_indices:
            st.warning(f"⚠️ Phát hiện **{len(valid_warn_indices)} dòng bị cảnh báo**. Vui lòng kiểm tra cột **'LÝ_DO_CẢNH_BÁO'** trong bảng bên dưới để sửa hoặc xóa!")
            
            # CHỨC NĂNG XÓA DÒNG CẢNH BÁO
            col_del_select, col_del_btn = st.columns([3, 1])
            
            options_dict = {}
            for w_idx in valid_warn_indices:
                if w_idx in df_clean.index:
                    row_data = df_clean.loc[w_idx]
                    stt_val = str(row_data.get('STT', w_idx + 1))
                    name_val = str(row_data.get('HO_TEN', ''))
                    bhxh_val = str(row_data.get('MA_SOBHXH', ''))
                    reason_val = warn_reasons.get(w_idx, '')
                    label = f"STT {stt_val} - {name_val} ({reason_val})"
                    options_dict[label] = w_idx

            selected_labels = col_del_select.multiselect(
                "🗑️ Chọn các dòng cảnh báo muốn XÓA:",
                options=list(options_dict.keys()),
                placeholder="Chọn một hoặc nhiều dòng để xóa...",
                key="select_rows_to_delete"
            )

            if col_del_btn.button("🔥 Xóa Dòng Đã Chọn", type="secondary"):
                if selected_labels:
                    indices_to_remove = [options_dict[lbl] for lbl in selected_labels if lbl in options_dict]
                    
                    for idx_rem in indices_to_remove:
                        if idx_rem in df_clean.index:
                            row_rem = df_clean.loc[idx_rem]
                            deleted_rows_log.append({
                                'STT Gốc': str(row_rem.get('STT', idx_rem + 1)),
                                'Họ và Tên': str(row_rem.get('HO_TEN', '')),
                                'Mã BHXH': str(row_rem.get('MA_SOBHXH', '')),
                                'Mã Thẻ': str(row_rem.get('MA_THE', '')),
                                'Lý do xóa': 'Xóa thủ công từ danh sách cảnh báo'
                            })
                            # Xóa lý do cảnh báo tương ứng
                            if idx_rem in warn_reasons:
                                del warn_reasons[idx_rem]
                    
                    df_clean.drop(index=indices_to_remove, inplace=True)
                    df_clean['STT'] = [str(i) for i in range(1, len(df_clean) + 1)]
                    
                    st.session_state['df_clean'] = df_clean
                    st.session_state['deleted_rows_log'] = deleted_rows_log
                    st.session_state['warn_reasons'] = warn_reasons
                    
                    st.toast(f"Đã xóa thành công {len(indices_to_remove)} dòng!", icon="✅")
                    st.rerun()

            st.markdown("---")

            # HIỂN THỊ BẢNG SỬA DỮ LIỆU CÓ THÊM CỘT LÝ DO CẢNH BÁO
            df_warn = df_clean.loc[valid_warn_indices].copy()
            
            # Chèn cột LÝ_DO_CẢNH_BÁO vào vị trí đầu tiên
            df_warn.insert(0, 'LÝ_DO_CẢNH_BÁO', [warn_reasons.get(i, '') for i in valid_warn_indices])
            
            edited_warn = st.data_editor(
                df_warn, 
                use_container_width=True, 
                key="warn_editor",
                disabled=["LÝ_DO_CẢNH_BÁO"] # Khóa không cho sửa cột lý do
            )
            
            # Cập nhật kết quả sửa vào df_clean (Loại bỏ cột LÝ_DO_CẢNH_BÁO trước khi gộp)
            updated_df_warn = edited_warn.drop(columns=['LÝ_DO_CẢNH_BÁO'])
            df_clean.update(updated_df_warn)
            st.session_state['df_clean'] = df_clean
        else:
            st.success("🎉 Tất cả dữ liệu đều đầy đủ thông tin, không có dòng nào bị cảnh báo!")

    # --- TAB 2: DANH SÁCH DÒNG BỊ XÓA ---
    with tab_del:
        if deleted_rows_log:
            st.dataframe(pd.DataFrame(deleted_rows_log), use_container_width=True)
        else:
            st.info("Không có dòng nào bị xóa!")

    # --- TAB 3: LỊCH SỬ TOOL TỰ ĐỘNG CHUẨN HÓA ---
    with tab_auto_history:
        if auto_clean_logs:
            st.subheader(f"📋 Chi tiết {len(auto_clean_logs)} vị trí dữ liệu Tool đã tự động can thiệp làm sạch:")
            st.dataframe(pd.DataFrame(auto_clean_logs), use_container_width=True)
        else:
            st.info("Dữ liệu đầu vào đã rất sạch, Tool không phải tự động can thiệp chỉnh sửa trường nào!")

    # Tạo tên file & Tải về (Tự loại bỏ cột LÝ_DO_CẢNH_BÁO nếu có)
    date_suffix = get_clean_date_str(df_clean)
    if date_suffix:
        output_filename = f"GiayRaVien_BHXH_{date_suffix}.xlsx" if "CT03" in option_mau else f"GiayChungNhan_BHXH_{date_suffix}.xlsx"
    else:
        output_filename = "GiayRaVien_BHXH_DaChuanHoa.xlsx" if "CT03" in option_mau else "GiayChungNhan_BHXH_DaChuanHoa.xlsx"

    # Đảm bảo file Excel xuất ra sạch sẽ không chứa cột cảnh báo
    df_export = df_clean.copy()
    if 'LÝ_DO_CẢNH_BÁO' in df_export.columns:
        df_export.drop(columns=['LÝ_DO_CẢNH_BÁO'], inplace=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, sheet_name='Dulieu_DaChuanHoa', index=False)
    output.seek(0)

    st.markdown("---")
    st.download_button(
        label=f"📥 Tải Về File Kết Quả Đã Chỉnh Sửa ({output_filename})",
        data=output,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
