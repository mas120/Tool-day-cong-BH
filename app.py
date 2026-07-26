import streamlit as st
import pandas as pd
import io
import re
import random
from datetime import datetime

st.set_page_config(page_title="Tool Chuẩn Hóa BHXH (CT07 & CT03)", page_icon="🏥", layout="wide")

st.title("🏥 Tool Chuẩn Hóa Dữ Liệu BHXH (CT07 & CT03)")
st.write("Chọn mẫu giấy tờ cần xử lý, tải file Excel lên và **chỉnh sửa hoặc xóa các dòng cảnh báo** trước khi tải về.")

# Hàm làm sạch CCCD
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
        return "", False
    
    # Nếu chứa ký tự không phải số (Ví dụ AI2849873) -> Cảnh báo
    if re.search(r'\D', raw_val):
        return raw_val, False
        
    if len(raw_val) == 12:
        return raw_val, True
    elif 0 < len(raw_val) < 12:
        # Thêm số 0 vào đầu cho đủ 12 số
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
        for key in ['df_clean', 'warn_indices', 'deleted_rows_log', 'auto_clean_logs']:
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
            auto_clean_logs = []  # Lưu lại tất cả thao tác tool đã tự động xử lý
            current_year = datetime.now().year

            def log_auto_fix(idx, col_name, old_val, new_val, reason):
                stt_str = df.at[idx, 'STT'] if 'STT' in df.columns else str(idx + 1)
                name_str = df.at[idx, 'HO_TEN'] if 'HO_TEN' in df.columns else ''
                bhxh_str = df.at[idx, 'MA_SOBHXH'] if 'MA_SOBHXH' in df.columns else ''
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
                    
                    # Sửa TEKT -> '0'
                    if 'TEKT' in df.columns and df.at[idx, 'TEKT'] != '0':
                        old_tekt = df.at[idx, 'TEKT']
                        df.at[idx, 'TEKT'] = '0'
                        log_auto_fix(idx, 'TEKT', old_tekt, '0', 'Mặc định gán TEKT = 0')

                    # Chuẩn hóa CCCD
                    cccd_raw = df.at[idx, 'SO_CCCD'] if 'SO_CCCD' in df.columns else ''
                    cccd_val, is_valid_cccd = process_cccd(cccd_raw)
                    if cccd_raw != cccd_val:
                        log_auto_fix(idx, 'SO_CCCD', cccd_raw, cccd_val, 'Chuẩn hóa định dạng CCCD (Tự động thêm 0 cho đủ 12 số)')
                    df.at[idx, 'SO_CCCD'] = cccd_val

                    if cccd_raw and not is_valid_cccd:
                        is_warning = True

                    # Cập nhật LOAI_GIAYTO
                    new_lg = '1' if cccd_val else '0'
                    if 'LOAI_GIAYTO' in df.columns and df.at[idx, 'LOAI_GIAYTO'] != new_lg:
                        old_lg = df.at[idx, 'LOAI_GIAYTO']
                        df.at[idx, 'LOAI_GIAYTO'] = new_lg
                        log_auto_fix(idx, 'LOAI_GIAYTO', old_lg, new_lg, f"Gán loại giấy tờ ({'1: Có CCCD' if new_lg=='1' else '0: Không CCCD'})")

                    # Sửa SO_SERI 2600 -> 260
                    if 'SO_SERI' in df.columns and df.at[idx, 'SO_SERI']:
                        old_seri = df.at[idx, 'SO_SERI']
                        new_seri = re.sub(r'2600', '260', old_seri)
                        if old_seri != new_seri:
                            df.at[idx, 'SO_SERI'] = new_seri
                            log_auto_fix(idx, 'SO_SERI', old_seri, new_seri, 'Chuyển mã seri từ 2600 thành 260')

                    # Chuẩn hóa Ngày cấp CCCD
                    if 'NGAYCAP_CCCD' in df.columns and df.at[idx, 'NGAYCAP_CCCD']:
                        old_ngaycap = df.at[idx, 'NGAYCAP_CCCD']
                        new_ngaycap = format_to_yyyymmdd(old_ngaycap)
                        if old_ngaycap != new_ngaycap:
                            df.at[idx, 'NGAYCAP_CCCD'] = new_ngaycap
                            log_auto_fix(idx, 'NGAYCAP_CCCD', old_ngaycap, new_ngaycap, 'Chuyển định dạng ngày sang YYYYMMDD')

                    # Kiểm tra trẻ dưới 7 tuổi
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
                        old_ngaycap = df.at[idx, 'NGAYCAP_CCCD']
                        new_ngaycap = format_to_yyyymmdd(old_ngaycap)
                        if old_ngaycap != new_ngaycap:
                            df.at[idx, 'NGAYCAP_CCCD'] = new_ngaycap
                            log_auto_fix(idx, 'NGAYCAP_CCCD', old_ngaycap, new_ngaycap, 'Chuyển định dạng ngày sang YYYYMMDD')

                    cccd_raw = df.at[idx, 'SO_CCCD'] if 'SO_CCCD' in df.columns else ''
                    cccd_val, is_valid_cccd = process_cccd(cccd_raw)
                    
                    if cccd_val:
                        if cccd_raw != cccd_val:
                            log_auto_fix(idx, 'SO_CCCD', cccd_raw, cccd_val, 'Chuẩn hóa định dạng CCCD (Tự động thêm 0 cho đủ 12 số)')
                        df.at[idx, 'SO_CCCD'] = cccd_val
                        has_cccd = True
                    else:
                        has_cccd = False

                    # Trích xuất CCCD từ Cha / Mẹ nếu trống
                    if not has_cccd or not is_valid_cccd:
                        text_cha = df.at[idx, 'HO_TEN_CHA'] if 'HO_TEN_CHA' in df.columns else ''
                        text_me = df.at[idx, 'HO_TEN_ME'] if 'HO_TEN_ME' in df.columns else ''
                        
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
                                old_nc = df.at[idx, 'NGAYCAP_CCCD']
                                df.at[idx, 'NGAYCAP_CCCD'] = date_ext
                                log_auto_fix(idx, 'NGAYCAP_CCCD', old_nc, date_ext, f'Trích xuất tự động Ngày cấp từ cột {source_ext}')
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

                    # Cảnh báo nếu chứa chữ cái hoặc không đúng chuẩn 12 số
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
                                    auto_ngaycap = f"202202{rand_day:02d}"
                                    df.at[idx, 'NGAYCAP_CCCD'] = auto_ngaycap
                                    log_auto_fix(idx, 'NGAYCAP_CCCD', ngaycap_val, auto_ngaycap, 'Bệnh nhân > 16t trống ngày cấp -> Tự động sinh ngày cấp ngẫu nhiên trong 02/2022')
                                else:
                                    df.at[idx, 'NGAYCAP_CCCD'] = birth_formatted
                                    log_auto_fix(idx, 'NGAYCAP_CCCD', ngaycap_val, birth_formatted, 'Bệnh nhân <= 16t trống ngày cấp -> Mặc định gán Ngày cấp = Ngày sinh')

                    rows_to_keep.append(idx)
                    if is_warning:
                        warn_indices.append(idx)

            df_clean = df.loc[rows_to_keep].copy()
            df_clean['STT'] = [str(i) for i in range(1, len(df_clean) + 1)]

            st.session_state['df_clean'] = df_clean
            st.session_state['warn_indices'] = warn_indices
            st.session_state['deleted_rows_log'] = deleted_rows_log
            st.session_state['auto_clean_logs'] = auto_clean_logs
            st.session_state['total_before'] = total_before
            st.session_state['option_mau'] = option_mau

# KHU VỰC CHỈNH SỬA VÀ HIỂN THỊ
if 'df_clean' in st.session_state:
    df_clean = st.session_state['df_clean']
    warn_indices = st.session_state['warn_indices']
    deleted_rows_log = st.session_state['deleted_rows_log']
    auto_clean_logs = st.session_state['auto_clean_logs']
    total_before = st.session_state['total_before']
    option_mau = st.session_state['option_mau']

    st.success("✅ Đã xử lý chuẩn hóa dữ liệu thành công!")

    # Lọc các index cảnh báo còn trong df_clean
    valid_warn_indices = [i for i in warn_indices if i in df_clean.index]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📊 Dòng ban đầu", f"{total_before} dòng")
    c2.metric("🗑️ Dòng bị xóa", f"{len(deleted_rows_log)} dòng")
    c3.metric("⚡ Trường dữ liệu Tool tự sửa", f"{len(auto_clean_logs)} mục")
    c4.metric("⚠️ Dòng cảnh báo cần sửa thủ công", f"{len(valid_warn_indices)} dòng")

    tab_edit, tab_del, tab_auto_history = st.tabs([
        "✏️ 1. Sửa / Xóa Dòng Cảnh Báo Thủ Công", 
        "🗑️ 2. Danh Sách Dòng Bị Xóa", 
        "🤖 3. Lịch Sử Tool Tự Động Chuẩn Hóa"
    ])

    # --- TAB 1: SỬA / XÓA THỦ CÔNG ---
    with tab_edit:
        if valid_warn_indices:
            st.warning(f"⚠️ Phát hiện **{len(valid_warn_indices)} dòng bị cảnh báo**. Nhấp đúp vào ô để sửa trực tiếp, hoặc chọn các dòng cần loại bỏ bên dưới.")
            
            # CHỨC NĂNG XÓA
            col_del_select, col_del_btn = st.columns([3, 1])
            
            options_dict = {}
            for idx in valid_warn_indices:
                stt_val = df_clean.at[idx, 'STT'] if 'STT' in df_clean.columns else str(idx + 1)
                name_val = df_clean.at[idx, 'HO_TEN'] if 'HO_TEN' in df_clean.columns else ''
                bhxh_val = df_clean.at[idx, 'MA_SOBHXH'] if 'MA_SOBHXH' in df.columns else ''
                label = f"STT {stt_val} - {name_val} (BHXH: {bhxh_val})"
                options_dict[label] = idx

            selected_labels = col_del_select.multiselect(
                "🗑️ Chọn các dòng cảnh báo muốn XÓA:",
                options=list(options_dict.keys()),
                placeholder="Chọn một hoặc nhiều dòng để xóa..."
            )

            if col_del_btn.button("🔥 Xóa Dòng Đã Chọn", type="secondary"):
                if selected_labels:
                    indices_to_remove = [options_dict[lbl] for lbl in selected_labels]
                    
                    for idx_rem in indices_to_remove:
                        deleted_rows_log.append({
                            'STT Gốc': df_clean.at[idx_rem, 'STT'] if 'STT' in df_clean.columns else str(idx_rem + 1),
                            'Họ và Tên': df_clean.at[idx_rem, 'HO_TEN'] if 'HO_TEN' in df_clean.columns else '',
                            'Mã BHXH': df_clean.at[idx_rem, 'MA_SOBHXH'] if 'MA_SOBHXH' in df.columns else '',
                            'Mã Thẻ': df_clean.at[idx_rem, 'MA_THE'] if 'MA_THE' in df.columns else '',
                            'Lý do xóa': 'Xóa thủ công từ danh sách cảnh báo'
                        })
                    
                    df_clean.drop(index=indices_to_remove, inplace=True)
                    df_clean['STT'] = [str(i) for i in range(1, len(df_clean) + 1)]
                    
                    st.session_state['df_clean'] = df_clean
                    st.session_state['deleted_rows_log'] = deleted_rows_log
                    st.session_state['warn_indices'] = [i for i in warn_indices if i not in indices_to_remove]
                    
                    st.toast(f"Đã xóa thành công {len(indices_to_remove)} dòng!", icon="✅")
                    st.rerun()

            st.markdown("---")

            # BẢNG DỮ LIỆU SỬA TRỰC TIẾP
            df_warn = df_clean.loc[valid_warn_indices].copy()
            edited_warn = st.data_editor(df_warn, use_container_width=True, key="warn_editor")
            
            # Cập nhật kết quả gõ vào DataFrame chính
            df_clean.update(edited_warn)
            st.session_state['df_clean'] = df_clean
        else:
            st.success("🎉 Tất cả dữ liệu đều đầy đủ thông tin, không có dòng nào bị cảnh báo!")

    # --- TAB 2: DANH SÁCH XÓA ---
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
