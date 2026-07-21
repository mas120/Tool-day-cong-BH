import streamlit as st
import pandas as pd
import io
import re

# Cấu hình giao diện
st.set_page_config(page_title="Tool Chuẩn Hóa BHXH (CT07)", page_icon="🏥", layout="wide")

st.title("🏥 Tool Chuẩn Hóa Giấy Chứng Nhận Hưởng BHXH (CT07)")
st.write("Tải lên file Excel để tự động cập nhật `MAU_SO`, `LOAI_GIAYTO` và trích xuất `SO_CCCD` / `NGAYCAP_CCCD` từ thông tin Bố/Mẹ.")

# Hàm trích xuất 12 số CCCD và Ngày cấp từ chuỗi văn bản
def extract_cccd_and_date(text):
    if pd.isna(text) or not isinstance(text, str):
        return None, None
    
    # Tìm dãy 12 số liên tiếp (CCCD)
    cccd_match = re.search(r'\b\d{12}\b', text)
    cccd = cccd_match.group(0) if cccd_match else None
    
    # Tìm ngày cấp dạng DD/MM/YYYY hoặc DD-MM-YYYY
    date_match = re.search(r'\b(\d{1,2}[/-]\d{1,2}[/-]\d{4})\b', text)
    date_val = date_match.group(0) if date_match else None
    
    return cccd, date_val

# Khu vực tải file
file_excel = st.file_uploader("📂 Kéo thả hoặc chọn file Excel (CT07) cần xử lý", type=["xlsx"])

if file_excel:
    if st.button("🚀 Tiến Hành Chuẩn Hóa Dữ Liệu", type="primary"):
        with st.spinner("Đang đọc và xử lý dữ liệu..."):
            # Đọc dữ liệu từ sheet đầu tiên
            df = pd.read_excel(file_excel, sheet_name=0)

            # Quy tắc 1 & 2: Gán MAU_SO = 'CT07' và LOAI_GIAYTO = 1
            df['MAU_SO'] = 'CT07'
            df['LOAI_GIAYTO'] = 1

            # Quy tắc 3: Điền SO_CCCD & NGAYCAP_CCCD nếu dòng đó bị trống
            count_fixed = 0
            for idx in df.index:
                cccd_val = str(df.at[idx, 'SO_CCCD']).strip() if pd.notna(df.at[idx, 'SO_CCCD']) else ''
                
                # Nếu CCCD trống hoặc bằng 'nan'
                if not cccd_val or cccd_val.lower() == 'nan':
                    text_cha = str(df.at[idx, 'HO_TEN_CHA']) if pd.notna(df.at[idx, 'HO_TEN_CHA']) else ''
                    text_me = str(df.at[idx, 'HO_TEN_ME']) if pd.notna(df.at[idx, 'HO_TEN_ME']) else ''
                    
                    # Thử trích xuất từ Bố trước, nếu không có thì trích xuất từ Mẹ
                    cccd_ext, date_ext = extract_cccd_and_date(text_cha)
                    if not cccd_ext:
                        cccd_ext, date_ext = extract_cccd_and_date(text_me)
                    
                    # Điền kết quả tìm được
                    if cccd_ext:
                        df.at[idx, 'SO_CCCD'] = cccd_ext
                        if date_ext and ('NGAYCAP_CCCD' in df.columns):
                            df.at[idx, 'NGAYCAP_CCCD'] = date_ext
                        count_fixed += 1

            # Thông báo thành công
            st.success(f"✅ Đã xử lý hoàn tất! Bổ sung thành công CCCD cho **{count_fixed}** trường hợp bị thiếu.")

            # Tạo file kết quả lưu trong bộ nhớ tạm
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

            # Bản xem trước dữ liệu
            st.subheader("📋 Xem trước 10 dòng dữ liệu đầu tiên:")
            st.dataframe(df.head(10), use_container_width=True)