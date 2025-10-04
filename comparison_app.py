

import streamlit as st
import pandas as pd
import os
import glob
import zipfile
import io

# Cố gắng import pypdf và hướng dẫn cài đặt nếu thiếu
try:
    import pypdf
except ImportError:
    st.error("Thư viện pypdf chưa được cài đặt. Vui lòng chạy lệnh sau trong terminal: pip install pypdf")
    st.stop()

st.set_page_config(page_title="Đối chiếu FPT", layout="wide", page_icon="📊")

st.title("📊 Đối chiếu dữ liệu Grab ")
st.write("Tải lên các tệp của bạn để bắt đầu đối chiếu và xử lý.")

# --- GIAO DIỆN NHẬP LIỆU ---
with st.container(border=True):
    col1, col2, col3 = st.columns([2, 2, 3])
    with col1:
        st.subheader("1. File Transport CSV")
        uploaded_transport_file = st.file_uploader("Tải lên file CSV transport", type=["csv"], label_visibility="collapsed")
    with col2:
        st.subheader("2. File Hóa đơn Excel")
        uploaded_invoice_file = st.file_uploader("Tải lên file Excel hóa đơn", type=["xls", "xlsx"], label_visibility="collapsed")
    with col3:
        st.subheader("3. Folder Báo cáo (.zip)")
        uploaded_zip_file = st.file_uploader("Tải lên file .zip của folder báo cáo", type=["zip"], label_visibility="collapsed")

# --- BẮT ĐẦU XỬ LÝ KHI CÓ ĐỦ FILE ---
if uploaded_transport_file is not None and uploaded_invoice_file is not None:
    try:
        # --- 1. ĐỌC VÀ LÀM SẠCH DỮ LIỆU GỐC ---
        df_transport = pd.read_csv(uploaded_transport_file, skiprows=7)
        try:
            df_invoice = pd.read_html(uploaded_invoice_file)[0]
        except Exception:
            df_invoice = pd.read_excel(uploaded_invoice_file, engine='xlrd')

        df_transport.columns = df_transport.columns.str.strip()
        df_invoice.columns = df_invoice.columns.str.strip()

        if df_invoice.shape[1] < 13:
            st.error(f"File Hóa đơn không có đủ 13 cột. Chỉ tìm thấy {df_invoice.shape[1]} cột.")
            st.stop()
        df_invoice.rename(columns={df_invoice.columns[1]: 'pdf_link_key', df_invoice.columns[12]: 'summary_ma_nhan_hoa_don'}, inplace=True)

        # --- 2. HỢP NHẤT DỮ LIỆU CSV VÀ EXCEL ---
        matching_ids = list(set(df_transport['Booking ID'].dropna()) & set(df_invoice['Booking'].dropna()))
        if not matching_ids:
            st.warning("Không tìm thấy Booking ID nào trùng khớp giữa file CSV và Excel.")
            st.stop()

        df_merged = pd.merge(df_transport[df_transport['Booking ID'].isin(matching_ids)], df_invoice[df_invoice['Booking'].isin(matching_ids)], left_on='Booking ID', right_on='Booking', suffixes=('_transport', '_invoice'))

        # --- 3. XỬ LÝ FOLDER PDF (NẾU CÓ) ---
        if uploaded_zip_file is not None:
            pdf_data = []
            with zipfile.ZipFile(uploaded_zip_file, 'r') as zip_ref:
                pdf_file_names = [name for name in zip_ref.namelist() if name.lower().endswith('.pdf') and not name.startswith('__MACOSX')]
                st.info(f"Bắt đầu xử lý {len(pdf_file_names)} file PDF từ tệp .zip...")
                progress_bar = st.progress(0, text="Đang xử lý file PDF...")
                for i, filename in enumerate(pdf_file_names):
                    try:
                        key_from_filename = os.path.basename(filename).split('_')[2]
                        with zip_ref.open(filename) as pdf_file:
                            pdf_content = pdf_file.read()
                            pdf_reader = pypdf.PdfReader(io.BytesIO(pdf_content))
                            text = "".join([page.extract_text() or "" for page in pdf_reader.pages])
                            
                            found_code = "Không tìm thấy trong PDF"
                            if "Mã nhận hóa đơn" in text:
                                parts = text.split("Mã nhận hóa đơn")
                                if len(parts) > 1:
                                    code = parts[1].split('\n')[0].replace(":", "").strip()
                                    if code:
                                        found_code = code
                            pdf_data.append({'pdf_link_key_str': key_from_filename, 'Mã hóa đơn từ PDF': found_code, 'pdf_content': pdf_content, 'pdf_filename': os.path.basename(filename)})
                    except Exception as e:
                        st.warning(f"Lỗi khi đọc file {filename} trong zip: {e}")
                    progress_bar.progress((i + 1) / len(pdf_file_names), text=f"Đang xử lý: {os.path.basename(filename)}")
            
            if pdf_data:
                df_pdf_data = pd.DataFrame(pdf_data)
                df_merged['pdf_link_key_str'] = df_merged['pdf_link_key'].astype(str)
                df_merged = pd.merge(df_merged, df_pdf_data, on='pdf_link_key_str', how='left')

        # --- 4. THỐNG KÊ VÀ HIỂN THỊ ---
        st.header("📈 Kết quả đối chiếu")
        with st.container(border=True):
            st.subheader("Bảng thống kê tổng hợp")
            agg_dict = {'Số chuyến': ('Booking ID', 'count'), 'Tổng tiền (VND)': ('Total Fare', 'sum')}
            if 'summary_ma_nhan_hoa_don' in df_merged.columns: agg_dict['Mã nhận hóa đơn (tóm tắt)'] = ('summary_ma_nhan_hoa_don', 'first')
            if 'Mã hóa đơn từ PDF' in df_merged.columns: agg_dict['Mã hóa đơn từ PDF'] = ('Mã hóa đơn từ PDF', lambda x: ", ".join(x.dropna().unique()))

            employee_stats = df_merged.groupby('Employee Name').agg(**agg_dict).reset_index()
            st.dataframe(employee_stats)

        # --- 5. KHU VỰC TẢI VỀ ---
        if 'pdf_content' in df_merged.columns:
            with st.container(border=True):
                st.subheader("📥 Tải về báo cáo PDF theo nhân viên")
                employee_list = sorted(df_merged['Employee Name'].unique())
                selected_employee = st.selectbox("Chọn nhân viên để tải về", employee_list)

                if selected_employee:
                    employee_df = df_merged[(df_merged['Employee Name'] == selected_employee) & (df_merged['pdf_content'].notna())]
                    if employee_df.empty:
                        st.warning(f"Không tìm thấy file PDF nào cho nhân viên {selected_employee}.")
                    else:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            for index, row in employee_df.iterrows():
                                zip_file.writestr(row['pdf_filename'], row['pdf_content'])
                        
                        st.download_button(
                            label=f"Tải xuống {len(employee_df)} file PDF cho {selected_employee}",
                            data=zip_buffer.getvalue(),
                            file_name=f"{selected_employee}_reports.zip",
                            mime="application/zip",
                            use_container_width=True
                        )

        # --- 6. CHI TIẾT CHUYẾN ĐI ---
        with st.container(border=True):
            st.subheader("📄 Chi tiết các chuyến đi (cho nhân viên có >1 chuyến)")
            multi_trip_employees = employee_stats[employee_stats['Số chuyến'] > 1]
            if multi_trip_employees.empty:
                st.info("Không có nhân viên nào có nhiều hơn một chuyến đi.")
            else:
                for index, row in multi_trip_employees.iterrows():
                    with st.expander(f"Xem chi tiết cho: {row['Employee Name']} ({row['Số chuyến']} chuyến)"):
                        st.dataframe(df_merged[df_merged['Employee Name'] == row['Employee Name']])

    except Exception as e:
        st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")
        st.exception(e) # In ra chi tiết lỗi để debug

