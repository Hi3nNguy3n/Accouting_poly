# Copyright by hiennm22, LocTH5 - BM UDPM
import streamlit as st
import pandas as pd
import os
import glob
import zipfile
import io
import re

# Cố gắng import pypdf và hướng dẫn cài đặt nếu thiếu
try:
    import pypdf
except ImportError:
    st.error("Thư viện pypdf chưa được cài đặt. Vui lòng chạy lệnh sau trong terminal: pip install pypdf")
    st.stop()

# --- HÀM HỖ TRỢ ---
def load_unit_mapping():
    """Đọc file Excel mapping và trả về một dictionary từ Employee Name sang Đơn vị."""
    try:
        df_mapping = pd.read_excel("FileMau/Tong hop _ Report.xlsx")
        # Cột B là Employee Name, Cột E là Đơn vị (theo yêu cầu của người dùng)
        name_col = df_mapping.columns[1]
        unit_col_map = df_mapping.columns[4]
        df_mapping = df_mapping.dropna(subset=[name_col])
        # Tạo map, xử lý trường hợp tên nhân viên trùng lặp (lấy đơn vị đầu tiên)
        return df_mapping.set_index(name_col)[unit_col_map].to_dict()
    except FileNotFoundError:
        st.error("Lỗi: Không tìm thấy file mapping 'FileMau/Tong hop _ Report.xlsx'. Vui lòng đảm bảo file tồn tại.")
        st.stop()
    except IndexError:
        st.error("Lỗi: File mapping 'Tong hop _ Report.xlsx' không có đủ 5 cột (để lấy cột B và E).")
        st.stop()
    except Exception as e:
        st.error(f"Lỗi khi đọc file mapping: {e}")
        st.stop()

st.set_page_config(page_title="Đối chiếu FPT", layout="wide", page_icon="📊")

st.title("📊 Đối chiếu dữ liệu Grab & Báo cáo PDF")
st.write("Tải lên các tệp của bạn để bắt đầu đối chiếu và xử lý.")
st.caption("Copyright by LocTH5, Hiennm22 - BM UDPM")

# --- GIAO DIỆN NHẬP LIỆU ---
with st.container(border=True):
    col1, col2, col3 = st.columns([2, 2, 3])
    file_types = ["csv", "xls", "xlsx"]
    with col1:
        st.subheader("1. File Transport")
        uploaded_transport_file = st.file_uploader("Tải lên file Transport", type=file_types, label_visibility="collapsed")
    with col2:
        st.subheader("2. File Hóa đơn")
        uploaded_invoice_file = st.file_uploader("Tải lên file Hóa đơn", type=file_types, label_visibility="collapsed")
    with col3:
        st.subheader("3. Folder Báo cáo (.zip)")
        uploaded_zip_file = st.file_uploader("Tải lên file .zip của folder báo cáo", type=["zip"], label_visibility="collapsed")

# --- BẮT ĐẦU XỬ LÝ KHI CÓ ĐỦ FILE ---
if uploaded_transport_file is not None and uploaded_invoice_file is not None:
    try:
        employee_to_unit_map = load_unit_mapping()

        # --- 1. ĐỌC VÀ LÀM SẠCH DỮ LIỆU GỐC ---
        # Đọc file transport (CSV hoặc Excel)
        if uploaded_transport_file.name.endswith('.csv'):
            df_transport = pd.read_csv(uploaded_transport_file, skiprows=8)
        else:
            df_transport = pd.read_excel(uploaded_transport_file, skiprows=8)

        # Đọc file hóa đơn (CSV hoặc Excel)
        if uploaded_invoice_file.name.endswith('.csv'):
            df_invoice = pd.read_csv(uploaded_invoice_file)
        elif uploaded_invoice_file.name.endswith('.xls'):
            try: # Xử lý trường hợp file .xls thực chất là HTML
                df_invoice = pd.read_html(uploaded_invoice_file)[0]
            except Exception:
                df_invoice = pd.read_excel(uploaded_invoice_file, engine='xlrd')
        else: # .xlsx
            df_invoice = pd.read_excel(uploaded_invoice_file)

        df_transport.columns = df_transport.columns.str.strip()
        df_invoice.columns = df_invoice.columns.str.strip()

        # Lấy tên cột địa chỉ theo chỉ số và đổi tên để tránh xung đột khi merge
        if len(df_transport.columns) > 9: # Phải có ít nhất 10 cột
            pickup_col_name = df_transport.columns[7]
            dropoff_col_name = df_transport.columns[9]
            df_transport.rename(columns={
                pickup_col_name: 'GEMINI_PICKUP_ADDRESS',
                dropoff_col_name: 'GEMINI_DROPOFF_ADDRESS'
            }, inplace=True)

        if df_invoice.shape[1] < 13:
            st.error(f"File Hóa đơn không có đủ 13 cột. Chỉ tìm thấy {df_invoice.shape[1]} cột.")
            st.stop()
        rename_dict = {
            df_invoice.columns[1]: 'pdf_link_key',
            df_invoice.columns[12]: 'summary_ma_nhan_hoa_don'
        }
        if len(df_invoice.columns) > 4:
            rename_dict[df_invoice.columns[4]] = 'GEMINI_NGAY_HD_INVOICE' # Cột 5 là Ngày HĐ
            rename_dict[df_invoice.columns[5]] = 'HINH_THUC_TT' # Cột 5 là Ngày HĐ
            rename_dict[df_invoice.columns[6]] = 'TIEN_TRC_THUE' # Cột 5 là Ngày HĐ
            rename_dict[df_invoice.columns[7]] = 'TIEN_THUE8' # Cột 5 là Ngày HĐ
            rename_dict[df_invoice.columns[8]] = 'TONG_TIEN' # Cột 5 là Ngày HĐ
            rename_dict[df_invoice.columns[15]] = 'NGAY_BOOKING' # Cột 5 là Ngày HĐ
            rename_dict[df_invoice.columns[16]] = 'SO_HOA_DON' # Cột 5 là Ngày HĐ
        df_invoice.rename(columns=rename_dict, inplace=True)

        # --- 2. HỢP NHẤT DỮ LIỆU CSV VÀ EXCEL ---
        matching_ids = list(set(df_transport['Booking ID'].dropna()) & set(df_invoice['Booking'].dropna()))
        if not matching_ids:
            st.warning("Không tìm thấy Booking ID nào trùng khớp giữa hai file đầu vào.")
            st.stop()

        df_merged = pd.merge(df_transport[df_transport['Booking ID'].isin(matching_ids)], df_invoice[df_invoice['Booking'].isin(matching_ids)], left_on='Booking ID', right_on='Booking', suffixes=('_transport', '_invoice'))

        # Áp dụng mapping để thêm cột 'Đơn vị'
        df_merged['Đơn vị'] = df_merged['Employee Name'].map(employee_to_unit_map)
        df_merged['Đơn vị'].fillna('Không xác định', inplace=True)

        # --- 3. XỬ LÝ FOLDER PDF (NẾU CÓ) ---
        count_no_pdf = 0
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
                            
                            # Trích xuất ngày hóa đơn từ PDF
                            ngay_hd_str = "Không tìm thấy"
                            match = re.search(r'Ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})', text, re.IGNORECASE)
                            if match:
                                day, month, year = match.groups()
                                ngay_hd_str = f"{day.zfill(2)}/{month.zfill(2)}/{year}"

                            pdf_data.append({
                                'pdf_link_key_str': key_from_filename, 
                                'Mã hóa đơn từ PDF': found_code, 
                                'Ngay_HD_pdf': ngay_hd_str,
                                'pdf_content': pdf_content, 
                                'pdf_filename': os.path.basename(filename)
                            })
                    except Exception as e:
                        st.warning(f"Lỗi khi đọc file {filename} trong zip: {e}")
                    progress_bar.progress((i + 1) / len(pdf_file_names), text=f"Đang xử lý: {os.path.basename(filename)}")
            
            if pdf_data:
                df_pdf_data = pd.DataFrame(pdf_data)
                df_merged['pdf_link_key_str'] = df_merged['pdf_link_key'].astype(str)
                df_merged = pd.merge(df_merged, df_pdf_data, on='pdf_link_key_str', how='left')
                count_no_pdf = df_merged['pdf_filename'].isnull().sum()

        # --- 4. THỐNG KÊ VÀ HIỂN THỊ ---
        if count_no_pdf > 0:
            st.warning(f"### 🔥 Chú ý: Có {count_no_pdf} hóa đơn không có file PDF tương ứng.")
        st.header("📈 Kết quả đối chiếu")
        with st.container(border=True):
            st.subheader("Bảng thống kê tổng hợp")
            agg_dict = {'Số chuyến': ('Booking ID', 'count'), 'Tổng tiền (VND)': ('Total Fare', 'sum')}
            if 'summary_ma_nhan_hoa_don' in df_merged.columns: agg_dict['Mã nhận hóa đơn (tóm tắt)'] = ('summary_ma_nhan_hoa_don', 'first')
            if 'Mã hóa đơn từ PDF' in df_merged.columns: agg_dict['Mã hóa đơn từ PDF'] = ('Mã hóa đơn từ PDF', lambda x: ", ".join(x.dropna().unique()))

            employee_stats = df_merged.groupby('Employee Name').agg(**agg_dict).reset_index()
            st.dataframe(employee_stats)

            # --- Chức năng tạo và tải Bảng kê Excel ---
            try:
                # --- Tìm tất cả các cột cần thiết một cách linh hoạt ---
                def find_col(df, possibilities):
                    for p in possibilities:
                        if p in df.columns:
                            return p
                    return None

                date_cols = ['Date', 'Date of Trip', 'Trip Date', 'Ngày', 'Date & Time (GMT+7)']
                date_col_name = find_col(df_merged, date_cols)
                # Sau khi đổi tên, chúng ta có thể tìm trực tiếp và chắc chắn hơn
                pickup_col_name = 'GEMINI_PICKUP_ADDRESS' if 'GEMINI_PICKUP_ADDRESS' in df_merged.columns else None
                dropoff_col_name = 'GEMINI_DROPOFF_ADDRESS' if 'GEMINI_DROPOFF_ADDRESS' in df_merged.columns else None

                if date_col_name is None:
                    # LỖI NGHIÊM TRỌNG: Cột ngày tháng là bắt buộc. Hiển thị lỗi và dừng lại.
                    try:
                        uploaded_transport_file.seek(0)
                        if uploaded_transport_file.name.endswith('.csv'):
                            header_list = pd.read_csv(uploaded_transport_file, skiprows=8, nrows=1).columns.tolist()
                        else:
                            header_list = pd.read_excel(uploaded_transport_file, skiprows=8, nrows=1).columns.tolist()
                        st.error(f"**Lỗi Bảng kê: Không thể tìm thấy cột Ngày tháng.** Vui lòng cho tôi biết tên cột ngày tháng chính xác từ danh sách bên dưới. **Các cột tìm được:** `{header_list}`")
                    except Exception as e:
                        st.error(f"Lỗi Bảng kê: Không tìm thấy cột ngày tháng. Lỗi khi đọc cột: {e}")
                else:
                    # Cột ngày tháng đã được tìm thấy, tiếp tục xử lý.
                    # Hiển thị cảnh báo nếu không tìm thấy các cột không bắt buộc.
                    if not pickup_col_name:
                        st.warning("Cảnh báo: Không tìm thấy cột 'Địa chỉ đón' (Pick-up Address). Cột này sẽ bị bỏ trống trong file Excel.")
                    if not dropoff_col_name:
                        st.warning("Cảnh báo: Không tìm thấy cột 'Địa chỉ trả' (Drop-off Address). Cột này sẽ bị bỏ trống trong file Excel.")

                    df_merged['Date_dt'] = pd.to_datetime(df_merged[date_col_name], errors='coerce')
                    start_date = df_merged['Date_dt'].min()
                    end_date = df_merged['Date_dt'].max()

                    @st.cache_data
                    def generate_bang_ke_excel(_df, s_date, e_date, _d_col, _p_col, _do_col):
                        from openpyxl import load_workbook
                        from openpyxl.styles import Font
                        template_path = "FileMau/BangKe.xlsx"
                        wb = load_workbook(template_path)
                        ws = wb.active
                        ws['C4'] = s_date.strftime('%d/%m/%Y') if pd.notna(s_date) else "N/A"
                        ws['C5'] = e_date.strftime('%d/%m/%Y') if pd.notna(e_date) else "N/A"
                        start_row = 8
                        for i, row in _df.reset_index(drop=True).iterrows():
                            ws.cell(row=start_row + i, column=1, value=i + 1)
                            ws.cell(row=start_row + i, column=2, value=row['Booking ID'])
                            ws.cell(row=start_row + i, column=3, value=f" {row['Employee Name']}")
                            if _p_col:
                                ws.cell(row=start_row + i, column=4, value=row[_p_col])
                            if _do_col:
                                ws.cell(row=start_row + i, column=5, value=row[_do_col])
                            # ws.cell(row=start_row + i, column=6, value=row['Date_dt'].strftime('%d/%m/%Y') if pd.notna(row['Date_dt']) else row[_d_col])
                            # ws.cell(row=start_row + i, column=7, value=row['Total Fare'])
                            if 'GEMINI_NGAY_HD_INVOICE' in row and pd.notna(row['GEMINI_NGAY_HD_INVOICE']):
                                ws.cell(row=start_row + i, column=6, value=row['GEMINI_NGAY_HD_INVOICE'])
                            ws.cell(row=start_row + i, column=7, value=row['HINH_THUC_TT'])
                            ws.cell(row=start_row + i, column=8, value="{:,.0f}".format(row['TIEN_TRC_THUE']))
                            ws.cell(row=start_row + i, column=9, value="{:,.0f}".format(row['TIEN_THUE8']))
                            ws.cell(row=start_row + i, column=10, value="{:,.0f}".format(row['TONG_TIEN']))
                            ws.cell(row=start_row + i, column=11, value=row['NGAY_BOOKING'])
                            ws.cell(row=start_row + i, column=12, value=row['pdf_link_key'])
                        
                        total_row_index = start_row + len(_df)
                        total_label_cell = ws.cell(row=total_row_index, column=7, value="Tổng cộng")
                        total_label_cell.font = Font(bold=True)

                        total_tien_trc_thue = _df['TIEN_TRC_THUE'].sum()
                        total_tien_thue8 = _df['TIEN_THUE8'].sum()
                        total_tong_tien = _df['TONG_TIEN'].sum()

                        total_value_cell_8 = ws.cell(row=total_row_index, column=8, value="{:,.0f}".format(total_tien_trc_thue))
                        total_value_cell_8.font = Font(bold=True)
                        total_value_cell_9 = ws.cell(row=total_row_index, column=9, value="{:,.0f}".format(total_tien_thue8))
                        total_value_cell_9.font = Font(bold=True)
                        total_value_cell_10 = ws.cell(row=total_row_index, column=10, value="{:,.0f}".format(total_tong_tien))
                        total_value_cell_10.font = Font(bold=True)
                        
                        excel_buffer = io.BytesIO()
                        wb.save(excel_buffer)
                        return excel_buffer.getvalue()

                    st.divider()
                    st.subheader("Tải Bảng Kê")

                    # --- Tùy chọn 1: Tải file ZIP theo Đơn vị ---
                    st.markdown("##### 📥 Tải Bảng kê theo Đơn vị (.zip)")
                    
                    unit_col = 'Đơn vị'
                    if unit_col not in df_merged.columns:
                        st.error(f"Lỗi: Không tìm thấy cột '{unit_col}'. Quá trình mapping có thể đã thất bại.")
                    else:
                        # --- Part 1: Download for a single unit with PDFs ---
                        st.markdown("###### Tải cho 1 đơn vị (kèm hóa đơn PDF)")
                        if 'pdf_content' not in df_merged.columns:
                            st.warning("Chưa có dữ liệu PDF. Vui lòng tải lên file .zip chứa các hóa đơn để sử dụng chức năng này.")
                        else:
                            unique_units_for_select = sorted(df_merged[unit_col].dropna().unique())
                            selected_unit = st.selectbox("Chọn đơn vị", unique_units_for_select)

                            if selected_unit:
                                df_unit = df_merged[df_merged[unit_col] == selected_unit]
                                
                                @st.cache_data
                                def create_zip_for_unit_with_pdfs(_unit_df, _unit_name):
                                    zip_buffer = io.BytesIO()
                                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                        start_date_unit = _unit_df['Date_dt'].min()
                                        end_date_unit = _unit_df['Date_dt'].max()
                                        excel_data = generate_bang_ke_excel(_unit_df, start_date_unit, end_date_unit, date_col_name, pickup_col_name, dropoff_col_name)
                                        safe_unit_name = "".join(c for c in str(_unit_name) if c.isalnum() or c in (' ', '_')).rstrip()
                                        excel_filename = f"BangKe_{safe_unit_name}_{start_date_unit.strftime('%Y%m%d') if pd.notna(start_date_unit) else 'nodate'}_{end_date_unit.strftime('%Y%m%d') if pd.notna(end_date_unit) else 'nodate'}.xlsx"
                                        zip_file.writestr(excel_filename, excel_data)

                                        df_pdfs = _unit_df[_unit_df['pdf_content'].notna()]
                                        for _, row in df_pdfs.iterrows():
                                            zip_file.writestr(f"HOA_DON_PDF/{row['pdf_filename']}", row['pdf_content'])
                                    return zip_buffer.getvalue()

                                zip_data_single = create_zip_for_unit_with_pdfs(df_unit, selected_unit)
                                
                                safe_unit_name_zip = "".join(c for c in str(selected_unit) if c.isalnum() or c in (' ', '_')).rstrip()
                                zip_filename_single = f"BangKe_va_HoaDon_{safe_unit_name_zip}.zip"
                                
                                st.download_button(
                                    label=f"📥 Tải ZIP cho '{selected_unit}'",
                                    data=zip_data_single,
                                    file_name=zip_filename_single,
                                    mime="application/zip",
                                    use_container_width=True
                                )

                        st.divider()

                        # --- Part 2: Download for all units (Excel only) ---
                        st.markdown("###### Tải cho tất cả đơn vị (chỉ Bảng kê Excel)")
                        st.info(f"Dữ liệu sẽ được gom nhóm theo cột 'Đơn vị' được map từ file 'Tong hop _ Report.xlsx'.")

                        if st.button("📦 Tạo và Tải file .zip cho tất cả đơn vị"):
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                unique_units = df_merged[unit_col].dropna().unique()
                                st.info(f"Bắt đầu tạo {len(unique_units)} file Bảng kê cho các đơn vị...")
                                progress_bar_zip = st.progress(0)
                                for i, unit in enumerate(unique_units):
                                    df_unit = df_merged[df_merged[unit_col] == unit]
                                    start_date_unit = df_unit['Date_dt'].min()
                                    end_date_unit = df_unit['Date_dt'].max()
                                    excel_data_unit = generate_bang_ke_excel(df_unit, start_date_unit, end_date_unit, date_col_name, pickup_col_name, dropoff_col_name)
                                    safe_unit_name = "".join(c for c in str(unit) if c.isalnum() or c in (' ', '_')).rstrip()
                                    file_name = f"BangKe_{safe_unit_name}_{start_date_unit.strftime('%Y%m%d') if pd.notna(start_date_unit) else 'nodate'}_{end_date_unit.strftime('%Y%m%d') if pd.notna(end_date_unit) else 'nodate'}.xlsx"
                                    zip_file.writestr(file_name, excel_data_unit)
                                    progress_bar_zip.progress((i + 1) / len(unique_units))

                                zip_data = zip_buffer.getvalue()
                                
                                st.download_button(
                                    label=f"✅ Tải về ZIP ({len(unique_units)} files)",
                                    data=zip_data,
                                    file_name="BangKe_Theo_Don_Vi.zip",
                                    mime="application/zip",
                                    use_container_width=True,
                                    key='download_all_units'
                                )

                    st.divider()
                    # --- Tùy chọn 2: Tải file tổng hợp ---
                    st.markdown("##### 📥 Tải Bảng kê Tổng hợp (1 file)")
                    excel_data = generate_bang_ke_excel(df_merged, start_date, end_date, date_col_name, pickup_col_name, dropoff_col_name)
                    st.download_button(
                        label="📥 Tải về Bảng kê Tổng hợp",
                        data=excel_data,
                        file_name=f"BangKe_TongHop_{start_date.strftime('%Y%m%d') if pd.notna(start_date) else ''}_{end_date.strftime('%Y%m%d') if pd.notna(end_date) else ''}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except FileNotFoundError:
                st.error("Lỗi: Không tìm thấy file mẫu 'FileMau/BangKe.xlsx'. Vui lòng đảm bảo file tồn tại.")
            except Exception as e:
                st.error(f"Đã xảy ra lỗi khi tạo file Bảng kê: {e}")


        # --- Bảng Hóa đơn không có PDF ---
        if uploaded_zip_file is not None and 'pdf_filename' in df_merged.columns:
            with st.container(border=True):
                st.subheader("🚫 Hóa đơn không có file PDF")
                invoices_no_pdf = df_merged[df_merged['pdf_filename'].isnull()].copy()
                if not invoices_no_pdf.empty:
                    st.warning(f"Tìm thấy {len(invoices_no_pdf)} hóa đơn không có file PDF tương ứng trong tệp zip.")
                    # Chọn các cột cần hiển thị để người dùng dễ dàng xác định hóa đơn bị thiếu
                    display_cols = ['Employee Name', 'Booking ID', 'Date', 'Trip Type', 'Total Fare', 'pdf_link_key']
                    # Lọc ra những cột thực sự tồn tại trong dataframe để tránh lỗi
                    existing_cols = [col for col in display_cols if col in invoices_no_pdf.columns]
                    st.dataframe(invoices_no_pdf[existing_cols])
                else:
                    st.success("Tất cả hóa đơn trong danh sách đối chiếu đều có file PDF tương ứng.")

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
