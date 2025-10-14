# Copyright by hiennm22, LocTH5 - BM UDPM
import streamlit as st
import pandas as pd
import os
import glob
import zipfile
import io
import re
import base64
import json
import os.path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Cố gắng import pypdf và hướng dẫn cài đặt nếu thiếu
try:
    import pypdf
except ImportError:
    st.error("Thư viện pypdf chưa được cài đặt. Vui lòng chạy lệnh sau trong terminal: pip install pypdf")
    st.stop()

# --- HÀM HỖ TRỢ ---
def find_col(df, possibilities):
    """Finds the first column in a dataframe that exists from a list of possibilities."""
    for p in possibilities:
        if p in df.columns:
            return p
    return None

def load_mapping_data():
    """Đọc file Excel mapping và trả về 2 dictionaries:
    1. Employee Name -> Đơn vị
    2. Đơn vị -> List of Emails
    """
    try:
        df_mapping = pd.read_excel("FileMau/Tong hop _ Report.xlsx")
        name_col = df_mapping.columns[1]
        email_col = df_mapping.columns[3]
        unit_col = df_mapping.columns[4]
        df_mapping = df_mapping.dropna(subset=[name_col, unit_col])
        employee_to_unit_map = df_mapping.set_index(name_col)[unit_col].to_dict()
        df_email_map = df_mapping.dropna(subset=[email_col])
        # Group by unit and create a list of unique emails for each unit
        unit_to_email_map = df_email_map.groupby(unit_col)[email_col].apply(lambda x: list(x.unique())).to_dict()
        return employee_to_unit_map, unit_to_email_map
    except FileNotFoundError:
        st.error("Lỗi: Không tìm thấy file mapping 'FileMau/Tong hop _ Report.xlsx'. Vui lòng đảm bảo file tồn tại.")
        st.stop()
    except IndexError:
        st.error("Lỗi: File mapping 'Tong hop _ Report.xlsx' không có đủ 5 cột (để lấy cột B, D, và E).")
        st.stop()
    except Exception as e:
        st.error(f"Lỗi khi đọc file mapping: {e}")
        st.stop()

# --- HÀM HỖ TRỢ OAUTH2 ---
SCOPES = ['https://www.googleapis.com/auth/gmail.send', 'https://www.googleapis.com/auth/gmail.readonly']
TOKEN_FILE = "token.json"

def get_google_credentials(credentials_json_content):
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            st.info("Refreshing expired credentials...")
            creds.refresh(Request())
        else:
            st.info("Credentials not found or invalid, starting authorization flow...")
            flow = InstalledAppFlow.from_client_config(
                json.loads(credentials_json_content), SCOPES)
            # run_local_server will open a browser tab for user authorization
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
            st.success(f"Credentials saved to {TOKEN_FILE}")
    return creds

def send_gmail_message(credentials, to, subject, body, attachments=None):
    """Sends an email with multiple attachments using Gmail API."""
    try:
        service = build('gmail', 'v1', credentials=credentials)
        user_profile = service.users().getProfile(userId='me').execute()
        sender_email = user_profile['emailAddress']
        sender_name = "Hệ thống đối chiếu tự động"
        
        message = MIMEMultipart()
        message['to'] = to
        message['from'] = formataddr((sender_name, sender_email))
        message['subject'] = subject
        msg_body = MIMEText(body, 'plain', 'utf-8')
        message.attach(msg_body)

        if attachments:
            for attachment in attachments:
                if attachment and attachment.get('data') and attachment.get('filename'):
                    part = MIMEApplication(attachment['data'], Name=attachment['filename'])
                    part['Content-Disposition'] = f'attachment; filename="{attachment["filename"]}"'
                    message.attach(part)

        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        create_message = {'raw': encoded_message}
        send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    except HttpError as error:
        st.error(f"An error occurred while sending email: {error}")
        raise error

# --- GIAO DIỆN CHÍNH ---
st.set_page_config(page_title="Đối chiếu FPT", layout="wide", page_icon="📊")
st.title("📊 Đối chiếu dữ liệu Grab & Báo cáo PDF")
st.write("Tải lên các tệp của bạn để bắt đầu đối chiếu và xử lý.")
st.caption("Copyright by LocTH5, Hiennm22 - BM UDPM")

# --- GIAO DIỆN NHẬP LIỆU ---
with st.container(border=True):
    st.subheader("Tải lên các file cần thiết")
    col1, col2, col3 = st.columns([2, 2, 3])
    file_types = ["csv", "xls", "xlsx"]
    with col1:
        uploaded_transport_file = st.file_uploader("1. File Transport", type=file_types)
    with col2:
        uploaded_invoice_file = st.file_uploader("2. File Hóa đơn", type=file_types)
    with col3:
        uploaded_zip_file = st.file_uploader("3. Folder Báo cáo (.zip)", type=["zip"])

# --- CẤU HÌNH OAUTH 2.0 ---
# The app now automatically loads 'credentials.json' from the local directory.
CREDENTIALS_FILE = "credentials.json"
st.session_state.credentials_loaded = False
if os.path.exists(CREDENTIALS_FILE):
    with open(CREDENTIALS_FILE, 'r', encoding='utf-8') as f:
        st.session_state.credentials_json_content = f.read()
    st.session_state.credentials_loaded = True

# --- BẮT ĐẦU XỬ LÝ KHI CÓ ĐỦ FILE ---
if uploaded_transport_file is not None and uploaded_invoice_file is not None:
    try:
        employee_to_unit_map, unit_to_email_map = load_mapping_data()

        # (The rest of the data processing logic remains the same as before)
        # --- 1. ĐỌC VÀ LÀM SẠCH DỮ LIỆU GỐC ---
        if uploaded_transport_file.name.endswith('.csv'):
            df_transport = pd.read_csv(uploaded_transport_file, skiprows=8)
        else:
            df_transport = pd.read_excel(uploaded_transport_file, skiprows=8)

        if uploaded_invoice_file.name.endswith('.csv'):
            df_invoice = pd.read_csv(uploaded_invoice_file)
        elif uploaded_invoice_file.name.endswith('.xls'):
            try:
                df_invoice = pd.read_html(uploaded_invoice_file)[0]
            except Exception:
                df_invoice = pd.read_excel(uploaded_invoice_file, engine='xlrd')
        else:
            df_invoice = pd.read_excel(uploaded_invoice_file)

        df_transport.columns = df_transport.columns.str.strip()
        df_invoice.columns = df_invoice.columns.str.strip()

        if len(df_transport.columns) > 9:
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
            rename_dict[df_invoice.columns[4]] = 'GEMINI_NGAY_HD_INVOICE'
            rename_dict[df_invoice.columns[5]] = 'HINH_THUC_TT'
            rename_dict[df_invoice.columns[6]] = 'TIEN_TRC_THUE'
            rename_dict[df_invoice.columns[7]] = 'TIEN_THUE8'
            rename_dict[df_invoice.columns[8]] = 'TONG_TIEN'
            rename_dict[df_invoice.columns[15]] = 'NGAY_BOOKING'
            rename_dict[df_invoice.columns[16]] = 'SO_HOA_DON'
        df_invoice.rename(columns=rename_dict, inplace=True)

        matching_ids = list(set(df_transport['Booking ID'].dropna()) & set(df_invoice['Booking'].dropna()))
        if not matching_ids:
            st.warning("Không tìm thấy Booking ID nào trùng khớp giữa hai file đầu vào.")
            st.stop()

        df_merged = pd.merge(df_transport[df_transport['Booking ID'].isin(matching_ids)], df_invoice[df_invoice['Booking'].isin(matching_ids)], left_on='Booking ID', right_on='Booking', suffixes=('_transport', '_invoice'))
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
                                    if code: found_code = code
                            ngay_hd_str = "Không tìm thấy"
                            match = re.search(r'Ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})', text, re.IGNORECASE)
                            if match:
                                day, month, year = match.groups()
                                ngay_hd_str = f"{day.zfill(2)}/{month.zfill(2)}/{year}"
                            pdf_data.append({'pdf_link_key_str': key_from_filename, 'Mã hóa đơn từ PDF': found_code, 'Ngay_HD_pdf': ngay_hd_str, 'pdf_content': pdf_content, 'pdf_filename': os.path.basename(filename)})
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
            st.warning(f"### ⚠️ Chú ý: Có {count_no_pdf} hóa đơn không có file PDF tương ứng.")
            with st.expander("Xem danh sách và thống kê các hóa đơn thiếu PDF"):
                df_missing_pdfs = df_merged[df_merged['pdf_filename'].isnull()]
                
                # 1. Display statistics table
                st.markdown("#### Thống kê theo Đơn vị")
                missing_stats = df_missing_pdfs.groupby('Đơn vị').agg(
                    so_hoa_don_thieu=('Booking ID', 'count'),
                    tong_tien_thieu=('Total Fare', 'sum')
                ).reset_index()
                missing_stats.rename(columns={
                    'so_hoa_don_thieu': 'Số hóa đơn thiếu PDF',
                    'tong_tien_thieu': 'Tổng tiền (ước tính)'
                }, inplace=True)
                st.dataframe(missing_stats)

                # 2. Display the raw list
                st.markdown("---")
                st.markdown("#### Danh sách chi tiết")
                
                cols_to_show = ['Booking ID', 'Employee Name', 'Đơn vị', 'pdf_link_key', 'Total Fare']
                date_cols = ['Date', 'Date of Trip', 'Trip Date', 'Ngày', 'Date & Time (GMT+7)']
                available_date_col = find_col(df_missing_pdfs, date_cols)
                if available_date_col:
                    cols_to_show.append(available_date_col)

                existing_cols_to_show = [col for col in cols_to_show if col in df_missing_pdfs.columns]
                st.dataframe(df_missing_pdfs[existing_cols_to_show])

        st.header("Kết quả đối chiếu")

        with st.container(border=True):
            st.subheader("Bảng thống kê tổng hợp")

            required_cols = {'Employee Name', 'Booking ID', 'Total Fare'}
            if not required_cols.issubset(df_merged.columns):
                st.warning("Không thể tạo bảng thống kê tổng hợp vì thiếu dữ liệu")
            else:
                agg_dict = {'So chuyen': ('Booking ID', 'count'), 'Tong tien (VND)': ('Total Fare', 'sum')}
                summary_df = (
                    df_merged
                    .groupby('Employee Name')
                    .agg(**agg_dict)
                    .reset_index()
                    .sort_values('Tong tien (VND)', ascending=False)
                    .reset_index(drop=True)
                )

                if summary_df.empty:
                    st.info("Không có đủ dữ liệu để thống kê.")
                else:
                    header_cols = st.columns([3, 1, 2])
                    header_cols[0].markdown("Tên Người dùng")
                    header_cols[1].markdown("Số chuyến")
                    header_cols[2].markdown("Tổng tiền (VND)")

                    current_employees = summary_df['Employee Name'].tolist()
                    expanded_state = st.session_state.setdefault("expanded_employees", {})
                    stale_keys = [name for name in expanded_state.keys() if name not in current_employees]
                    for name in stale_keys:
                        del expanded_state[name]
                    for name in current_employees:
                        expanded_state.setdefault(name, False)

                    def toggle_employee_expansion(employee_name: str) -> None:
                        state = st.session_state["expanded_employees"]
                        state[employee_name] = not state.get(employee_name, False)

                    date_cols = ['Date', 'Date of Trip', 'Trip Date', 'Ngay', 'Date & Time (GMT+7)']
                    date_col_name = find_col(df_merged, date_cols)
                    pickup_col_name = 'GEMINI_PICKUP_ADDRESS' if 'GEMINI_PICKUP_ADDRESS' in df_merged.columns else None
                    dropoff_col_name = 'GEMINI_DROPOFF_ADDRESS' if 'GEMINI_DROPOFF_ADDRESS' in df_merged.columns else None

                    for idx, row in summary_df.iterrows():
                        employee_name = row['Employee Name']
                        expanded = st.session_state["expanded_employees"][employee_name]
                        row_container = st.container()
                        with row_container:
                            row_cols = st.columns([3, 1, 2])
                            with row_cols[0]:
                                label = f"{'✅ ' if expanded else ''}{employee_name}"
                                st.button(
                                    label,
                                    key=f"summary_row_{idx}",
                                    use_container_width=True,
                                    on_click=toggle_employee_expansion,
                                    args=(employee_name,),
                                )
                            with row_cols[1]:
                                st.write(int(row['So chuyen']))
                            with row_cols[2]:
                                st.write(f"{row['Tong tien (VND)']:,.0f}")

                        if st.session_state["expanded_employees"].get(employee_name):
                            employee_df = df_merged[df_merged['Employee Name'] == employee_name]
                            usage_date_col = 'NGAY_BOOKING' if 'NGAY_BOOKING' in employee_df.columns else date_col_name
                            detail_cols = [
                                'Employee Name',
                                'Booking ID',
                                pickup_col_name,
                                dropoff_col_name,
                                'GEMINI_NGAY_HD_INVOICE',
                                'HINH_THUC_TT',
                                'TIEN_TRC_THUE',
                                'TIEN_THUE8',
                                'TONG_TIEN',
                                usage_date_col,
                                'SO_HOA_DON',
                            ]
                            final_cols = [col for col in detail_cols if col and col in employee_df.columns]

                            detail_df = employee_df[final_cols].copy()
                            detail_df.insert(0, "STT", range(1, len(detail_df) + 1))
                            if 'SO_HOA_DON' in detail_df.columns:
                                detail_df['SO_HOA_DON'] = (
                                    detail_df['SO_HOA_DON']
                                    .astype(str)
                                    .str.split('_')
                                    .str[0]
                                    .replace('nan', '')
                                )

                            rename_map = {
                                'Employee Name': 'Người sử dụng',
                                'Booking ID': 'Mã đặt chỗ',
                                'GEMINI_PICKUP_ADDRESS': 'Điểm đón',
                                'GEMINI_DROPOFF_ADDRESS': 'Điểm đến',
                                'GEMINI_NGAY_HD_INVOICE': 'Ngày HĐ',
                                'HINH_THUC_TT': 'Hình thức thanh toán',
                                'TIEN_TRC_THUE': 'Tổng tiền trước thuế',
                                'TIEN_THUE8': 'Tổng tiền thuế (8%)',
                                'TONG_TIEN': 'Tổng tiền đã có thuế',
                                'NGAY_BOOKING': 'Ngày sử dụng',
                                'SO_HOA_DON': 'Số hóa đơn',
                            }
                            if usage_date_col and usage_date_col not in rename_map:
                                rename_map[usage_date_col] = 'Ngày sử dụng'
                            detail_df.rename(columns={k: v for k, v in rename_map.items() if k in detail_df.columns}, inplace=True)

                            employee_display_col = rename_map.get('Employee Name', 'Employee Name')
                            money_cols = [
                                rename_map.get('TIEN_TRC_THUE', 'TIEN_TRC_THUE'),
                                rename_map.get('TIEN_THUE8', 'TIEN_THUE8'),
                                rename_map.get('TONG_TIEN', 'TONG_TIEN'),
                            ]
                            total_row = {col: "" for col in detail_df.columns}
                            total_row['STT'] = ""
                            if employee_display_col in total_row:
                                total_row[employee_display_col] = 'Tổng cộng'
                            for money_col in money_cols:
                                if money_col in detail_df.columns:
                                    numeric_values = pd.to_numeric(detail_df[money_col], errors='coerce')
                                    total_row[money_col] = numeric_values.sum()
                            detail_df = pd.concat([detail_df, pd.DataFrame([total_row])], ignore_index=True)

                            for money_col in money_cols:
                                if money_col in detail_df.columns:
                                    numeric_series = pd.to_numeric(detail_df[money_col], errors='coerce')
                                    detail_df[money_col] = numeric_series.apply(
                                        lambda value: f"{value:,.0f}" if pd.notna(value) else ""
                                    )

                            st.dataframe(detail_df, use_container_width=True, hide_index=True)

                        st.markdown('<hr style="margin-top:0.25rem; margin-bottom:0.25rem;">', unsafe_allow_html=True)
        try:
            date_cols = ['Date', 'Date of Trip', 'Trip Date', 'Ngày', 'Date & Time (GMT+7)']
            date_col_name = find_col(df_merged, date_cols)
            pickup_col_name = 'GEMINI_PICKUP_ADDRESS' if 'GEMINI_PICKUP_ADDRESS' in df_merged.columns else None
            dropoff_col_name = 'GEMINI_DROPOFF_ADDRESS' if 'GEMINI_DROPOFF_ADDRESS' in df_merged.columns else None

            if date_col_name is None:
                try:
                    uploaded_transport_file.seek(0)
                    header_list = pd.read_csv(uploaded_transport_file, skiprows=8, nrows=1).columns.tolist() if uploaded_transport_file.name.endswith('.csv') else pd.read_excel(uploaded_transport_file, skiprows=8, nrows=1).columns.tolist()
                    st.error(f"**Lỗi Bảng kê: Không thể tìm thấy cột Ngày tháng.** Vui lòng cho tôi biết tên cột ngày tháng chính xác từ danh sách bên dưới. **Các cột tìm được:** `{header_list}`")
                except Exception as e:
                    st.error(f"Lỗi Bảng kê: Không tìm thấy cột ngày tháng. Lỗi khi đọc cột: {e}")
            else:
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
                        if _p_col: ws.cell(row=start_row + i, column=4, value=row[_p_col])
                        if _do_col: ws.cell(row=start_row + i, column=5, value=row[_do_col])
                        if 'GEMINI_NGAY_HD_INVOICE' in row and pd.notna(row['GEMINI_NGAY_HD_INVOICE']): ws.cell(row=start_row + i, column=6, value=row['GEMINI_NGAY_HD_INVOICE'])
                        ws.cell(row=start_row + i, column=7, value=row['HINH_THUC_TT'])
                        ws.cell(row=start_row + i, column=8, value="{:,.0f}".format(row['TIEN_TRC_THUE']))
                        ws.cell(row=start_row + i, column=9, value="{:,.0f}".format(row['TIEN_THUE8']))
                        ws.cell(row=start_row + i, column=10, value="{:,.0f}".format(row['TONG_TIEN']))
                        ws.cell(row=start_row + i, column=11, value=row['NGAY_BOOKING'])
                        ws.cell(row=start_row + i, column=12, value=row['pdf_link_key'])
                    total_row_index = start_row + len(_df)
                    total_label_cell = ws.cell(row=total_row_index, column=7, value="Tổng cộng"); total_label_cell.font = Font(bold=True)
                    total_value_cell_8 = ws.cell(row=total_row_index, column=8, value="{:,.0f}".format(_df['TIEN_TRC_THUE'].sum())); total_value_cell_8.font = Font(bold=True)
                    total_value_cell_9 = ws.cell(row=total_row_index, column=9, value="{:,.0f}".format(_df['TIEN_THUE8'].sum())); total_value_cell_9.font = Font(bold=True)
                    total_value_cell_10 = ws.cell(row=total_row_index, column=10, value="{:,.0f}".format(_df['TONG_TIEN'].sum())); total_value_cell_10.font = Font(bold=True)
                    excel_buffer = io.BytesIO(); wb.save(excel_buffer); return excel_buffer.getvalue()

                st.subheader("📧 Gửi Bảng Kê qua Email (bằng Gmail)")
                if 'credentials_json_content' not in st.session_state:
                    st.warning("Vui lòng tải file `credentials.json` ở trên để kích hoạt chức năng gửi email.")
                else:
                    # --- SINGLE SEND ---
                    st.markdown("###### Gửi cho 1 đơn vị")
                    unit_col = 'Đơn vị'
                    unique_units_for_select_email = sorted(df_merged[unit_col].dropna().unique())
                    selected_unit_email = st.selectbox("Chọn đơn vị để gửi email", unique_units_for_select_email, key="email_unit_select")

                    if selected_unit_email:
                        recipient_emails = unit_to_email_map.get(selected_unit_email, [])
                        if recipient_emails:
                            st.info(f"Bảng kê cho '{selected_unit_email}' sẽ được gửi đến các địa chỉ sau:")
                            st.markdown(f"**{', '.join(recipient_emails)}**")
                        else:
                            st.warning(f"Không tìm thấy địa chỉ email cho đơn vị '{selected_unit_email}'.")

                        if st.button(f"📧 Gửi Email đến '{selected_unit_email}'", use_container_width=True, key="send_email_btn"):
                            if not recipient_emails:
                                st.error(f"Không thể gửi email: Không tìm thấy email cho đơn vị '{selected_unit_email}'.")
                            else:
                                to_field = ", ".join(recipient_emails)
                                with st.spinner(f"Đang xác thực và gửi email đến {to_field}..."):
                                    try:
                                        creds = get_google_credentials(st.session_state.credentials_json_content)
                                        df_unit = df_merged[df_merged[unit_col] == selected_unit_email]
                                        
                                        # 1. Create Excel attachment
                                        excel_data_email = generate_bang_ke_excel(df_unit, df_unit['Date_dt'].min(), df_unit['Date_dt'].max(), date_col_name, pickup_col_name, dropoff_col_name)
                                        safe_unit_name = "".join(c for c in str(selected_unit_email) if c.isalnum() or c in (' ', '_')).rstrip()
                                        excel_filename = f"BangKe_{safe_unit_name}.xlsx"
                                        attachments = [{'data': excel_data_email, 'filename': excel_filename}]

                                        # 2. Create zipped PDF attachments for each employee
                                        df_unit_with_pdfs = df_unit[df_unit['pdf_content'].notna()]
                                        if not df_unit_with_pdfs.empty:
                                            employee_groups = df_unit_with_pdfs.groupby('Employee Name')
                                            st.info(f"Đang tạo và đính kèm {len(employee_groups)} file .zip (chứa PDF) cho từng nhân viên...")
                                            for employee_name, employee_df in employee_groups:
                                                zip_buffer = io.BytesIO()
                                                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_f:
                                                    for _, row in employee_df.iterrows():
                                                        pdf_filename = os.path.basename(row['pdf_filename'])
                                                        zip_f.writestr(pdf_filename, row['pdf_content'])
                                                
                                                safe_employee_name = "".join(c for c in str(employee_name) if c.isalnum() or c in (' ', '_')).rstrip()
                                                zip_filename = f"{safe_employee_name}.zip"
                                                
                                                attachments.append({
                                                    'data': zip_buffer.getvalue(),
                                                    'filename': zip_filename
                                                })

                                        # 3. Send email
                                        subject = f"HOA DON GRAP"
                                        body = f"Kính gửi Cơ sở {selected_unit_email},\n\nTrung tâm xin gửi hóa đơn Grap phát sinh trong kỳ. Cán bộ thanh toán cơ sở vui lòng xem các file bảng kê và hóa đơn (nếu có) được đính kèm trong email này và thực hiện hồ sơ thanh toán đúng hạn.\n\nMọi thông tin thắc mắc, xin vui lòng liên hệ: lientt3@fe.edu.vn\nĐây là hệ thống đối chiếu tự động, vui lòng không reply email.\n\nTrân trọng"
                                        send_gmail_message(creds, to_field, subject, body, attachments)
                                        st.success(f"✅ Đã gửi email thành công đến {to_field} cho đơn vị '{selected_unit_email}'.")
                                    except Exception as e:
                                        st.error(f"Lỗi khi gửi email: {e}")

                    # --- BULK SEND ---
                    st.divider()
                    st.markdown("###### Gửi cho tất cả các đơn vị")
                    if st.button("🚀 Gửi Email cho TẤT CẢ các đơn vị", use_container_width=True, key="send_all_emails_btn"):
                        with st.spinner("Bắt đầu quá trình gửi email hàng loạt..."):
                            creds = get_google_credentials(st.session_state.credentials_json_content)
                            units_to_email = sorted(df_merged[unit_col].dropna().unique())
                            progress_bar = st.progress(0, text="Bắt đầu...")
                            success_count = 0
                            failed_units = []
                            for i, unit in enumerate(units_to_email):
                                progress_text = f"Đang xử lý: {unit} ({i+1}/{len(units_to_email)})"; progress_bar.progress((i + 1) / len(units_to_email), text=progress_text)
                                recipient_emails = unit_to_email_map.get(unit, [])
                                if not recipient_emails:
                                    failed_units.append((unit, "Không tìm thấy email trong file mapping."))
                                    continue
                                try:
                                    df_unit = df_merged[df_merged[unit_col] == unit]
                                    
                                    # 1. Create Excel attachment
                                    excel_data_email = generate_bang_ke_excel(df_unit, df_unit['Date_dt'].min(), df_unit['Date_dt'].max(), date_col_name, pickup_col_name, dropoff_col_name)
                                    safe_unit_name = "".join(c for c in str(unit) if c.isalnum() or c in (' ', '_')).rstrip()
                                    excel_filename = f"BangKe_{safe_unit_name}.xlsx"
                                    attachments = [{'data': excel_data_email, 'filename': excel_filename}]

                                    # 2. Create zipped PDF attachments for each employee in the unit
                                    df_unit_with_pdfs = df_unit[df_unit['pdf_content'].notna()]
                                    if not df_unit_with_pdfs.empty:
                                        for employee_name, employee_df in df_unit_with_pdfs.groupby('Employee Name'):
                                            zip_buffer = io.BytesIO()
                                            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_f:
                                                for _, row in employee_df.iterrows():
                                                    pdf_filename = os.path.basename(row['pdf_filename'])
                                                    zip_f.writestr(pdf_filename, row['pdf_content'])
                                            
                                            safe_employee_name = "".join(c for c in str(employee_name) if c.isalnum() or c in (' ', '_')).rstrip()
                                            zip_filename = f"{safe_employee_name}.zip"
                                            
                                            attachments.append({
                                                'data': zip_buffer.getvalue(),
                                                'filename': zip_filename
                                            })

                                    # 3. Send email
                                    subject = f"Bảng kê đối chiếu Grab cho đơn vị '{unit}'"
                                    body = f"Kính gửi Quý đơn vị {unit},\n\nVui lòng xem các file bảng kê và hóa đơn (nếu có) được đính kèm trong email này.\n\nTrân trọng,\nSerder mail."
                                    to_field = ", ".join(recipient_emails)
                                    send_gmail_message(creds, to_field, subject, body, attachments)
                                    success_count += 1
                                except Exception as e:
                                    failed_units.append((unit, str(e)))
                            progress_bar.empty()
                            st.success(f"✅ Hoàn tất! Đã gửi thành công {success_count}/{len(units_to_email)} email.")
                            if failed_units:
                                st.error(f"❌ Có {len(failed_units)} email gửi thất bại.")
                                with st.expander("Xem chi tiết lỗi"):
                                    for unit, reason in failed_units: st.write(f"- **{unit}**: {reason}")
        except Exception as e:
            st.error(f"Đã xảy ra lỗi khi tạo file Bảng kê hoặc gửi mail: {e}")

    except Exception as e:
        st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")
        st.exception(e)
