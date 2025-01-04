import os
import base64
import streamlit as st
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Đọc và giải mã biến môi trường chứa file JSON
def get_credentials_from_env():
    # Đọc biến môi trường
    encoded_json = os.getenv("SERVICE_ACCOUNT_JSON")
    if not encoded_json:
        raise ValueError("Biến môi trường SERVICE_ACCOUNT_JSON không tồn tại!")
    
    # Giải mã chuỗi Base64 và trả về thông tin xác thực
    json_data = base64.b64decode(encoded_json).decode("utf-8")
    credentials = Credentials.from_service_account_info(eval(json_data), scopes=["https://www.googleapis.com/auth/documents"])
    return credentials

# Kết nối tới Google Docs API
def connect_to_google_docs():
    credentials = get_credentials_from_env()
    service = build("docs", "v1", credentials=credentials)
    return service

# Ghi nội dung vào Google Docs
def write_to_google_docs(doc_id, date, content):
    service = connect_to_google_docs()

    # Tạo nội dung cần thêm vào Google Docs
    requests = [
        {
            "insertText": {
                "location": {
                    "index": 1  # Chèn vào đầu tài liệu
                },
                "text": f"1. {date}\n"
                        f"2. {content}\n\n"
            }
        }
    ]

    # Gửi yêu cầu thêm nội dung vào tài liệu
    result = service.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()
    return result

# Tạo giao diện bằng Streamlit
st.title("Nhật ký sự kiện khoa xét nghiệm")

# Trường "Ngày" tự động lấy thời gian thực
current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
st.write("Ngày:", current_datetime)

# Trường nhập nội dung sự kiện
event_content = st.text_area("Nội dung sự kiện")

# Xử lý khi bấm nút lưu
if st.button("Lưu vào Google Docs"):
    if not event_content:
        st.warning("Vui lòng nhập nội dung sự kiện!")
    else:
        try:
            # ID Google Docs (trích xuất từ URL)
            doc_id = "1YRqAYASyH72iDfxnlFPaXwpnOlWp0A3XctIdwB8qcWI"
            
            # Ghi dữ liệu vào Google Docs
            result = write_to_google_docs(doc_id, current_datetime, event_content)
            st.success("Đã lưu thành công vào Google Docs!")
        except Exception as e:
            st.error(f"Có lỗi xảy ra: {e}")
