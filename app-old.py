import json
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import streamlit as st

# Hàm đọc thông tin từ Streamlit Secrets
def get_credentials_from_secrets():
    try:
        # Lấy nội dung từ Streamlit Secrets
        credentials_json = st.secrets["GOOGLE_CREDENTIALS"]
        # Chuyển nội dung JSON thành dictionary
        credentials_info = json.loads(credentials_json)
        # Tạo thông tin xác thực từ dictionary
        credentials = Credentials.from_service_account_info(credentials_info)
        return credentials
    except KeyError:
        raise ValueError("GOOGLE_CREDENTIALS không được tìm thấy trong Streamlit Secrets!")
    except json.JSONDecodeError:
        raise ValueError("Dữ liệu GOOGLE_CREDENTIALS không hợp lệ!")

# Hàm kết nối đến Google Docs API
def connect_to_google_docs():
    credentials = get_credentials_from_secrets()
    service = build("docs", "v1", credentials=credentials)
    return service

# Hàm lấy số thứ tự hiện tại trong Google Docs
def get_current_entry_number(doc_id):
    service = connect_to_google_docs()
    document = service.documents().get(documentId=doc_id).execute()
    content = document.get("body").get("content", [])

    # Ghép nội dung thành chuỗi văn bản nếu có nội dung
    text = ""
    for element in content:
        if "textRun" in element.get("paragraph", {}).get("elements", [{}])[0]:
            text += element["paragraph"]["elements"][0]["textRun"]["content"]

    # Lọc các dòng bắt đầu bằng số thứ tự
    numbers = []
    for line in text.split("\n"):
        if line.strip().startswith(tuple(str(i) for i in range(10))):  # Kiểm tra dòng bắt đầu bằng số
            try:
                number = int(line.split(".")[0])  # Lấy số trước dấu "."
                numbers.append(number)
            except ValueError:
                continue

    # Trả về số thứ tự tiếp theo
    return max(numbers) + 1 if numbers else 1

# Hàm ghi dữ liệu vào Google Docs
def write_to_google_docs(doc_id, date, content):
    service = connect_to_google_docs()
    # Xác định số thứ tự mới và vị trí cuối tài liệu
    entry_number = get_current_entry_number(doc_id)
    document = service.documents().get(documentId=doc_id).execute()
    # Kiểm tra vị trí cuối nếu tài liệu rỗng
    end_index = document.get("body").get("content")[-1].get("endIndex", 1) - 1 if document.get("body").get("content") else 1

    # Định dạng nội dung
    formatted_content = "\n".join([f"- {line}" for line in content.split("\n")])
    requests = [
        # Thêm tiêu đề "Ngày"
        {
            "insertText": {
                "location": {"index": end_index},  # Chèn vào cuối tài liệu
                "text": f"{entry_number}. Ngày: {date}\n"
            }
        },
        # Thêm "Nội dung" in đậm và gạch dưới
        {
            "insertText": {
                "location": {"index": end_index + len(f"{entry_number}. Ngày: {date}\n")},
                "text": "Nội dung:\n"
            }
        },
        {
            "updateTextStyle": {
                "range": {
                    "startIndex": end_index + len(f"{entry_number}. Ngày: {date}\n"),
                    "endIndex": end_index + len(f"{entry_number}. Ngày: {date}\nNội dung:")
                },
                "textStyle": {"bold": True, "underline": True},
                "fields": "bold,underline"
            }
        },
        # Thêm nội dung với xuống dòng
        {
            "insertText": {
                "location": {"index": end_index + len(f"{entry_number}. Ngày: {date}\nNội dung:\n")},
                "text": f"{formatted_content}\n\n"
            }
        }
    ]
    # Gửi yêu cầu cập nhật đến Google Docs API
    service.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()

# Giao diện Streamlit
st.title("Nhật ký sự kiện khoa xét nghiệm")

# Trạng thái ứng dụng
if "event_content" not in st.session_state:
    st.session_state.event_content = ""

# Trường "Ngày" tự động lấy thời gian thực
current_datetime = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
st.write("Ngày:", current_datetime)

# Trường nhập nội dung sự kiện
event_content = st.text_area("Nội dung sự kiện", value=st.session_state.event_content)

# Xử lý khi bấm nút lưu
if st.button("Lưu vào Google Docs"):
    if not event_content:
        st.warning("Vui lòng nhập nội dung sự kiện!")
    else:
        try:
            # ID của Google Docs từ link bạn cung cấp
            doc_id = "1YRqAYASyH72iDfxnlFPaXwpnOlWp0A3XctIdwB8qcWI"
            write_to_google_docs(doc_id, current_datetime, event_content)
            st.success("Đã lưu thành công vào Google Docs!")
            st.session_state.event_content = ""  # Reset nội dung sau khi lưu
        except Exception as e:
            st.error(f"Có lỗi xảy ra: {e}")

# Nút "Tạo mới" để reset các trường nhập
if st.button("Tạo mới"):
    st.session_state.event_content = ""  # Reset nội dung sự kiện
