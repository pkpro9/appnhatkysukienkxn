import streamlit as st
from datetime import datetime
import pytz
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Hàm kết nối đến Google Docs API
def connect_to_google_docs():
    credentials_info = {
        "type": st.secrets["gcp_service_account"]["type"],
        "project_id": st.secrets["gcp_service_account"]["project_id"],
        "private_key_id": st.secrets["gcp_service_account"]["private_key_id"],
        "private_key": st.secrets["gcp_service_account"]["private_key"],
        "client_email": st.secrets["gcp_service_account"]["client_email"],
        "client_id": st.secrets["gcp_service_account"]["client_id"],
        "auth_uri": st.secrets["gcp_service_account"]["auth_uri"],
        "token_uri": st.secrets["gcp_service_account"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["gcp_service_account"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["gcp_service_account"]["client_x509_cert_url"]
    }

    credentials = Credentials.from_service_account_info(credentials_info, scopes=["https://www.googleapis.com/auth/documents"])
    service = build("docs", "v1", credentials=credentials)
    return service

# Hàm lấy số thứ tự hiện tại trong Google Docs
def get_current_entry_number(doc_id):
    service = connect_to_google_docs()
    document = service.documents().get(documentId=doc_id).execute()
    content = document.get("body").get("content")

    text = ""
    for element in content:
        if "textRun" in element.get("paragraph", {}).get("elements", [{}])[0]:
            text += element["paragraph"]["elements"][0]["textRun"]["content"]

    numbers = []
    for line in text.split("\n"):
        if line.strip().startswith(tuple(str(i) for i in range(10))):
            try:
                number = int(line.split(".")[0])
                numbers.append(number)
            except ValueError:
                continue

    return max(numbers) + 1 if numbers else 1

# Hàm ghi dữ liệu vào Google Docs
def write_to_google_docs(doc_id, date, content):
    service = connect_to_google_docs()
    entry_number = get_current_entry_number(doc_id)
    document = service.documents().get(documentId=doc_id).execute()
    end_index = document.get("body").get("content")[-1].get("endIndex") - 1 if len(document.get("body").get("content")) > 1 else 1

    formatted_content = "\n".join([f"- {line}" for line in content.split("\n")])
    requests = [
        {
            "insertText": {
                "location": {"index": end_index},
                "text": f"{entry_number}. Ngày: {date}\n"
            }
        },
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
        {
            "insertText": {
                "location": {"index": end_index + len(f"{entry_number}. Ngày: {date}\nNội dung:\n")},
                "text": f"{formatted_content}\n\n"
            }
        }
    ]
    service.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()

# Giao diện Streamlit
st.title("Quản lý thông tin")

menu = st.sidebar.radio("Chọn chức năng", ["Nhật ký sự kiện Khoa XN", "Giao ban viện"])

if menu == "Nhật ký sự kiện Khoa XN":
    st.header("Nhật ký sự kiện Khoa XN")

    if "event_content" not in st.session_state:
        st.session_state.event_content = ""
    if "event_date" not in st.session_state:
        timezone = pytz.timezone("Asia/Ho_Chi_Minh")
        st.session_state.event_date = datetime.now(timezone).strftime("%d-%m-%Y %H:%M:%S")

    event_date = st.text_input("Ngày:", value=st.session_state.event_date)
    st.session_state.event_date = event_date

    event_content = st.text_area("Nội dung sự kiện", value=st.session_state.event_content)

    if st.button("Lưu vào Google Docs"):
        if not event_content:
            st.warning("Vui lòng nhập nội dung sự kiện!")
        else:
            try:
                doc_id = "1YRqAYASyH72iDfxnlFPaXwpnOlWp0A3XctIdwB8qcWI"
                write_to_google_docs(doc_id, st.session_state.event_date, event_content)
                st.success("Đã lưu thành công vào Google Docs!")
                st.session_state.event_content = ""
            except Exception as e:
                st.error(f"Có lỗi xảy ra: {e}")

    if st.button("Tạo mới"):
        st.session_state.event_content = ""
        timezone = pytz.timezone("Asia/Ho_Chi_Minh")
        st.session_state.event_date = datetime.now(timezone).strftime("%d-%m-%Y %H:%M:%S")

elif menu == "Giao ban viện":
    st.header("Giao ban viện")

    if "meeting_date" not in st.session_state:
        timezone = pytz.timezone("Asia/Ho_Chi_Minh")
        st.session_state.meeting_date = datetime.now(timezone).strftime("%d-%m-%Y %H:%M:%S")

    meeting_date = st.text_input("Ngày:", value=st.session_state.meeting_date)
    st.session_state.meeting_date = meeting_date

    expertise = st.text_area("Chuyên môn")
    dissemination = st.text_area("Phổ biến")

    if st.button("Lưu vào Google Docs (Giao ban viện)"):
        if not expertise and not dissemination:
            st.warning("Vui lòng nhập thông tin chuyên môn và phổ biến!")
        else:
            try:
                doc_id = "1wdpbDQeLhyHhrjN_6GPZbH4s_ZiqkOyU4J2NvlPmpWY"
                content = f"Chuyên môn: {expertise}\nPhổ biến: {dissemination}"
                write_to_google_docs(doc_id, st.session_state.meeting_date, content)
                st.success("Đã lưu thành công vào Google Docs!")
            except Exception as e:
                st.error(f"Có lỗi xảy ra: {e}")
