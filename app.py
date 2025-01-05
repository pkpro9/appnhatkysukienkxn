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
                "text": "- Địa điểm:\n"
            }
        },
        {
            "insertText": {
                "location": {"index": end_index + len(f"{entry_number}. Ngày: {date}\n- Địa điểm:\n")},
                "text": content["location"]
            }
        },
        {
            "insertText": {
                "location": {"index": end_index + len(f"{entry_number}. Ngày: {date}\n- Địa điểm:\n{content['location']}\n")},
                "text": "- Thành phần tham dự:\n"
            }
        },
        {
            "insertText": {
                "location": {"index": end_index + len(f"{entry_number}. Ngày: {date}\n- Địa điểm:\n{content['location']}\n- Thành phần tham dự:\n")},
                "text": content["attendees"]
            }
        },
        {
            "insertText": {
                "location": {"index": end_index + len(f"{entry_number}. Ngày: {date}\n- Địa điểm:\n{content['location']}\n- Thành phần tham dự:\n{content['attendees']}\n")},
                "text": "- Nội dung cuộc họp:\n"
            }
        },
        {
            "insertText": {
                "location": {"index": end_index + len(f"{entry_number}. Ngày: {date}\n- Địa điểm:\n{content['location']}\n- Thành phần tham dự:\n{content['attendees']}\n- Nội dung cuộc họp:\n")},
                "text": content["meeting_content"]
            }
        }
    ]
    service.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()

# Giao diện Streamlit
st.title("Quản lý thông tin")

menu = st.sidebar.radio("Chọn chức năng", ["Nhật ký sự kiện Khoa XN", "Giao ban viện", "Biên bản họp KXN"])

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

elif menu == "Biên bản họp KXN":
    st.header("Biên bản họp KXN")

    if "meeting_minutes_date" not in st.session_state:
        timezone = pytz.timezone("Asia/Ho_Chi_Minh")
        st.session_state.meeting_minutes_date = datetime.now(timezone).strftime("%d-%m-%Y %H:%M:%S")

    meeting_minutes_date = st.text_input("Ngày:", value=st.session_state.meeting_minutes_date)
    st.session_state.meeting_minutes_date = meeting_minutes_date

    location = st.text_area("Địa điểm")
    attendees = st.text_area("Thành phần tham dự")
    meeting_content = st.text_area("Nội dung cuộc họp")

    if st.button("Lưu vào Google Docs (Biên bản họp KXN)"):
        if not location and not attendees and not meeting_content:
            st.warning("Vui lòng nhập đầy đủ thông tin cuộc họp!")
        else:
            try:
                doc_id = "17bJaGses0Pss7AxiWvrKNiV75PBdszYytiovAbITGlE"
                content = {
                    "location": f"\n+ {location.replace('\n', '\n+ ')}\n",
                    "attendees": f"\n+ {attendees.replace('\n', '\n+ ')}\n",
                    "meeting_content": f"\n+ {meeting_content.replace('\n', '\n+ ')}\n"
                }
                write_to_google_docs(doc_id, st.session_state.meeting_minutes_date, content)
                st.success("Đã lưu thành công vào Google Docs!")
            except Exception as e:
                st.error(f"Có lỗi xảy ra: {e}")
