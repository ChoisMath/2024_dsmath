import smtplib
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import pandas as pd
import pytz
from datetime import datetime
import gspread
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from gspread_formatting import DataValidationRule, BooleanCondition, set_data_validation_for_cell_range

credental_json = {
    "type": "service_account",
    "project_id": "chois-python-connect",
    "private_key_id": st.secrets["private_key_id"],
    "private_key": st.secrets["private_key"],
    "client_email": st.secrets["client_email"],
    "client_id": "116464278440047112678",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/chois-python-connect%40chois-python-connect.iam.gserviceaccount.com",
}

# Google Sheets 및 Drive 서비스 설정
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = service_account.Credentials.from_service_account_info(credental_json, scopes=scope)
client = gspread.authorize(credentials)
service = build('drive', 'v3', credentials=credentials)


using_spreadsheet_id = '1DwMKa9x9mHZnKUFgylhgQahEoFaTmfHCr4yeCVNVpT4'
spreadsheet = client.open_by_key(using_spreadsheet_id)
chehumsheet = spreadsheet.worksheet('2024dgmath')

def get_seoul_time():
    # 서울 시간대 객체 생성
    seoul_tz = pytz.timezone('Asia/Seoul')

    # 현재 UTC 시간을 가져온 후 서울 시간대로 변환
    utc_dt = datetime.utcnow()
    utc_dt = utc_dt.replace(tzinfo=pytz.utc)  # UTC 시간대 정보 추가
    seoul_dt = utc_dt.astimezone(seoul_tz)  # 서울 시간대로 변환

    # strftime을 사용하여 원하는 형식의 문자열로 변환
    formatted_time = seoul_dt.strftime('%Y-%m-%d %H:%M:%S')
    return formatted_time

def data_load(sheet_name):
    all_data = pd.DataFrame(sheet_name.get_all_values())
    all_data.columns = list(all_data.iloc[0])
    all_data = all_data.iloc[1:]
    return all_data

def insert_row(sheet, input_list, last_row, success_message):
    validation_rule = DataValidationRule(
        BooleanCondition('BOOLEAN', ['TRUE', 'FALSE']),
        # condition'type' and 'values', defaulting to TRUE/FALSE
        showCustomUi=True)
    set_data_validation_for_cell_range(sheet, "A" + str(last_row), validation_rule)  # inserting checkbox
    input_list.append(get_seoul_time())
    sheet.update(range_name="B" + str(last_row), values=[input_list])
    sheet.update_cell(last_row, 11, value="=\"Daugu-2024-\"&text(J" + str(last_row) + ",\"00#\")")
    st.success(success_message)

def data_input_chehum(sheet, input_list, success_message):
    last_row = sheet.row_count + 1
    sheet.add_rows(1)
    insert_row(sheet, input_list, last_row, success_message)





def get_pdfdata_from_sheet(using_spreadsheet_id, gid):
    """Google Sheets에서 PDF 데이터를 가져오는 함수"""
    url = f"https://docs.google.com/spreadsheets/d/{using_spreadsheet_id}/export?exportFormat=pdf&gid={gid}"
    response = service._http.request(url)
    if response[0].status != 200:
        raise Exception(f"Failed to fetch PDF: {response[0].status}")
    pdf_data = response[1]
    return pdf_data

def send_mail(pdf_data, fromaddr, toaddr, school_name, name):
    """이메일 발송 함수"""
    if not isinstance(toaddr, str):
        raise ValueError("The email address must be a string.")

    msg = MIMEMultipart('alternative')
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = Header(f"{school_name} {name}학생의 대구수학페스티벌 이수증", 'utf-8')

    part = MIMEBase('application', "octet-stream")
    part.set_payload(pdf_data)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=f"{school_name}_{name}.pdf")
    msg.attach(part)

    # 이메일 발송
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(fromaddr, st.secrets["email_secret"])
    server.sendmail(fromaddr, toaddr, msg.as_string())
    server.quit()

def send_mail2(pdf_content, email, school, name):
    fromaddr = "complete860127@gmail.com"
    password = st.secrets["email_secret"]
    toaddr = email

    # instance of MIMEMultipart
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = "2024학년도 대구수학페스티벌 참가확인서"

    # attachment
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(pdf_content)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=f"{school}_{name}.pdf")
    msg.attach(part)

    # attach the instance 'part' to instance 'msg'
    msg.attach(part)

    # creates SMTP session
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(fromaddr, password)
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    server.quit()
