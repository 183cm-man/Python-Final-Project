import smtplib
import imaplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openpyxl
from openpyxl import Workbook
from datetime import datetime
import random
import string
import time
import concurrent.futures
import pandas as pd
import os

RECEIVE_TIME_LIMIT = 60*3
# 配置
SMTP_USER = "icplfinalproject@gmail.com"
SMTP_PASSWORD = "suyp ldja vkvr rluf"
SMTP_SERVER = "smtp.gmail.com"
IMAP_SERVER = "imap.gmail.com"

# 已知學生列表
file_path = "known_students.xlsx"
if not os.path.exists(file_path):
    raise FileNotFoundError(f"{file_path} not found. Please check the file path.")
known_students_df = pd.read_excel(file_path)
known_students = list(known_students_df["Email"])


# 定義初始化Excel文件的函數
def init_excel(file_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Attendance"
    sheet.append(["Email", "Number", "Date", "Status"])
    workbook.save(file_path)

# 檢查並初始化出席記錄文件
attendance_file = "Attendance.xlsx"
if not os.path.exists(attendance_file):
    init_excel(attendance_file)


# 生成隨機6位包含數字和大小寫字母的密碼
def generate_password(length=6):
    characters = string.ascii_letters + string.digits
    return ''.join(random.choice(characters) for _ in range(length))

# 發送點名郵件的函數
def send_attendance_email(student_email, course_name):
    subject = f"Attendance for {course_name}"
    body = "Please reply to this email with the attendance code.\n" \
           "Please respond within 3 minutes. The correct code is case-sensitive."

    msg = MIMEMultipart()
    msg["From"] = SMTP_USER
    msg["To"] = student_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    server = smtplib.SMTP_SSL(SMTP_SERVER, 465)
    server.login(SMTP_USER, SMTP_PASSWORD)
    server.sendmail(SMTP_USER, student_email, msg.as_string())
    server.quit()

# 發送點名成功的郵件
def send_success_email(student_email, status):
    subject = "Attendance Confirmation"
    body = f"Your attendance code is correct. You are marked as {status}."

    msg = MIMEMultipart()
    msg["From"] = SMTP_USER
    msg["To"] = student_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    server = smtplib.SMTP_SSL(SMTP_SERVER, 465)
    server.login(SMTP_USER, SMTP_PASSWORD)
    server.sendmail(SMTP_USER, student_email, msg.as_string())
    server.quit()

# 發送錯誤密碼的郵件
def send_error_email(student_email, is_final=False):
    if is_final:
        subject = "Final Attendance Code Incorrect"
        body = "The attendance code you provided is incorrect again. You are marked as absent."
    else:
        subject = "Attendance Code Incorrect"
        body = "The attendance code you provided is incorrect. Please try again with the correct code."

    msg = MIMEMultipart()
    msg["From"] = SMTP_USER
    msg["To"] = student_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    server = smtplib.SMTP_SSL(SMTP_SERVER, 465)
    server.login(SMTP_USER, SMTP_PASSWORD)
    server.sendmail(SMTP_USER, student_email, msg.as_string())
    server.quit()

# 發送缺席的郵件
def send_absent_email(student_email):
    subject = "Attendance Confirmation"
    body = "You did not reply to the attendance code on time. You are marked as absent."

    msg = MIMEMultipart()
    msg["From"] = SMTP_USER
    msg["To"] = student_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    server = smtplib.SMTP_SSL(SMTP_SERVER, 465)
    server.login(SMTP_USER, SMTP_PASSWORD)
    server.sendmail(SMTP_USER, student_email, msg.as_string())
    server.quit()

# 檢查並處理郵件回覆
def check_email():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(SMTP_USER, SMTP_PASSWORD)
    mail.select("inbox")

    status, data = mail.search(None, "UNSEEN")
    mail_ids = data[0]

    id_list = mail_ids.split()
    for i in id_list:
        status, data = mail.fetch(i, "(RFC822)")
        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                email_from = email.utils.parseaddr(msg["from"])[1]
                email_subject = msg["subject"]

                # 處理郵件內容
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            email_body = part.get_payload(decode=True).decode("utf-8").strip()
                            first_line = email_body.split('\n', 1)[0].strip()  # 只取第一行
                            process_attendance_response(email_from, first_line)
                            break
                else:
                    email_body = msg.get_payload(decode=True).decode("utf-8").strip()
                    first_line = email_body.split('\n', 1)[0].strip()  # 只取第一行
                    process_attendance_response(email_from, first_line)

# 處理點名回覆
def process_attendance_response(email_from, email_body):
    print(f"收到來自 {email_from} 的郵件，內容為：{email_body}")  # 調試用日誌

    if email_from in known_students:
        if email_body in valid_codes:
            print(f"密碼正確：{email_body}")  # 調試用日誌
            if email_from in incorrect_attempts_once:
                record_attendance(email_from, "Late")
                status = "Late"
                LateList.add(email_from)
                incorrect_attempts_once.remove(email_from)
                
            else:
                record_attendance(email_from, "Present")
                status = "Present"
                used_codes.add(email_body)
                PresentList.add(email_from)

            send_success_email(email_from, status)
            valid_codes.remove(email_body)
            responded_students.add(email_from)
           
        else:
            
            if email_body in used_codes : 
                print(f"密碼已使用過：{email_body} ")  
            else :
                print(f"密碼錯誤：{email_body} ")

            if email_from in incorrect_attempts_once :
                record_attendance(email_from, "Absent")
                send_error_email(email_from, is_final=True)
                incorrect_attempts_twice.add(email_from)
                incorrect_attempts_once.remove(email_from)
            else :
                send_error_email(email_from)
                incorrect_attempts_once.add(email_from)

            responded_students.add(email_from)
    else:
        print(f"未知的郵件地址：{email_from}")
        record_attendance(email_from, "Invalid Email")



def record_attendance(email, status):
    file_path = attendance_file
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    if email in known_students:
        row_number = known_students.index(email)
        name = known_students_df.iloc[row_number]["Number"]
    else:
        name = "Unknown"
    sheet.append([email, name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), status])
    workbook.save(file_path)

# 初始化 Excel 文件
init_excel("Attendance.xlsx")

# 計時任務
def timed_check_email():
    start_time = time.time()
    while time.time() - start_time < RECEIVE_TIME_LIMIT:
        check_email()
        time.sleep(2)  # 每2秒檢查一次郵件

# 初始化必需的集合
valid_codes = set(generate_password() for _ in range(len(known_students)))
print(f"生成的點名密碼：{valid_codes}")  # 調試用日誌
used_codes = set()
incorrect_attempts_once = set()
incorrect_attempts_twice = set()
responded_students = set()


PresentList = set()
LateList = set()
noResponsedList = set()

# 並行發送初始點名郵件
with concurrent.futures.ThreadPoolExecutor() as executor:
    futures = [executor.submit(send_attendance_email, student_email, "Python Class") for student_email in known_students]
    concurrent.futures.wait(futures)

# 定期檢查郵件
timed_check_email()


# 記錄未回覆正確密碼的學生為缺席
for student_email in known_students:
    if student_email not in responded_students:
        record_attendance(student_email, "Absent")
        send_absent_email(student_email)
        noResponsedList.add(student_email)
    elif student_email in incorrect_attempts_once :
        record_attendance(student_email, "Absent")
        send_absent_email(student_email)
        


print ("-------點名結束-------")
print ("出席：" + str(list(PresentList)))
print ("遲到：" + str(list(LateList)))
print ("缺席(兩次密碼錯誤）：" + str(list(incorrect_attempts_twice)))
print ("缺席（輸入過一次錯誤密碼而未回覆第二次）：" + str(list(incorrect_attempts_once)))
print ("缺席（從未回覆）：" + str(list(noResponsedList)))
print ("出席結果已儲存在 Attendance.xlsx ")

