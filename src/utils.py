import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from contextlib import contextmanager
from colorama import Fore, Style
from requests.sessions import Session as Session
import yaml
from binance import Client
import win32com.client
import xlwings as xw
import pywintypes


with open("config.yaml") as f:
    CONFIGS = yaml.safe_load(f)

# excel_path = "D:\Google Drive\AJ\Excel"
excel_path = CONFIGS["EXCEL_PATH"]


@contextmanager
def read_and_save_workbook():
    workbook = xw.Book(excel_path)
    yield workbook
    workbook.save()


def color_print(msg, color=None):
    if color=="green":
        print(Fore.GREEN+msg)
    elif color=="red":
        print(Fore.RED+msg)
    else:
        print(Style.RESET_ALL+msg)

def send_email(subject:str, body:str):
    # Create the email message
    password = CONFIGS["EMAIL_PASS"]
    from_email = CONFIGS["EMAIL_FROM"]
    to_emails = CONFIGS["EMAIL_TO"]
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = ', '.join(to_emails)
    msg['Subject'] = subject

    # Attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))

    # Set up the server and send the email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)  # Specify your SMTP server and port
        server.starttls()  # Secure the connection
        server.login(from_email, password)
        text = msg.as_string()
        server.sendmail(from_email, to_emails, text)
        server.quit()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {str(e)}")

def format_usd_ntl(ntl:str)->float:
    # return float(ntl.replace("US$", ""))
    return float(ntl)

def format_currency(ntl:float)->str:
    return "US${:10,.8f}".format(ntl)


def call_vb(wb:xw.Book):
    print("calling vba......")
    macro = wb.macro("Sheet9.BinanceRate")
    try:
        macro()
    except pywintypes.com_error:
        print("some error happen when calling vba")
        return

if CONFIGS["TESTNET"] is True:
    color_print("Running on testnet", "green")

CLIENT = Client(CONFIGS["API_KEY"], CONFIGS["SECRET_KEY"], testnet=CONFIGS["TESTNET"])