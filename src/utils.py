import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from openpyxl import load_workbook
from contextlib import contextmanager
from colorama import Fore, Style
import yaml
from binance import Client

# excel_path = "D:\Google Drive\AJ\Excel"
excel_path = "/Users/lei/code/python-scripts/crypto.xlsm"

@contextmanager
def read_sheet(sheetname, is_save=True):
    workbook = load_workbook(excel_path)
    yield workbook[sheetname]
    if is_save:
        workbook.save(excel_path)

def color_print(msg, color=None):
    if color=="green":
        print(Fore.GREEN+msg)
    elif color=="red":
        print(Fore.RED+msg)
    else:
        print(Style.RESET_ALL+msg)

def send_email(subject, body, to_emails, from_email, password):
    # Create the email message
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

# Load the private key.
# In this example the key is expected to be stored without encryption,
# but we recommend using a strong password for improved security.
with open("config.yaml") as f:
    CONFIGS = yaml.safe_load(f)

if CONFIGS["TESTNET"] is True:
    color_print("Running on testnet", "Green")

CLIENT = Client(CONFIGS["API_KEY"], CONFIGS["SECRET_KEY"], testnet=CONFIGS["TESTNET"])