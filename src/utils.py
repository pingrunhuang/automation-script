from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from contextlib import contextmanager
from typing import Dict
from binance.client import BaseClient
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
    print("Calling vba......")
    macro = wb.macro("Sheet9.BinanceRate")
    try:
        macro()
    except pywintypes.com_error:
        print("some error happen when calling vba")
        return
    finally:
        print("Done vba")

if CONFIGS["TESTNET"] is True:
    color_print("Running on testnet", "green")



class MyBNCClient(Client):
    def create_order(self, **params):
        print(f"Creating sell order with parameters: {params}")
        uri = self._create_api_uri("order", True, BaseClient.PUBLIC_API_VERSION)
        print(f"Endpoing: {uri}")
        self.sell_params = params
        return super().create_order(**params)
    def generate_reject_email(self):
        """
        - Symbol:	XXX (take from "symbol" crop out ending USDT)
        - Profit:		YYY.YY (take from “cummulativeQuoteQty")
        - Expd Profit:	YYY.YY (take from columnMcell)
        - line break
        - posted api
                        - line break
        - full details of 5.2
        """
        pass


    def generate_sell_email(self, resp:dict):
        """
        ) send email with subject: “Crypto-Binance-SellDone”, Body:
        - Symbol:	XXX (take from "symbol" crop out ending USDT)
        - Profit:		YYY.YY (take from “cummulativeQuoteQty")
        - Expd Profit:	YYY.YY (take from columnMcell)
        - Qty:		-ZZZ.ZZ (take from "executedQty")
        - line break
        - posted api
        - line break
        - full details of 5.2.3.2
        """
        lines = [
            f"Symbol: {resp["symbol"]}, Profit: {resp["cummulativeQuoteQty"]}, Expd Profit: {resp["columnMcell"]}, Qty: {resp["executedQty"]}",
            f"API endpoint: {self._create_api_uri("order", True, BaseClient.PUBLIC_API_VERSION)}"
            f"Parameters: {self.sell_params}",
            f"Responses: {resp}"
        ]
        return "\n".join(lines)

    def generate_sell_error_mail(self):
        """
        5.2.3.2.1.1) send email with subject: “Crypto-Binance-SellOrderError”, Body:
                        - posted api
        - line break
        - full details of 5.2.3.2
        """

    def generate_balance_email(self):
        """
        - Total-USD:		XXX.XX (cell value from AQ2)
        - Total-AED:		XXX.XX (cell value from AS2)
        - P/L%:			YY.Y (cell value from AQ4)
        - AJ-USDT:		XXX.XX (cell value from AN23)
        - Binance-USDT:	XXX.XX (USDT balance received from Binance)
        - line break
        - posted api
        - line break
        - full details of 7.2
        """
        pass


CLIENT = MyBNCClient(CONFIGS["API_KEY"], CONFIGS["SECRET_KEY"], testnet=CONFIGS["TESTNET"])


