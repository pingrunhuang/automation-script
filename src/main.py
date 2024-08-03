from utils import read_and_save_workbook, CLIENT, send_email
import buy_module, sell_module, reset_module
import time


if __name__ == "__main__":

    with read_and_save_workbook() as workbook:
        start = time.time()
        sell_module.run(workbook)
        # reset_module.run(workbook)
        # buy_module.run(workbook)
        sheet = workbook.sheets["Overview"]
        total_usd = sheet["AQ2"].value
        total_aed = sheet["AS2"].value
        pl_pct = sheet["AQ4"].value
        aj_usdt = sheet["AN23"].value
        body = CLIENT.generate_balance_email(int(total_usd), int(total_aed), pl_pct, aj_usdt, start)
        send_email("Crypto-Binance-End", body)
        print("Finished all modules")