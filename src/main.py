import utils
import buy_module, sell_module, reset_module, min_module
import time
import xlwings as xw
import traceback


if __name__ == "__main__":
    try:
        workbook = xw.Book(utils.excel_path)
        app = workbook.app
        start = time.time()
        utils.call_vb(workbook, "AllRate")
        sheet = workbook.sheets["Stats"]
        statsF16 = sheet["F16"].value
        float(statsF16)
        sell_module.run(workbook)
        reset_module.run(workbook)
        min_module.run(workbook)
        buy_module.run(workbook)
        sheet = workbook.sheets["Stats"]
        total_usd = sheet["F16"].value
        total_aed = sheet["F17"].value
        pl_pct = sheet["F18"].value
        aj_usdt = sheet["H2"].value
        body = utils.CLIENT.generate_balance_email(int(total_usd), int(total_aed), pl_pct, aj_usdt, start)
        utils.send_email("Crypto-Binance-End", body)
        print("Finished all modules")
    except (TypeError, ValueError) as e:
        utils.send_email("Crypto-Binance-Error")
        tb = traceback.format_exc()
        print(tb)
    finally:
        workbook.save()
        workbook.close()
        app.kill()
        if not utils.PROXY:
            utils.SMTP_SERVER.close()