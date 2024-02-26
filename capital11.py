# python3.9
# selenium4.4.0
# pyinstaller5.3

import datetime
import win32com.client
from asset_func import risona, manex_value

apl = win32com.client.Dispatch("Excel.Application")
apl.Visible = 1
wb1 = apl.Workbooks.open(r'Z:\文書\総資産.xlsx')
ws = wb1.Worksheets("総資産")
wb2 = apl.Workbooks.open(r'Z:\文書\資産運用成績.xlsm')
ws1 = wb2.Worksheets("利率")

today = datetime.date.today()

now = ws.Range('A2').value
now_ex = datetime.date(now.year, now.month, now.day)

if today != now_ex:
    ws.Range("2:2").Insert()
    ws.Range("B2:K2").style = "通貨"
    ws.Range("A2:O2").HorizontalAlignment = -4152
    ws.Range("A2").value = today.strftime("%Y/%m/%d")
    ws.Range('A2').NumberFormat = ws.Range('A3').NumberFormat
    ws.Range('H2').value = int(0)
    ws.Range('I2').value = int(6342000)
    ws.Range('J2').formula = "=SUM(B2:I2)"
    ws.Range('J2').Interior.color = ws.Range('J3').Interior.color
    ws.Range('K2').formula = "=J2"
    ws.Range('K3').formula = "=K2"
    ws.Range('L2').formula = "=J2-J3"
    ws.Range('M2').value = int(68299343)
    ws.Range('N2').formula = "=N3-13333.3"
    ws.Range('O2').formula = "=(J2-N2)/13333.3"

ws.Range('B2').value = ws1.Range("C2").value

# try:
#     ws.Range('C2').value, ws.Range('D2').value = manex_value()
# except Exception:
#     pass

ws.Range('C2').value, ws.Range('D2').value = manex_value()



try:
    ws.Range('E2').value, ws.Range('F2').value = risona()
except Exception:
    pass
# try:
#     ws.Range('G2').value = roukin()
# except Exception:
#     pass

ws.Range('A1').value = datetime.datetime.now().strftime('%Y/%m/%d %H:%M')

wb1.Save()
