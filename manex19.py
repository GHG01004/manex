# python3.9
# selenium4.10.0
# pyinstaller5.6.2

import datetime

import win32com.client
from selenium.webdriver.common.by import By

from asset_func import comma
from asset_func import manex_login

# options = Options() options.add_argument('--headless')  # headlessモードを使用する options.add_argument(
# f'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.79 '
# f'Safari/537.36')


# url = "https://mst.monex.co.jp/pc/ITS/login/LoginIDPassword.jsp"
# driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
# driver.implicitly_wait(10)
# driver.get(url)
# tag_t1 = driver.find_element(By.NAME, "loginid")
# tag_t1.send_keys("92007709")
# tag_t2 = driver.find_element(By.NAME, "passwd")
# tag_t2.send_keys("ABC123abc")
# tag_t3=driver.find_element(By.CSS_SELECTOR,"#contents > div > p.mb-20 > input")
# tag_t3.submit()

driver = manex_login()

tag_t4 = driver.find_elements(By.CLASS_NAME, "nav02")
tag_t4[0].click()

# tag_t5=driver.find_element(By.CSS_SELECTOR,"#contents > table:nth-child(8) > tbody")
tag_t5 = driver.find_elements(By.CSS_SELECTOR, ".table-block.table-cmn_01.amountList.s-mt-10")

tag_t6 = tag_t5[0].find_elements(By.TAG_NAME, "tr")
del tag_t6[0]
list01 = []
for i in tag_t6:
    i.find_elements(By.TAG_NAME, "td")
    list01.append(i.text.split(" "))
dict1 = {i[0].split("\n")[0]: comma(i[3]) for i in list01}

# tag_t7 = driver.find_element(By.CSS_SELECTOR,"#contents > table:nth-child(16) > tbody")
# tag_t7=driver.find_elements(By.CSS_SELECTOR,"table-block table-cmn_01 amountList s-mt-10")
tag_t8 = tag_t5[1].find_elements(By.TAG_NAME, "tr")
del tag_t8[0]
list02 = []
for i in tag_t8:
    i.find_elements(By.TAG_NAME, "td")
    list02.append(i.text.split(" "))
dict2 = {i[0].split("\n")[0]: comma(i[2]) for i in list02}

apl = win32com.client.Dispatch("Excel.Application")
apl.Visible = 1

wb1 = apl.Workbooks.open(r'Z:\文書\新資産運用【日々】.xlsx')
ws11 = wb1.Worksheets("2021415")
ws12 = wb1.Worksheets("マネックス")
ws13 = wb1.Worksheets("乖離率")

dict1.update(dict2)

c = int(1)
for a, b in dict1.items():
    d = str("A") + str(c)
    e = str("B") + str(c)
    ws11.Range(d).value = a
    ws11.Range(e).value = b
    c = c + 1

wb2 = apl.Workbooks.open(r'Z:\文書\資産運用成績.xlsm')
ws21 = wb2.Worksheets("利率")

apl.Run('資産運用成績.xlsm!銘柄利率')
apl.Run('資産運用成績.xlsm!グラフ記入')

wb3 = apl.Workbooks.open(r'Z:\文書\総資産.xlsx')
ws31 = wb3.Worksheets("総資産")

today = datetime.date.today()
now = ws31.Range('A2').value
now_ex = datetime.date(now.year, now.month, now.day)
now1 = ws13.Range("A2").value
now_ex1 = datetime.date(now1.year, now1.month, now1.day)

if today == now_ex:
    ws31.Range('B2').value = ws21.Range("C2")

if today != now_ex1:
    ws13.Range("A2").EntireRow.Insert(-4121)
    ws13.Range("A2").value = today.strftime("%Y/%m/%d")
    ws13.Range("B2:J2").NumberFormat = "0.00%"
    ws13.Range("B2:J2").HorizontalAlignment = -4152

ws13.Range("B2").value = ws12.Range("AD41").value
ws13.Range("C2").value = ws12.Range("AD55").value
ws13.Range("D2").value = ws12.Range("AD56").value
ws13.Range("E2").value = ws12.Range("AD43").value
ws13.Range("F2").value = ws12.Range("AD44").value
ws13.Range("G2").value = ws12.Range("AD46").value
ws13.Range("H2").value = ws12.Range("AD47").value
ws13.Range("I2").value = ws12.Range("AD48").value
ws13.Range("J2").value = ws12.Range("AD49").value

wb1.Save()
wb2.Save()
wb3.Save()

driver.quit()
