import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By


def manex_login():
    url03 = "https://mst.monex.co.jp/pc/ITS/login/LoginIDPassword.jsp"
    # options = Options()
    # options.add_argument('--headless')
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    # driver = Chrome(options=options)
    # driver = Chrome(r"c:\webdriver\chromedriver.exe")

    driver.implicitly_wait(10)
    driver.get(url03)

    tag_v1 = driver.find_element(By.NAME, "loginid")
    tag_v1.send_keys("92007709")

    tag_v2 = driver.find_element(By.NAME, "passwd")
    tag_v2.send_keys("ABC123abc")

    tag_v3 = driver.find_element(By.CSS_SELECTOR, "#contents > div > p.mb-20 > input")
    tag_v3.submit()

    return driver


def manex_value():
    driver = manex_login()
    time.sleep(10)
    pr04 = int(driver.find_element(By.CSS_SELECTOR, "#buyingpower").text.replace(",", ""))

    # tag_v5 = driver.find_element(By.XPATH, "/html/body/div[3]/div[1]/div/div[2]/div/ul/li[2]/a")
    # tag_v5=driver.find_element(By.CSS_SELECTOR,"#gnav > li.nav02.gnav-item-asset.gnav-item-reporthelp.gn_custAsset.current.is-current > a")
    tag_v5 = driver.find_element(By.LINK_TEXT, "保有残高・口座管理")
    tag_v5.click()

    tag_v6 = driver.find_element(By.LINK_TEXT, "iDeCo残高")
    # tag_v6=driver.find_element(By.CSS_SELECTOR,"#gn_custAsset-lm_custAsset > div.contents > div > div.nav-cmn-tab_04.type-size-l > ul > li:nth-child(7) > a")
    tag_v6.click()
    time.sleep(10)
    pr05 = int(
        driver.find_element(By.CSS_SELECTOR, "#pensionAssetValuation").text.replace(",",
                                                                                    ""))
    # tag_v4 = driver.find_element(By.XPATH, "/html/body/div[3]/div[1]/div/div[1]/div[2]/p/a")
    tag_v4 = driver.find_element(By.LINK_TEXT, "ログアウト")
    tag_v4.click()
    driver.quit()
    return pr04, pr05


def roukin():
    url02 = "https://www.parasol.anser.ne.jp/ib/index.do?PT=BS&CCT0080=2963"
    # driver = Chrome(r"c:\anaconda3\webdriver\chromedriver.exe", options=options)
    # driver = Chrome(r"c:\webdriver\chromedriver.exe")
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

    driver.implicitly_wait(10)
    driver.get(url02)

    # tag_r0 = driver.find_element_by_id("pswd001_check")
    # tag_r0.click()

    # tag_r00=driver.find_element_by_class_name("ui-button-text")
    # tag_r00.click()

    tag_r1 = driver.find_element(By.ID, "txtBox001")
    tag_r1.send_keys("0002887722")

    tag_r2 = driver.find_element(By.ID, "pswd001")
    tag_r2.send_keys("ABC123abc")

    tag_r3 = driver.find_element(By.ID, "btn999")
    tag_r3.click()

    num01 = int(driver.find_element(By.ID, "msg002-1").text) - 1
    num02 = int(driver.find_element(By.ID, "msg002-2").text) - 1

    num = str(4610497951)

    input_num = num[num01] + num[num02]

    tag_r5 = driver.find_element(By.ID, "pswd001")
    tag_r5.send_keys(input_num)

    tag_r6 = driver.find_element(By.ID, "btn002")
    tag_r6.click()

    try:
        pr03 = driver.find_element(By.ID, "msg107-1").text.replace(",", "").replace("円", "")
        pr03 = int(pr03)
    except Exception:
        tag_r7 = driver.find_element(By.ID, "cs_globalButton_logout")
        tag_r7.click()

        pr03 = driver.find_element(By.ID, "msg107-1").text.replace(",", "").replace("円", "")
        pr03 = int(pr03)

    tag_r7 = driver.find_element(By.ID, "cs_globalButton_logout")
    tag_r7.click()

    driver.quit()
    return pr03


def risona():
    url01 = "https://ib.saitamaresona.co.jp/IB/0102/SC_N_0102_010.aspx"
    # driver = Chrome(service=service, options=options)
    # driver = Chrome(r"c:\anaconda3\webdriver\chromedriver.exe", options=options)
    # driver = Chrome(r"c:\webdriver\chromedriver.exe")
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

    # driver = Chrome(options=options)
    driver.implicitly_wait(10)
    driver.get(url01)

    tag_t0 = driver.find_element(By.XPATH, "/html/body/div[1]/form/div[5]/div/span/table[1]/tbody/tr/td[2]/input")
    tag_t0.click()

    tag_t1 = driver.find_element(By.XPATH, "/html/body/div[1]/form/div[5]/div/span/table[1]/tbody/tr/td[1]/p/input")
    tag_t1.send_keys("GHG01004")

    tag_t2 = driver.find_element(By.XPATH, "/html/body/div[1]/form/div[5]/div/span/ul[2]/li[2]/input")
    tag_t2.click()

    tag_t3 = driver.find_element(By.XPATH, "/html/body/div[1]/form/div[5]/div/div[2]/table/tbody/tr[3]/td[2]/input")
    tag_t3.click()

    tag_t4 = driver.find_element(By.XPATH, "/html/body/div[1]/form/div[5]/div/div[2]/table/tbody/tr[3]/td[1]/p/input")
    tag_t4.send_keys("ABC123abc")

    tag_t5 = driver.find_element(By.XPATH, "/html/body/div[1]/form/div[5]/div/ul[2]/li/input")
    tag_t5.click()

    # tag_t7 = driver.find_element(By.XPATH, "/html/body/div[1]/form/div[5]/div/ul[2]/div/li/input")
    # tag_t7.click()

    tag_t6 = driver.find_element(By.XPATH, "/html/body/div[1]/form/div[4]/div[2]/ul[1]/li[2]/ul/li[1]/input")
    tag_t6.click()

    pr01 = driver.find_element(By.XPATH,
                               "/html/body/div[1]/form/div[4]/div[1]/div[3]/div/table/tbody/tr[2]/td[5]").text.replace(
        ",", "").replace("円",
                         "")
    pr01 = int(pr01)

    tag_t8 = driver.find_element(By.XPATH, "/html/body/div[1]/form/div[4]/div[2]/ul[1]/li[3]/ul/li[4]/input")
    tag_t8.click()

    pr02 = driver.find_element(By.XPATH,
                               "/html/body/div[1]/form/div[4]/div[1]/div[6]/div/div/table/tbody/tr[2]/td[5]").text.replace(
        ",", "").replace(
        "円",
        "")
    pr02 = int(pr02)

    tag_t9 = driver.find_element(By.XPATH, "/html/body/div[1]/form/div[3]/ul/div/li[1]/input")
    tag_t9.click()
    driver.quit()
    return pr01, pr02


def comma(a):
    if "\n" in a:
        c = a.rsplit("\n")
        if len(c) == 3:
            b_0, b_1, b_2 = a.rsplit("\n")
            if b_0 != '－':
                return float(b_0.replace(",", ""))
            else:
                return float(b_1.replace(",", ""))
        else:
            b_0, b_1 = a.rsplit("\n")
            return float(b_0.replace(",", ""))
