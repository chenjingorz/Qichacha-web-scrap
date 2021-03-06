# names with 100% on qichacha are stored in SQL, else stored in csv 'unmatched entities' for manual search

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import pyodbc
from time import sleep

# import the excel file with entities
entity = pd.read_excel('C:\Work Data (Keppel ABC Audit)\PYTHON PROJECT\QICHACHA\Copy of entity.xlsx')

connection = pyodbc.connect('Driver={SQL Server Native Client 11.0};Server=SGKCLNB18030528;Database=Entity1;Trusted_Connection=yes;')
cursor = connection.cursor()

# ONLY TO CLEAR ALL ROWS IN THE TABLE
cursor.execute("truncate table MainMembers")
cursor.execute("truncate table StockHolders")
cursor.execute("truncate table EntityInfo")
connection.commit()

# run in BROWSER MODE if u want to see whats happening
driver = webdriver.Chrome('C:/Program Files/Java/chromedriver.exe')
driver.get('https://www.qichacha.com/')

# run in headless mode for faster execution, need to download google canary
#chrome_options = Options()
#chrome_options.add_argument("--headless")
#chrome_options.add_argument("--window-size=1920x1080")
#driver = webdriver.Chrome('C:/Program Files/Java/chromedriver.exe', chrome_options=chrome_options)
#driver.get('https://www.qichacha.com/')

# open with encoding so that when u write no need encoding
f = open('Unmatched entities2.csv','w', encoding = 'utf-8-sig')
f.write("Unmatched Entities|Search1|Search2|Search3|\n")

# login first
login = driver.find_element_by_xpath('/html/body/header/div/ul[2]/li[7]/a').click()
qq = driver.find_element_by_xpath('//*[@id="qrcodeLoginPanel"]/div[2]/div/div[3]/a[2]').click()
# switch to the different iframe embedded in the login page
driver.switch_to.frame(0)
qq = driver.find_element_by_xpath('//*[@id="switcher_plogin"]').click()
user = driver.find_element_by_xpath('//*[@id="u"]').send_keys('953726816')
pw = driver.find_element_by_xpath('//*[@id="p"]').send_keys('KepCorp@123')
login = driver.find_element_by_xpath('//*[@id="login_button"]').click()

sleep(5)

# for i in range(entity.shape[0]):
for i in range(326, entity.shape[0]):
    if i != 0:
        newSearch = driver.find_element_by_xpath('//*[@id="headerKey"]')
        for k in range(len(key)):
            newSearch.send_keys(Keys.BACKSPACE)
        key = entity.iloc[i]['BUTXT']
        newSearch.send_keys(key)
        select = driver.find_element_by_xpath('/html/body/header/div/form/div/div/span/button').click()
    else:
        key = entity.iloc[0]['BUTXT']
        search = driver.find_element_by_xpath('//*[@id="searchkey"]').send_keys(key)
        search = driver.find_element_by_xpath('//*[@id="V3_Search_bt"]').click()

    try:
        name = driver.find_elements_by_class_name('ma_h1')
        if len(name) == 0:
            f.write(key + "\n")
        else:
            company = name[0].text
            if company == key:
                companyInfo = driver.find_elements_by_class_name('m-t-xs')
                date = companyInfo[0].text.split('成立时间：')[1]
                email = companyInfo[1].text.split(' ')[0][3:]
                phone = companyInfo[1].text.split(' ')[1][3:]
                address = companyInfo[2].text[3:]
                command = "insert into EntityInfo(CompanyName, FoundedOn, Phone, Email, Address) values (N'" + key + "', '" + date + "', '" + phone + "' , '" + email + "', N'" + address + "')"
                cursor.execute(command)
                name[0].click()
                driver.switch_to.window(driver.window_handles[1])

                # legalRep
                try:
                    comInfo = driver.find_element_by_id("Cominfo")
                    legalRep = comInfo.find_element_by_class_name('bname').text
                    command = "update EntityInfo set LegalRep = N'" + legalRep + "' where CompanyName = N'" + key + "'"
                    cursor.execute(command)
                except NoSuchElementException:
                    command = "update EntityInfo set LegalRep = '-' where CompanyName = N'" + key + "'"
                    cursor.execute(command)

                # use assigned variableName.find instead of driver.find to find the element under that variable
                try:
                    mainMember = driver.find_element_by_id('Mainmember')
                    # name and position list
                    mainName = mainMember.find_elements_by_xpath(
                        './/a[@class = "c_a"]')  # .// current processing starts at the current node
                    mainPosition = mainMember.find_elements_by_xpath('.//td[@class = "text-center"]')
                    memberNo = mainMember.find_element_by_xpath('//*[@id="Mainmember"]/div/span[1]').text
                    for j in range(int(memberNo)):
                        # store into SQL
                        command = "insert into MainMembers(CompanyName, MainMember, Position) values (N'" + key + "', N'" + \
                                  mainName[j].text + "', N'" + mainPosition[j].text + "')"
                        cursor.execute(command)
                except NoSuchElementException:
                    pass

                try:
                    stockInfo = driver.find_element_by_id('Sockinfo')
                    # stockholder list
                    stockHolders = stockInfo.find_elements_by_xpath('.//a[not(@class ="btn-touzi")]')
                    stockNo = stockInfo.find_element_by_xpath('//*[@id="Sockinfo"]/div/span[1]').text
                    for j in range(int(stockNo)):
                        command = "insert into StockHolders(CompanyName, StockHolder) values (N'" + key + "', N'" + \
                                  stockHolders[j + 1].text + "')"
                        cursor.execute(command)
                except NoSuchElementException:
                    pass
                connection.commit()

                # close current window and switch to original page
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
            else:
                if (len(name) < 3):
                    write = key
                    for i in range(len(name)):
                        write += '|' + name[i].text
                    f.write(write + '\n')
                else:
                    f.write(key + '|' + name[0].text + '|' + name[1].text + '|' + name[2].text + '\n')
    except NoSuchElementException:
        f.write(key + '\n')

f.close()
driver.close()
