# import win32clipboard as w
import time
import xlrd
# import pyperclip
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
br = webdriver.Firefox()

ws = xlrd.open_workbook('dataconfirm.xls')
tab = ws.sheets()[0]
nrows = tab.nrows
data_rows = []
for i in range(nrows):
    data_rows.append(tab.row_values(i))

for row in data_rows:
    try:
        print(row[0])
        datas = [row[0],row[1],row[2],'2017',row[3],'234300']
        br.get("http://60.173.222.146/firstWeb/szksLogin.jsp?loginType=01")
        time.sleep(1)
        br.find_element_by_id('grq_szksdl_fxm').send_keys(row[0])
        time.sleep(0.5)
        br.find_element_by_id('grq_szksdl_fsfzh').send_keys(row[1])
        time.sleep(0.5)
        br.find_element_by_id('grq_szksdl_fdhhm').send_keys(row[2])
        time.sleep(1)

        vtxt = input('captcha:')
        br.find_element_by_id('jcaptcha').send_keys(vtxt)

        br.find_element_by_id('login-btn').click()
        time.sleep(2)
        br.find_element_by_partial_link_text('报考中职学校').click()
        
        c = input(':')
        if c == 'q':
            break
    except:
        print('error')
br.quit()
