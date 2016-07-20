# import win32clipboard as w
import xlrd
import time
# import pyperclip
from selenium import webdriver
br = webdriver.Firefox()

# def set_clip(data):
#     w.OpenClipboard()
#     w.EmptyClipboard()
#     w.SetClipboardText(data)
#     w.CloseClipboard()

# def set_clip(data):
#     pyperclip.copy(data)


ws = xlrd.open_workbook('data.xls')
tab = ws.sheets()[0]
nrows = tab.nrows
data_rows = []
for i in range(nrows):
    data_rows.append(tab.row_values(i))
# print(data_rows)

# for row in data_rows:
#     print(row[0])
#     datas = [row[2],row[7],row[5],'2016',row[6],'234300',row[2],row[7],row[5],'宿州环保工程学校']
#     c = ''
#     for data in datas:
#         set_clip(data)
#         print(data,end=' ')
#         c = input(':')
#         if c == 'q':
#             break
#     if c == 'q':
#         break

for row in data_rows:
    datas = [row[2],row[7],row[5],'2016',row[6],'234300']
    br.get("http://zhk.ahzsks.cn/firstWeb/szksLogin.jsp?loginType=01")
    br.find_element_by_name('currentTZzbmb.fxm').send_keys(row[2])
    br.find_element_by_name('currentTZzbmb.fsfzh').send_keys(row[7])
    br.find_element_by_name('currentTZzbmb.fdhhm').send_keys(row[5])
    print(row[0],row[2])
    # br.find_element_by_link_text("注册").click()

    # br.find_elements_by_name('currentTZzbmb.fxm')[1].send_keys(row[2])
    # br.find_elements_by_name('currentTZzbmb.fsfzh')[1].send_keys(row[7])
    # br.find_elements_by_name('currentTZzbmb.fdhhm')[1].send_keys(row[5])
    # br.find_element_by_name('currentTZzbmb.fbynd').send_keys(2016)
    # br.find_element_by_name('currentTZzbmb.ftxdz').send_keys(row[6])
    # br.find_element_by_name('currentTZzbmb.fyzbm').send_keys('234300')

    # br.find_element_by_name('currentTZzbmb.fmmdm').send_keys('13')
    # br.find_element_by_name('currentTZzbmb.fmzdm').send_keys('01')
    # br.find_element_by_name('currentTZzbmb.fxldm').send_keys('7')
    # br.find_element_by_name('currentTZzbmb.fkslbdm').send_keys('3')
    # br.find_element_by_name('currentTZzbmb.fdsdm').send_keys('22')
    # br.find_element_by_name('currentTZzbmb.fxqdm').send_keys('25')


    c = ''
    # for data in datas:

    #     set_clip(data)
    #     print(data,end=' ')
    #     c = input(':')
    #     if c == 'q':
    #         break
    vtxt = input('captcha:')
    br.find_element_by_name('jcaptcha').send_keys(vtxt)
    br.find_element_by_id('login-btn').click()

    # br.find_element_by_link_text('报考中职学校').click()

    for i in range(50):
        flag = True
        try:
            br.find_element_by_link_text('报考中职学校').click()
        except:
            flag = False
            print('sleep')
        time.sleep(2)
        if flag:
            break



    for i in range(50):
        flag = True
        try:
            br.find_element_by_partial_link_text('确认').click()
        except:
            flag = False
            print('sleep')
        time.sleep(2)
        if flag:
            break

    for i in range(50):
        flag = True
        try:
            br.find_element_by_partial_link_text('确定').click()
        except:
            flag = False
            print('sleep')
            time.sleep(1)
            br.find_element_by_partial_link_text('确认').click()
        time.sleep(0.5)

        if flag:
            break

    for i in range(50):
        flag = True
        try:
            br.find_element_by_partial_link_text('注销').click()
        except:
            flag = False
            print('sleep')
        time.sleep(2)
        if flag:
            break



    # time.sleep(10)
    # br.find_element_by_partial_link_text('确认').click()
    # time.sleep(10)
    # br.find_element_by_partial_link_text('确定').click()
    # # br.implicitly_wait(30)
    # time.sleep(45)
    # br.find_element_by_partial_link_text('注销').click()
    time.sleep(0.5)
    # c = input(':')
    # if c == 'q':
    #     break


br.quit()
