# import win32clipboard as w
import xlrd
import pyperclip
from selenium import webdriver
br = webdriver.Firefox()

# def set_clip(data):
#     w.OpenClipboard()
#     w.EmptyClipboard()
#     w.SetClipboardText(data)
#     w.CloseClipboard()

def set_clip(data):
    pyperclip.copy(data)


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
    br.find_element_by_link_text("注册").click()

    br.find_elements_by_name('currentTZzbmb.fxm')[1].send_keys(row[2])
    br.find_elements_by_name('currentTZzbmb.fsfzh')[1].send_keys(row[7])
    br.find_elements_by_name('currentTZzbmb.fdhhm')[1].send_keys(row[5])
    br.find_element_by_name('currentTZzbmb.fbynd').send_keys(2016)
    br.find_element_by_name('currentTZzbmb.ftxdz').send_keys(row[6])
    br.find_element_by_name('currentTZzbmb.fyzbm').send_keys('234300')


    selectors = br.find_elements_by_css_selector("a.textbox-icon")
    select_option_ids = ['_easyui_combobox_i1_2','_easyui_combobox_i2_0','_easyui_combobox_i6_1',
        '_easyui_combobox_i3_2','_easyui_combobox_i4_12','_easyui_combobox_i5_4'
        ]
    for slt,slt_id in zip(selectors,select_option_ids):
        slt.click()
        br.find_element_by_id(slt_id).click()
    br.find_element_by_partial_link_text('保存').click()
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
    br.find_element_by_partial_link_text('登录').click()
    c = input(':')
    if c == 'q':
        break
br.quit()
