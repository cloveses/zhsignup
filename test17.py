# import win32clipboard as w
import time
import xlrd
# import pyperclip
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
br = webdriver.Firefox()

ws = xlrd.open_workbook('data.xls')
tab = ws.sheets()[0]
nrows = tab.nrows
data_rows = []
for i in range(nrows):
    data_rows.append(tab.row_values(i))

for row in data_rows:
    try:
        datas = [row[0],row[1],row[2],'2017',row[3],'234300']
        br.get("http://60.173.222.146/firstWeb/szksLogin.jsp?loginType=01")
        time.sleep(1)
        br.find_element_by_id('grq_szksdl_fxm').send_keys(row[0])
        time.sleep(0.5)
        br.find_element_by_id('grq_szksdl_fsfzh').send_keys(row[1])
        print(row[0])
        time.sleep(0.5)
        br.find_element_by_id('grq_szksdl_fdhhm').send_keys(row[2])
        time.sleep(0.5)
        br.find_element_by_link_text("注册").click()
        time.sleep(1)

        br.find_elements_by_id('grq_szksdl_fxm')[1].send_keys(row[0])
        br.find_elements_by_id('grq_szksdl_fsfzh')[1].send_keys(row[1])
        br.find_elements_by_id('grq_szksdl_fdhhm')[1].send_keys(row[2])
        br.find_element_by_id('grq_szksdl_fbynd').send_keys(2017)
        br.find_element_by_id('grq_szksdl_ftxdz').send_keys(row[3])
        br.find_element_by_id('grq_szksdl_fyzbm').send_keys('234300')
        
        br.find_element_by_css_selector('select#grq_szksdl_fmmdm + span').click()
        br.find_element_by_id('_easyui_combobox_i1_2').click()

        br.find_element_by_css_selector('select#grq_szksdl_fmzdm + span').click()
        br.find_element_by_id('_easyui_combobox_i2_0').click()

        br.find_element_by_css_selector('select#grq_szksdl_fxldm + span').click()
        br.find_element_by_id('_easyui_combobox_i6_1').click()

        br.find_element_by_css_selector('select#grq_szksdl_fkslbdm + span').click()
        br.find_element_by_id('_easyui_combobox_i3_2').click()

        br.find_element_by_css_selector('select#grq_szksdl_fdsdm + span').click()
        br.find_element_by_id('_easyui_combobox_i4_12').click()
        
        time.sleep(0.5)
        br.find_element_by_css_selector('select#grq_szksdl_fxqdm + span').click()
        br.find_element_by_id('_easyui_combobox_i5_4').click()

        br.find_element_by_partial_link_text('保存').click()

        vtxt = input('captcha:')
        br.find_element_by_id('jcaptcha').send_keys(vtxt)
        
        br.find_element_by_id('login-btn').click()
        time.sleep(2)
        br.find_element_by_partial_link_text('报考中职学校').click()
        time.sleep(4)
        
        br.find_element_by_css_selector('input#grq_f5bm_fyxmc1 + a').click()
        time.sleep(1)
        br.find_element_by_name('currentSchool.fyxmc').send_keys("宿州环保工程学校")
        br.find_element_by_id('yxxz_select_search_ok').click()
        time.sleep(1)
        br.find_element_by_id('datagrid-row-r1-2-0').click()
        br.find_element_by_id('yxxz_select_submit').click()

        ## 选择专业
        time.sleep(3)
        br.find_element_by_css_selector('input#grq_f5bm_fzymc1 + a').click()
        c = input(':')
        br.find_element_by_id('zyxz_select_submit').click()
        time.sleep(1)
        br.find_element_by_partial_link_text('核实无误后').click()
        br.execute_script("$('div.panel:nth-child(26)').eq(0).attr('style','z-index:10000')")
        br.find_element_by_css_selector('.messager-button > a:nth-child(1) > span:nth-child(1) > span:nth-child(1)').click()
        # time.sleep(5)
        # br.find_element_by_partial_link_text('注销').click()
        # c = input(':')
        time.sleep(1)
        br.find_element_by_css_selector('.cupid-blue > span:nth-child(1) > span:nth-child(1)').click()
    except:
        print('error')
    c = input(':')
    if c == 'q':
        break
br.quit()
