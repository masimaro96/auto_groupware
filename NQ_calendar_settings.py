import time, json, openpyxl
import random, testlink
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import WebDriverException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.color import Color
from openpyxl import Workbook

import pathlib
from pathlib import Path
import os
from sys import platform
import NQ_function

from framework_sample import *
from NQ_login_function import driver, data, ValidateFailResultAndSystem, Logging, TesCase_LogResult#, TestlinkResult_Fail, TestlinkResult_Pass

#chrome_path = os.path.dirname(Path(__file__).absolute())+"\\chromedriver.exe"

n = random.randint(1,3000)
m = random.randint(3000,6000)

def calendar(domain_name):
    driver.get(domain_name + "calendar/list/mycal/")
    Logging(" ")
    Logging('============ Menu Calendar ============')
    try:
        popup = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["update"]["title"])))
        if popup.is_displayed():
            Wait10s_ClickElement(data["calendar"]["update"]["button_ok"])
            Logging("Pop up update display")
        else:
            Logging("Pop up not display")
    except WebDriverException:
        Logging("=> Pop up not display")
    time.sleep(5)

    ''' Access to menu '''
    Logging("- Access menu")

    try:
        page_title = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//list-view//h1/span")))
        if page_title.text == 'Custom Calendar':
            setting_user()
            try:
                admin_user = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["calendar_admin"])))
                if admin_user.is_displayed():
                    Logging("- Account admin")
                    admin_user.click()
                    Logging("- Admin calendar")
                    admin_setting()
            except WebDriverException:
                Logging("=> Account user")    
        else:
            Logging("- Calendar: Category type -> Change to folder type")
            time.sleep(5)
            try:
                admin_user_category = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["admin_setting"])))
                if admin_user_category.is_displayed():
                    Logging("- Account admin")
                    change_settings()
                    setting_user()
                try:
                    admin_user = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["calendar_admin"])))
                    if admin_user.is_displayed():
                        Logging("- Account admin")
                        admin_user.click()
                        Logging("- Admin calendar")
                        admin_setting()
                except WebDriverException:
                    Logging("=> Account user")
            except WebDriverException:
                Logging("=> Account user") 

    except WebDriverException:
        Logging("=> Fail")   

def color():
    color_list = int(len(driver.find_elements_by_xpath(data["calendar"]["settings"]["list_color"])))
            
    list_color_calendar = []
    y = 0
    for y in range(color_list):
        y += 1
        list_color_folder = driver.find_element_by_xpath("//*[@id='calendar_setting_form']//form[contains(@name,'categoryForm')]/div//ul/li[" + str(y) + "]")
        list_color_calendar.append(y)

    a = random.choice(list_color_calendar)
    filter_color = driver.find_element_by_xpath("//*[@id='calendar_setting_form']//form[contains(@name,'categoryForm')]/div//ul/li[" + str(a) + "]")
    filter_color.click()

def delete():
    Wait10s_ClickElement(data["calendar"]["settings"]["button_delete"][0])
    Logging("- Select delete sub folder")
    time.sleep(2)

    Wait10s_ClickElement(data["calendar"]["settings"]["button_delete"][1])
    Logging("=> Confirm delete sub folder successfully")
    time.sleep(5)
    
def share_permission():
    Logging("- Select permission for user")
    options_list = ["Permission to Read/Write", "Permission to Read/Write/Delete", "Permission to Read/Write/Modify", "Permission to Read/Write/Modify/Delete"]

    sel = Select(driver.find_element_by_xpath(data["calendar"]["settings"]["dropdown"]))
    sel.select_by_visible_text(random.choice(options_list))
    Logging("=> Select: " + str(random.choice(options_list)))
    time.sleep(2)

def org():
    key_user = data["calendar"]["settings"]["organization_input_1"]
    Wait10s_ClickElement(data["calendar"]["settings"]["select_organization"])
    Logging("- Select organization")

    Wait10s_EnterElement(data["calendar"]["settings"]["organization_input"], key_user)
    Logging("- Input organization")
    time.sleep(2)

    try:
        Wait10s_ClickElement(data["calendar"]["settings"]["organization_select"])
        Logging("- Choose organization")
        time.sleep(2)
        Wait10s_ClickElement(data["calendar"]["settings"]["organization_add"])
        Logging("- Add organization")
        time.sleep(2)
        Wait10s_ClickElement(data["calendar"]["settings"]["organization_save"])
        Logging("- Save organization")
    except:
        Wait10s_ClickElement(data["calendar"]["settings"]["organization_save"])
        pass
    try:
        share_permission()
    except:
        pass

def category():
    input_category = data["calendar"]["settings"]["category_name"]

    Wait10s_ClickElement(data["calendar"]["settings"]["category"])
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["settings"]["category"]))).click()
    Logging("- Click Manage Categories")
    time.sleep(2)

    Wait10s_EnterElement(data["calendar"]["settings"]["input_category_name"], input_category)
    Logging("- Input category name")
    time.sleep(5)

    ''' Change color '''
    Wait10s_ClickElement(data["calendar"]["settings"]["color"])
    time.sleep(3)

    try:
        color()
    except:
        pass

    Wait10s_ClickElement(data["calendar"]["settings"]["button_save"][1])
    Logging("=> Edit success")

def folder_mycalendar():
    namefolder = data["calendar"]["settings"]["name_folder"] + str(n)
    try:
        Logging(" ")
        Logging("** Create folder in my calendar **")
        Wait10s_ClickElement(data["calendar"]["settings"]["add_button"])
        Logging("- Click add button")
        Wait10s_EnterElement(data["calendar"]["settings"]["input_folder_name"], namefolder)
        Logging("- Input name successfully")
        Wait10s_ClickElement(data["calendar"]["settings"]["button_save"][0])
        Logging("- Add folder successfully")
    except:
        pass

    try:
        folder_exist = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, "//*[@id='calendar_setting_form']//span[contains(., 'A folder with the same name already exists. Force apply?')]")))
        if folder_exist.is_displayed():
            Logging("A folder with the same name already exists")
            Wait10s_Clear_InputElement(data["calendar"]["settings"]["input_folder_name"], namefolder)
            Wait10s_ClickElement(data["calendar"]["settings"]["button_save"][2])
            Logging("- Input another name")
    except:
        pass
    time.sleep(6)

    try:
        ''' Check if folder have been create '''
        Logging("** Check if folder have been create **")
        folder_mycalendar = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='calendar_setting_form']//span//a[contains(., '" + namefolder + "')]")))
        if folder_mycalendar.is_displayed():
            Logging("=> Create folder success")
            TesCase_LogResult(**data["testcase_result"]["calendar"]["add_folder_setting"]["pass"])
        else:
            Logging("=> Create folder fail")
            TesCase_LogResult(**data["testcase_result"]["calendar"]["add_folder_setting"]["fail"])
            ValidateFailResultAndSystem("<div>[Calendar]Create folder in my calendar fail </div>")
    except:
        Logging("=> Create folder fail")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["add_folder_setting"]["fail"])
        pass
    time.sleep(5)
    return namefolder

def sub_folder_mycalendar(namefolder):
    name_subfolder = data["calendar"]["settings"]["name_folder_2"] + str(n)
    try:
        Logging(" ")
        Logging("** Create sub folder of main folder have create")
        Wait10s_ClickElement(data["calendar"]["settings"]["select_folder"])
        Logging("- Show dropdowm parent folder")
        time.sleep(2)
        Wait10s_ClickElement("//*[@id='calendar_setting_form']//form//li//a[contains(., '" + namefolder + "')]")
        Logging("- Select parent folder")
        Wait10s_InputElement(data["calendar"]["settings"]["input_folder_name"], name_subfolder)
        Logging("- Input name successfully")
        Wait10s_ClickElement(data["calendar"]["settings"]["button_save"][0])
        Logging("- Click save successfully")
    except:
        pass

    time.sleep(5)

    ''' Check if sub folder have been create '''
    try:
        Logging("** Check if sub folder have been create **")
        Wait10s_ClickElement("//*[@id='calendar_setting_form']//span//a[contains(., '" + namefolder + "')]")
        time.sleep(3)              
        subfolder_mycalendar = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='calendar_setting_form']//span//a[contains(., '" + name_subfolder + "')]")))
        if subfolder_mycalendar.is_displayed():
            Logging("=> Create sub folder success => Pass")
            TesCase_LogResult(**data["testcase_result"]["calendar"]["add_subfolder_setting"]["pass"])
        else:
            Logging("=> Create sub folder fail => Fail")
            TesCase_LogResult(**data["testcase_result"]["calendar"]["add_subfolder_setting"]["fail"])
            ValidateFailResultAndSystem("<div>[Calendar]Create sub folder fail </div>")
        time.sleep(5)
    except:
        Logging("=> Create sub folder fail => Fail")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["add_subfolder_setting"]["fail"])
        pass
    return name_subfolder

def edit_sub_folder(name_subfolder):
    ''' Pull srcoll bar to down '''
    driver.find_element_by_tag_name("body").send_keys(Keys.HOME)
    time.sleep(5)
    
    ''' Edit sub folder '''
    try:
        Logging(" ")
        Logging("** Edit sub folder **")
        Wait10s_ClickElement("//*[@id='calendar_setting_form']//span//a[contains(., '" + name_subfolder + "')]")
        Logging("- Select sub folder to edit")
        time.sleep(2)

        try:
            category()
        except:
            pass

        TesCase_LogResult(**data["testcase_result"]["calendar"]["edit_subfolder_setting"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["edit_subfolder_setting"]["fail"])
        pass       
    time.sleep(5)

def delete_sub_folder(name_subfolder):
    try:
        ''' Delete sub folder '''
        Logging(" ")
        Logging("** Delete sub folder **")
        Wait10s_ClickElement("//*[@id='calendar_setting_form']//li//a[contains(., '" + name_subfolder + "')]")
        Logging("- Select sub folder")
        time.sleep(2)
        delete()
        TesCase_LogResult(**data["testcase_result"]["calendar"]["delete_subfolder_setting"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["delete_subfolder_setting"]["fail"])
        pass

def edit_folder(namefolder):
    ''' Edit main folder '''
    try:
        Logging(" ")
        Logging("** Edit main folder in my calendar **")
        Wait10s_ClickElement("//*[@id='calendar_setting_form']//span[contains(., '" + namefolder + "')]")
        Logging("- Click folder to edit")
        time.sleep(2)
        Wait10s_ClickElement(data["calendar"]["settings"]["share_folder"])
        Logging("- Select share folder")

        driver.find_element_by_tag_name("body").send_keys(Keys.END)
        time.sleep(5)
        try: 
            org()
        except:
            pass
        Wait10s_ClickElement(data["calendar"]["settings"]["button_save"][0])
        time.sleep(5)
        driver.find_element_by_tag_name("body").send_keys(Keys.HOME)
        time.sleep(5)

        Wait10s_ClickElement("//*[@id='calendar_setting_form']//span[contains(., '" + namefolder + "')]")
        time.sleep(2)
        try:
            category()
        except:
            pass
        TesCase_LogResult(**data["testcase_result"]["calendar"]["edit_folder_setting"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["edit_folder_setting"]["fail"])
        pass

    time.sleep(5)

def delete_folder(namefolder):
    driver.find_element_by_tag_name("body").send_keys(Keys.HOME)
    try:
        ''' Delete main folder '''
        Logging(" ")
        Logging("** Delete main folder")
        Wait10s_ClickElement(data["calendar"]["settings"]["share_folder"])
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[@id='calendar_setting_form']//li//a[contains(., '" + namefolder + "')]"))).click()
        Logging("- Select folder")
        time.sleep(2)
        # Keys.End to scroll down
        driver.find_element_by_tag_name("body").send_keys(Keys.HOME)
        delete()
        TesCase_LogResult(**data["testcase_result"]["calendar"]["delete_folder_setting"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["delete_folder_setting"]["fail"])
        pass

def share_user():
    key_org = data["calendar"]["settings"]["organization_input_1"]

    Wait10s_ClickElement(data["calendar"]["Admin"]["share"])
    Logging("- Select share folder")
    time.sleep(5)
    
    driver.find_element_by_tag_name("body").send_keys(Keys.END)
    time.sleep(5)
    Wait10s_ClickElement(data["calendar"]["Admin"]["organization"])
    time.sleep(5)

    Wait10s_InputElement(data["calendar"]["Admin"]["input_organization"], key_org)
    Logging("- Input name organization")
    time.sleep(5)
    Wait10s_ClickElement(data["calendar"]["Admin"]["select_user"])
    time.sleep(5)
    Wait10s_ClickElement(data["calendar"]["Admin"]["add_user"])
    Logging("- Add user organization")
    time.sleep(5)
    Wait10s_ClickElement(data["calendar"]["Admin"]["save_user"])
    Logging("- Share company calendar success") 
    time.sleep(5)

def setting_view():
    view_list = int(len(driver.find_elements_by_xpath(data["calendar"]["settings"]["setting_view"])))
    view_mode_list = []
    i = 0
    for i in range(view_list):
        i += 1
        mode_list = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["settings"]["view_mode"] + "[" + str(i) + "]/label")))
        view_mode_list.append(mode_list.text)
    
    Logging("- Total of view mode list: " + str(len(view_mode_list)))
    x = random.choice(view_mode_list)
    time.sleep(2)
    select_mode_view = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["settings"]["setting_view"] + "[contains(.,'" + str(x) + "')]")))
    select_mode_view.click()
    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["settings"]["button_save"][3]))).click()

def folder_company():
    name = data["calendar"]["Admin"]["name_input"] + str(n)
    ''' Go to manage company folders '''
    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["company_folder"]))).click()
    Logging(" ")
    Logging("** Manage company folders")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["folder_name"]))).send_keys(name)  
    Logging("- Input name of folder")
    time.sleep(2) 

    try:
        ''' Change color '''
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["color"]))).click()
        time.sleep(2)

        color_list = int(len(driver.find_elements_by_xpath(data["calendar"]["Admin"]["list_color"])))
        
        list_color_calendar = []
        y = 0
        for y in range(color_list):
            y += 1
            list_color_folder = driver.find_element_by_xpath("//*[@id='boot-strap-valid']//a[contains(@data-toggle, 'dropdown')]/following-sibling::ul/li[" + str(y) + "]")
            list_color_calendar.append(y)

        a = random.choice(list_color_calendar)
        filter_color = driver.find_element_by_xpath("//*[@id='boot-strap-valid']//a[contains(@data-toggle, 'dropdown')]/following-sibling::ul/li[" + str(a) + "]")
        filter_color.click()
    except:
        pass

    driver.find_element_by_tag_name("body").send_keys(Keys.END)
    time.sleep(3)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["save"]))).click()
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["close_dialog"]))).click()
    Logging("=> Create folder in company calendar")
    time.sleep(5)

    try:
        ''' Check folder have create '''
        folder_company = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.calendar.admin_company']//span//a[contains(., '" + name + "')]")))
        if folder_company.is_displayed:
            Logging("=> Create folder success => Pass")
            TesCase_LogResult(**data["testcase_result"]["calendar"]["add_folder"]["pass"])
        else:
            Logging("=> Create folder fail => Fail")
            TesCase_LogResult(**data["testcase_result"]["calendar"]["add_folder"]["fail"])
            ValidateFailResultAndSystem("<div>[Calendar]Create folder admin fail </div>")
    except:
        Logging("=> Create folder fail => Fail")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["add_folder"]["fail"])
        pass
    return name

def edit_folder_company(name):
    driver.find_element_by_tag_name("body").send_keys(Keys.HOME)
    time.sleep(5)
    edit_name = data["calendar"]["Admin"]["category_input"] + str(n)

    try:
        ''' Edit - add category for folder company '''
        Logging(" ")
        Logging("** Edit - add category for folder company")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.calendar.admin_company']//span//a[contains(., '" + name + "')]"))).click()
        Logging("- Select folder to edit")

        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["category_company"]))).click()
        Logging("- Click icon add Category List")
        time.sleep(2) 
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["category_company_input"]))).send_keys(edit_name)
        Logging("- Input name category")
        time.sleep(2)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["add_category"]))).click()
        Logging("- Click save")
        time.sleep(2) 
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["close_dialog"]))).click()
        Logging("=> Confirm add category success")
        time.sleep(5)

        ''' Set permission to share user '''
        try:
            share_user()
        except:
            driver.find_element_by_xpath(data["calendar"]["Admin"]["save_user"]).click()
            pass
        try:
            share_permission()
        except:
            pass
        
        driver.find_element_by_tag_name("body").send_keys(Keys.END)
        time.sleep(5)
        driver.find_element_by_xpath(data["calendar"]["Admin"]["save"]).click()
        time.sleep(5)
        driver.find_element_by_xpath(data["calendar"]["Admin"]["close_dialog"]).click()
        time.sleep(5)
        Logging("- Save edit folder company calendar")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["edit_folder"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["edit_folder"]["fail"])
        pass

def search_folderlist():
    ''' Go to calendar folder list '''
    try:
        Logging(" ")
        Logging("** Calendar folder list")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["folder_list"]))).click()
        time.sleep(5)
        search = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["search"])))
        search.send_keys(data["calendar"]["Admin"]["search_input"])
        Logging("- Input search name")
        search.send_keys(Keys.ENTER)
        Logging("=> Search success")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["search_folder"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["search_folder"]["fail"])
        pass

def delete_folder_company(name):
    try:
        Logging(" ")
        Logging("** Delete folder after search")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.calendar.admin_folder']//table//tr[contains(., '" + name + "')]"))).click()
        Logging("- Select folder")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["button_delete"][0]))).click()
        Logging("- Click delete button")
        driver.find_element_by_xpath(data["calendar"]["Admin"]["close_dialog_1"]).click()
        Logging("- Confirm delete folder")
        driver.find_element_by_xpath(data["calendar"]["Admin"]["close_dialog"]).click()
        Logging("=> Delete folder success")  
        TesCase_LogResult(**data["testcase_result"]["calendar"]["delete_folder"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["delete_folder"]["fail"])
        pass

def setting_calendar_company():
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["setting_calendar"]))).click()
    Logging("** Setting admin")

    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["loading_dialog"])))

    ''' Can't check input have select or not '''    
    try:        
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["birthday"]))).click()
        Logging("- Check show birthday")

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["holiday"]))).click()
        Logging("- Check Hide Holiday Calendar")
        
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["dept"]))).click()
        Logging("- Check Hide Dept. Calendar")

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["save_settings"]))).click()
        Logging("=> Save setting")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["select_setting_admin"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["select_setting_admin"]["fail"])
        pass

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["close_popup"]))).click()
    Logging("=> Close pop up")
    time.sleep(5)
    '''WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["category_type"]))).click()
    Logging("** Change to Category Type")
    time.sleep(10)'''

def change_settings():
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["admin_setting"]))).click()
    Logging("** Setting admin")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["folder_type"]))).click()
    Logging("=> Use folder type")
    time.sleep(20)

def category_folder():
    name_category_type = data["calendar"]["category_type"]["input_text"] + str(m)
    name_category_type_edit = data["calendar"]["category_type"]["category_edit_name"] + str(n)

    Logging(" ")
    Logging("============ Calendar - Category type ============")
    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["settings"])))
    element.location_once_scrolled_into_view
    driver.find_element_by_xpath(data["calendar"]["category_type"]["settings"]).click()
    Logging("=> Click settings")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["add_category"]))).click()
    Logging("- Add category")
    time.sleep(5)

    ''' Change color '''
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["color"]))).click()
    time.sleep(2)

    color_list = int(len(driver.find_elements_by_xpath(data["calendar"]["category_type"]["list_color"])))
    
    list_color_calendar = []
    y = 0
    for y in range(color_list):
        y += 1
        list_color_folder = driver.find_element_by_xpath("//div[contains(@ng-click, 'close($event)')]//form[contains(@method, 'post')]/div//ul/li[" + str(y) + "]")
        list_color_calendar.append(y)

    a = random.choice(list_color_calendar)
    filter_color = driver.find_element_by_xpath("//div[contains(@ng-click, 'close($event)')]//form[contains(@method, 'post')]/div//ul/li[" + str(a) + "]")
    filter_color.click()

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["input"]))).send_keys(name_category_type)
    Logging("- Input name category")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["share"]))).click()
    Logging("- Share category")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["org"]))).click()
    Logging("- Click org")
    time.sleep(5)
    name = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["input_org"])))
    name.send_keys(data["calendar"]["category_type"]["org_text"])
    time.sleep(2)
    name.send_keys(Keys.ENTER)
    Logging("- Search org")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["select_user_1"]))).click()
    Logging("- Select user")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["add_user"]))).click()
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["button_save_org"]))).click()
    Logging("- Save org")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["button_save"]))).click()
    Logging("=> Save category")
    time.sleep(5)

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.calendarnew.settings']//tr//td[contains(., '" + name_category_type + "')]//following-sibling::td//a[contains(@ng-click, 'edit(e, item)')]"))).click()
    time.sleep(2)

    ''' Change color '''
    name = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["category_edit"])))
    name.clear()
    name.send_keys(name_category_type_edit)
    Logging("- Edit name category")

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["descriptions"]))).click()
    Logging("- Add descriptions")

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["org"]))).click()
    Logging("- Click org")
    time.sleep(5)
    name = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["input_org"])))
    name.send_keys(data["calendar"]["category_type"]["org_text"])
    time.sleep(2)
    name.send_keys(Keys.ENTER)
    Logging("- Search org")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["select_user_2"]))).click()
    Logging("- Select user")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["add_user"]))).click()
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["button_save_org"]))).click()
    Logging("- Save org")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["button_save"]))).click()
    Logging("=> Save category")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.calendarnew.settings']//tr//td[contains(., '" + name_category_type_edit + "')]//following-sibling::td//a[contains(@ng-click, 'delete(e, item)')]"))).click()
    Logging("- Select delete category")
    time.sleep(2)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["delete"]))).click()
    Logging("=> Delete category")
    time.sleep(5)

def setting_user():
    Logging(" ")
    Logging("============ Calendar - Folder type ============")
    driver.find_element_by_xpath(data["calendar"]["settings"]["settings_calendar"]).click()
    Logging("- Click settings")
    try:
        title = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.calendar.setting']//h1[contains(., 'Settings')]")))
        if title.text == 'Settings':
            Logging("Access settings")
    except:
        pass
    Logging(" ")
    Logging("============ Test case settings calendar ============")

    ''' Create folder in my calendar '''
    try:
        namefolder = folder_mycalendar()
    except:
        namefolder = None
    
    if bool(namefolder) == True:
        name_subfolder = sub_folder_mycalendar(namefolder)
        if bool(name_subfolder) == True:
            edit_sub_folder(name_subfolder)
            delete_sub_folder(name_subfolder)
        else:
            Logging(">> Can't continue execution")
    else:
        Logging("=> Create sub folder fail => Fail")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["add_subfolder_setting"]["fail"])

    try:
        edit_folder(namefolder)
        delete_folder(namefolder)
    except:
        Logging(">> Can't continue execution")
        pass

def admin_setting():
    Logging(" ")
    Logging("============ Test case Admin calendar ============")

    try:
        name = folder_company()
    except:
        name = None
    
    if bool(name) == True:
        try:
            edit_folder_company(name)
            search_folderlist()
            delete_folder_company(name)
            setting_calendar_company()
        except:
            Logging(">> Can't continue execution")
            pass
    else:
        Logging("=> Create folder fail => Fail")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["add_folder"]["fail"])

