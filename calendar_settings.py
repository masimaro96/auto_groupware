import time, json, openpyxl
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import WebDriverException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from datetime import datetime
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
date_time = datetime.now().strftime("%Y/%m/%d, %H:%M:%S")

def calendar(domain_name):
    driver.get(domain_name + "calendar/list/mycal/")
    Logging(" ")
    PrintGreen('============ Menu Calendar ============')
    try:
        popup = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["update"]["title"])))
        if popup.is_displayed():
            Commands.Wait20s_ClickElement(data["calendar"]["update"]["button_ok"])
            Logging("Pop up update display")
        else:
            Logging("Pop up not display")
    except WebDriverException:
        Logging("=> Pop up not display")
    

    ''' Access to menu '''
    Logging("- Access menu")

    try:
        page_title = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//list-view//h1/span")))
        if page_title.text == 'Custom Calendar':
            try:
                admin_user = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["calendar_admin"])))
                if admin_user.is_displayed():
                    setting_user()
                    PrintGreen("- Account admin")
                    admin_user.click()
                    PrintGreen("- Admin calendar")
                    admin_setting()
            except WebDriverException:
                PrintGreen("=> Account user")
                setting_user()    
        else:
            Logging("- Calendar: Category type -> Change to folder type")
            
            try:
                admin_user_category = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["admin_setting"])))
                if admin_user_category.is_displayed():
                    PrintGreen("- Account admin")
                    change_settings()
                    setting_user()
                try:
                    admin_user = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["calendar_admin"])))
                    if admin_user.is_displayed():
                        PrintGreen("- Account admin")
                        admin_user.click()
                        PrintGreen("- Admin calendar")
                        admin_setting()
                except WebDriverException:
                    PrintGreen("=> Account user")
                    setting_user()
            except WebDriverException:
                PrintGreen("=> Account user")
                setting_user() 

    except WebDriverException:
        PrintRed("=> Fail")   

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
    Commands.Wait20s_ClickElement(data["calendar"]["settings"]["button_delete"][0])
    Logging("- Select delete sub folder")
    time.sleep(5)
    Commands.Wait20s_ClickElement(data["calendar"]["settings"]["button_delete"][1])
    Logging("=> Confirm delete sub folder successfully")
    
    
def share_permission():
    Logging("- Select permission for user")
    options_list = ["Permission to Read/Write", "Permission to Read/Write/Delete", "Permission to Read/Write/Modify", "Permission to Read/Write/Modify/Delete"]
    time.sleep(5)
    sel = Select(driver.find_element_by_xpath(data["calendar"]["settings"]["dropdown"]))
    sel.select_by_visible_text(random.choice(options_list))
    Logging("=> Select: " + str(random.choice(options_list)))
    

def org():
    key_user = data["calendar"]["settings"]["organization_input_1"]
    Commands.Wait20s_ClickElement(data["calendar"]["settings"]["select_organization"])
    Logging("- Select organization")
    time.sleep(5)
    Commands.Wait20s_EnterElement(data["calendar"]["settings"]["organization_input"], key_user)
    Logging("- Input organization")
    

    try:
        Commands.Wait20s_ClickElement(data["calendar"]["settings"]["organization_select"])
        Logging("- Choose organization")
        
        Commands.Wait20s_ClickElement(data["calendar"]["settings"]["organization_add"])
        Logging("- Add organization")
        
        Commands.Wait20s_ClickElement(data["calendar"]["settings"]["organization_save"])
        Logging("- Save organization")
    except:
        Commands.Wait20s_ClickElement(data["calendar"]["settings"]["organization_save"])
        pass
    try:
        share_permission()
    except:
        pass

def category():
    input_category = data["calendar"]["settings"]["category_name"]

    Commands.Wait20s_ClickElement(data["calendar"]["settings"]["category"])
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["settings"]["category"]))).click()
    Logging("- Click Manage Categories")
    time.sleep(5)
    Commands.Wait20s_EnterElement(data["calendar"]["settings"]["input_category_name"], input_category)
    Logging("- Input category name")
    
    ''' Change color '''
    Commands.Wait20s_ClickElement(data["calendar"]["settings"]["color"])
    time.sleep(5)
    try:
        color()
    except:
        pass

    Commands.Wait20s_ClickElement(data["calendar"]["settings"]["button_save"][1])
    Logging("=> Edit success")

def folder_mycalendar():
    namefolder = data["calendar"]["settings"]["name_folder"] + str(n)
    try:
        Logging(" ")
        Logging("** Create folder in my calendar **")
        Commands.Wait20s_ClickElement(data["calendar"]["settings"]["add_button"])
        Logging("- Click add button")
        Commands.Wait20s_InputElement(data["calendar"]["settings"]["input_folder_name"], namefolder)
        Logging("- Input name successfully")
        Commands.Wait20s_ClickElement(data["calendar"]["settings"]["button_save"][0])
        Logging("- Add folder successfully")
    except:
        pass

    try:
        folder_exist = Waits.Wait20s_ElementLoaded("//*[@id='calendar_setting_form']//span[contains(., 'A folder with the same name already exists. Force apply?')]")
        if folder_exist.is_displayed():
            Logging("A folder with the same name already exists")
            Commands.Wait20s_Clear_InputElement(data["calendar"]["settings"]["input_folder_name"], namefolder)
            Commands.Wait20s_ClickElement(data["calendar"]["settings"]["button_save"][2])
            Logging("- Input another name")
    except:
        pass
    

    try:
        ''' Check if folder have been create '''
        Logging("** Check if folder have been create **")
        Waits.Wait20s_ElementLoaded("//*[@id='calendar_setting_form']//span//a[contains(., '" + str(namefolder) + "')]")
        Logging("=> Create folder success")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["add_folder_setting"]["pass"])
    except:
        Logging("=> Create folder fail")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["add_folder_setting"]["fail"])
        pass
    
    
    return namefolder

def sub_folder_mycalendar(namefolder):
    name_subfolder = data["calendar"]["settings"]["name_sub_folder"] + str(n)
    try:
        Logging(" ")
        Logging("** Create sub folder of main folder have create")
        Commands.Wait20s_ClickElement(data["calendar"]["settings"]["select_folder"])
        Logging("- Show dropdowm parent folder")
        
        Commands.Wait20s_ClickElement("//*[@id='calendar_setting_form']//form//li//a[contains(., '" + namefolder + "')]")
        Logging("- Select parent folder")
        Commands.Wait20s_InputElement(data["calendar"]["settings"]["input_folder_name"], name_subfolder)
        Logging("- Input name successfully")
        Commands.Wait20s_ClickElement(data["calendar"]["settings"]["button_save"][0])
        Logging("- Click save successfully")
    except:
        pass

    ''' Check if sub folder have been create '''
    try:
        Logging("** Check if sub folder have been create **")
        Commands.Wait20s_ClickElement("//*[@id='calendar_setting_form']//span//a[contains(., '" + namefolder + "')]")
        time.sleep(5)              
        Waits.Wait20s_ElementLoaded("//*[@id='calendar_setting_form']//span//a[contains(., '" + str(name_subfolder) + "')]")
        Logging("=> Create sub folder success => Pass")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["add_subfolder_setting"]["pass"])
    except:
        Logging("=> Create sub folder fail => Fail")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["add_subfolder_setting"]["fail"])

    return name_subfolder

def edit_sub_folder(name_subfolder):
    ''' Pull srcoll bar to down '''
    driver.find_element_by_tag_name("body").send_keys(Keys.HOME)
    
    
    ''' Edit sub folder '''
    try:
        Logging(" ")
        Logging("** Edit sub folder **")
        Commands.Wait20s_ClickElement("//*[@id='calendar_setting_form']//span//a[contains(., '" + str(name_subfolder) + "')]")
        Logging("- Select sub folder to edit")
        

        try:
            category()
        except:
            pass

        TesCase_LogResult(**data["testcase_result"]["calendar"]["edit_subfolder_setting"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["edit_subfolder_setting"]["fail"])      
   

def delete_sub_folder(name_subfolder):
    try:
        ''' Delete sub folder '''
        Logging(" ")
        Logging("** Delete sub folder **")
        Commands.Wait20s_ClickElement("//*[@id='calendar_setting_form']//li//a[contains(., '" + str(name_subfolder) + "')]")
        Logging("- Select sub folder")
        
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
        Commands.Wait20s_ClickElement("//*[@id='calendar_setting_form']//span[contains(., '" + namefolder + "')]")
        Logging("- Click folder to edit")
        
        Commands.Wait20s_ClickElement(data["calendar"]["settings"]["share_folder"])
        Logging("- Select share folder")

        driver.find_element_by_tag_name("body").send_keys(Keys.END)
        
        try: 
            org()
        except:
            pass
        Commands.Wait20s_ClickElement(data["calendar"]["settings"]["button_save"][0])
        
        driver.find_element_by_tag_name("body").send_keys(Keys.HOME)
        

        Commands.Wait20s_ClickElement("//*[@id='calendar_setting_form']//span[contains(., '" + namefolder + "')]")
        
        try:
            category()
        except:
            pass
        TesCase_LogResult(**data["testcase_result"]["calendar"]["edit_folder_setting"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["edit_folder_setting"]["fail"])
        pass

    

def delete_folder(namefolder):
    driver.find_element_by_tag_name("body").send_keys(Keys.HOME)
    try:
        ''' Delete main folder '''
        Logging(" ")
        Logging("** Delete main folder")
        Commands.Wait20s_ClickElement("//*[@id='calendar_setting_form']//li//a[contains(., '" + namefolder + "')]")
        Logging("- Select folder")
        
        # Keys.End to scroll down
        driver.find_element_by_tag_name("body").send_keys(Keys.HOME)
        delete()
        TesCase_LogResult(**data["testcase_result"]["calendar"]["delete_folder_setting"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["delete_folder_setting"]["fail"])
        pass

def share_user():
    key_org = data["calendar"]["settings"]["organization_input_1"]

    Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["share"])
    Logging("- Select share folder")
    
    
    driver.find_element_by_tag_name("body").send_keys(Keys.END)
    
    Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["organization"])
    

    Commands.Wait20s_InputElement(data["calendar"]["Admin"]["input_organization"], key_org)
    Logging("- Input name organization")
    
    Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["select_user"])
    
    Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["add_user"])
    Logging("- Add user organization")
    
    Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["save_user"])
    Logging("- Share company calendar success") 
    

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
    
    select_mode_view = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["settings"]["setting_view"] + "[contains(.,'" + str(x) + "')]")))
    select_mode_view.click()
    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["settings"]["button_save"][3]))).click()

def folder_company():
    name = data["calendar"]["Admin"]["name_input"] + str(n)
    # name = data["title"] + date_time
    ''' Go to manage company folders '''
    Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["company_folder"])
    Logging(" ")
    Logging("** Manage company folders")
    
    Commands.Wait20s_InputElement(data["calendar"]["Admin"]["folder_name"], name) 
    Logging("- Input name of folder")
     

    try:
        ''' Change color '''
        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["color"])
        time.sleep(5)
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
        Logging("- Select color")
    except:
        pass

    driver.find_element_by_tag_name("body").send_keys(Keys.END)
    
    Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["save"])
    
    Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["close_dialog"])
    Logging("=> Create folder in company calendar")
    

    try:
        ''' Check folder have create '''
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.calendar.admin_company']//span//a[contains(., '" + name + "')]")))
        Logging("=> Create folder success => Pass")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["add_folder"]["pass"])
    except:
        Logging("=> Create folder fail => Fail")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["add_folder"]["fail"]) 
        pass
    return name

def edit_folder_company(name):
    driver.find_element_by_tag_name("body").send_keys(Keys.HOME)
    
    edit_name = data["calendar"]["Admin"]["category_input"] + str(n)

    try:
        ''' Edit - add category for folder company '''
        Logging(" ")
        Logging("** Edit - add category for folder company")
        Commands.Wait20s_ClickElement("//*[@id='ngw.calendar.admin_company']//span//a[contains(., '" + name + "')]")
        Logging("- Select folder to edit")

        
        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["category_company"])
        Logging("- Click icon add Category List")
         
        Commands.Wait20s_InputElement(data["calendar"]["Admin"]["category_company_input"], edit_name)
        Logging("- Input name category")
        

        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["add_category"])
        Logging("- Click save")
         
        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["close_dialog"])
        Logging("=> Confirm add category success")
        

        ''' Set permission to share user '''
        try:
            share_user()
        except:
            Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["save_user"])
            pass
        try:
            share_permission()
        except:
            pass
        
        driver.find_element_by_tag_name("body").send_keys(Keys.END)
        
        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["save"])
        
        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["close_dialog"])
        
        Logging("- Save edit folder company calendar")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["edit_folder"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["edit_folder"]["fail"])
        pass

def search_folderlist():
    key_search = data["calendar"]["Admin"]["search_input"]
    ''' Go to calendar folder list '''
    try:
        Logging(" ")
        Logging("** Calendar folder list")
        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["folder_list"])
        
        Commands.Wait20s_EnterElement(data["calendar"]["Admin"]["search"], key_search)
        Logging("- Input search name")
        Logging("=> Search success")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["search_folder"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["search_folder"]["fail"])
        pass

def delete_folder_company(name):
    try:
        Logging(" ")
        Logging("** Delete folder after search")
        Commands.Wait20s_ClickElement("//*[@id='ngw.calendar.admin_folder']//table//tr[contains(., '" + name + "')]")
        Logging("- Select folder")
        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["button_delete"][0])
        Logging("- Click delete button")
        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["close_dialog_1"])
        Logging("- Confirm delete folder")
        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["close_dialog"])
        Logging("=> Delete folder success")  
        TesCase_LogResult(**data["testcase_result"]["calendar"]["delete_folder"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["delete_folder"]["fail"])
        pass

def setting_calendar_company():
    Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["setting_calendar"])
    Logging("** Setting admin")

    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["loading_dialog"])))

    ''' Can't check input have select or not '''    
    try:        
        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["birthday"])
        Logging("- Check show birthday")

        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["holiday"])
        Logging("- Check Hide Holiday Calendar")
        
        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["dept"])
        Logging("- Check Hide Dept. Calendar")

        Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["save_settings"])
        Logging("=> Save setting")
        TesCase_LogResult(**data["testcase_result"]["calendar"]["select_setting_admin"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["calendar"]["select_setting_admin"]["fail"])
        pass

    Commands.Wait20s_ClickElement(data["calendar"]["Admin"]["close_popup"])
    Logging("=> Close pop up")
    
    '''WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["Admin"]["category_type"]))).click()
    Logging("** Change to Category Type")
    '''

def change_settings():
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["admin_setting"])
    Logging("** Setting admin")
    
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["folder_type"])
    Logging("=> Use folder type")
    

def category_folder():
    name_category_type = data["calendar"]["category_type"]["input_text"] + str(m)
    name_category_type_edit = data["calendar"]["category_type"]["category_edit_name"] + str(n)
    org_key = data["calendar"]["category_type"]["org_text"]

    Logging(" ")
    Logging("============ Calendar - Category type ============")
    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["calendar"]["category_type"]["settings"])))
    element.location_once_scrolled_into_view
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["settings"])
    Logging("=> Click settings")
    
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["add_category"])
    Logging("- Add category")
    

    ''' Change color '''
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["color"])
    

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

    Commands.Wait20s_InputElement(data["calendar"]["category_type"]["input"], name_category_type)
    Logging("- Input name category")
    
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["share"])
    Logging("- Share category")
    
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["org"])
    Logging("- Click org")
    
    Commands.Wait20s_EnterElement(data["calendar"]["category_type"]["input_org"], org_key)
    Logging("- Search org")
    
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["select_user_1"])
    Logging("- Select user")
    
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["add_user"])
    
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["button_save_org"])
    Logging("- Save org")
    
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["button_save"])
    Logging("=> Save category")
    

    Commands.Wait20s_ClickElement("//*[@id='ngw.calendarnew.settings']//tr//td[contains(., '" + name_category_type + "')]//following-sibling::td//a[contains(@ng-click, 'edit(e, item)')]")
    

    ''' Change color '''
    Commands.Wait20s_Clear_InputElement(data["calendar"]["category_type"]["category_edit"], name_category_type_edit)
    Logging("- Edit name category")

    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["descriptions"])
    Logging("- Add descriptions")

    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["org"])
    Logging("- Click org")
    
    Commands.Wait20s_EnterElement(data["calendar"]["category_type"]["input_org"], org_key)
    Logging("- Search org")
    
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["select_user_2"])
    Logging("- Select user")
    
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["add_user"])
    
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["button_save_org"])
    Logging("- Save org")
    
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["button_save"])
    Logging("=> Save category")
    
    Commands.Wait20s_ClickElement("//*[@id='ngw.calendarnew.settings']//tr//td[contains(., '" + name_category_type_edit + "')]//following-sibling::td//a[contains(@ng-click, 'delete(e, item)')]")
    Logging("- Select delete category")
    
    Commands.Wait20s_ClickElement(data["calendar"]["category_type"]["delete"])
    Logging("=> Delete category")
    

def setting_user():
    Logging(" ")
    Logging("============ Calendar - Folder type ============")
    Commands.Wait20s_ClickElement(data["calendar"]["settings"]["settings_calendar"])
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

