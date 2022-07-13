import time, json, openpyxl
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import WebDriverException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from datetime import datetime
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains
from openpyxl import Workbook

import pathlib
from pathlib import Path
import os
from sys import platform
import NQ_function

from framework_sample import *
from NQ_login_function import driver, data, ValidateFailResultAndSystem, Logging, TesCase_LogResult

n = random.randint(1,3000)
m = random.randint(3000,6000)

date_time = datetime.now().strftime("%Y/%m/%d, %H:%M:%S")

def todo(domain_name):
    driver.get(domain_name + "todo/setting/setting/")

    Logging(" ")
    PrintGreen("============ Menu To-Do ============")

    ''' Access to page To-Do '''
    PrintGreen("- Access menu")
    
    try:
        
        admin = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, data["todo"]["admin"])))
        if admin.is_displayed():
            setting_execution()
            PrintGreen("- Account admin")
            admin.click()
            
            admin_execution()
    except:
        PrintGreen("=> Account user")
        setting_execution()

def add_folder():
    name_folder = data["title"] + date_time

    PrintYellow("-----Add folder------")
    Commands.Wait20s_ClickElement(data["todo"]["add_button"])
    Logging("- Select button add folder")
    
    Commands.Wait20s_InputElement(data["todo"]["add_folder_name"], name_folder)
    Logging("- Input name folder")
    

    ''' Check button save have work '''
    try:
        Commands.Wait20s_ClickElement(data["todo"]["save_button"])
        
        Logging("=> Save success")

        ''' Check folder to-do have create '''
        PrintYellow("** Check folder have save yet!!")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.todo.list.setting']//span//a[contains(., '" + name_folder + "')]")))
        Logging("=> Folder have add successfully")
        TesCase_LogResult(**data["testcase_result"]["todo"]["add_folder"]["pass"])
            
    except WebDriverException:
        Logging("=> Save fail")
        TesCase_LogResult(**data["testcase_result"]["todo"]["add_folder"]["fail"])
    

    return name_folder

def delete_folder(name_folder):   
    try:
        Commands.Wait20s_ClickElement("//*[@id='ngw.todo.list.setting']//span//a[contains(., '" + name_folder + "')]")
        Logging("- Select folder")
        Commands.Wait20s_ClickElement(data["todo"]["del_button"])
        
        Commands.Wait20s_ClickElement(data["todo"]["del_button_1"])
        PrintYellow("=> Delete folder")
        TesCase_LogResult(**data["testcase_result"]["todo"]["delete_folder"]["pass"])
        
    except:
        TesCase_LogResult(**data["testcase_result"]["todo"]["delete_folder"]["fail"])
        pass

def setting_execution():
    Commands.Wait20s_ClickElement(data["todo"]["settings_todo"])
    

    page_title = driver.find_element_by_xpath("//*[@id='ngw.todo.list.setting']//span")
    if page_title.text == 'Settings':
        Logging("- Click settings success")
    else:
        Logging("- Click settings fail")
    
    Logging(" ")
    PrintGreen("============ Test case settings To-Do ============")

    try:
        name_folder = add_folder()
    except:
        name_folder = None

    if bool(name_folder) == True:
        try:
            delete_folder(name_folder)
        except:
            Logging(">> Can't countinue execution")
            pass
    else:
        PrintRed("=> Add folder todo fail")
        TesCase_LogResult(**data["testcase_result"]["todo"]["add_folder"]["fail"])

def categories():
    categories_name = data["title"] + date_time

    Commands.Wait20s_ClickElement(data["todo"]["manage_categories"])
    Commands.Wait20s_ClickElement(data["todo"]["manage_categories_add"])
    
    Commands.Wait20s_InputElement(data["todo"]["manage_categories_input"], categories_name)
    
    Commands.Wait20s_ClickElement(data["todo"]["manage_categories_save"])
    Logging("=> Add manage categories")
    
    try:
        Logging("** Check categories have been create!!")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["todo"]["manage_categories_name"]  % str(categories_name))))
        Logging("=> Categories have been create")
        TesCase_LogResult(**data["testcase_result"]["todo"]["add_categories"]["pass"])
    except:
        Logging("=> Categories create fail")
        TesCase_LogResult(**data["testcase_result"]["todo"]["add_categories"]["fail"])
    
    return categories_name

def edit_categories():
    categories_name_edit = data["title"] + date_time

    Commands.Wait20s_ClickElement(data["todo"]["edit"])
    Commands.Wait20s_Clear_InputElement(data["todo"]["manage_categories_input"], categories_name_edit)
    Commands.Wait20s_ClickElement(data["todo"]["manage_categories_save"])
    Logging("=> Edit manage categories")

    try:
        Logging("** Check categories have been create!!")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["todo"]["manage_categories_name"]  % str(categories_name_edit))))
        Logging("=> Categories have been edit")
        TesCase_LogResult(**data["testcase_result"]["todo"]["edit_categories"]["pass"])
    except:
        Logging("=> Categories edit fail")
        TesCase_LogResult(**data["testcase_result"]["todo"]["edit_categories"]["fail"])

    return categories_name_edit

def search():
    search_key = data["title"]
    try:
        Commands.Wait20s_EnterElement(data["todo"]["search"], search_key)
        Logging("=> Search manage categories")
        TesCase_LogResult(**data["testcase_result"]["todo"]["search_categories"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["todo"]["search_categories"]["fail"])
        pass

def delete(categories_name_edit):
    try:
        Commands.Wait20s_ClickElement(data["todo"]["manage_categories_name"]  % str(categories_name_edit))
        Logging("- Select category")
        Commands.Wait20s_ClickElement(data["todo"]["del_category"])
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["todo"]["button_ok"])
        TesCase_LogResult(**data["testcase_result"]["todo"]["delete_categories"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["todo"]["delete_categories"]["fail"])
        pass

    
def admin_execution():
    Logging(" ")
    PrintGreen("============ Test case settings admin To-Do ============")
    Logging("- Admin to-do")

    try:
        categories_name = categories()
    except:
        categories_name = None
    
    if bool(categories_name) == True:
        try:
            categories_name_edit = edit_categories()
            search()
            
            if bool(categories_name_edit) == True:
                try:
                    delete(categories_name_edit)
                except:
                    PrintRed(">> Can't continue execution")
                    pass   
            else: 
                PrintRed("=> Categories can't delete")
                TesCase_LogResult(**data["testcase_result"]["todo"]["delete_categories"]["fail"])
        except:
            categories_name_edit = None
    else:
        PrintRed("=> Categories can't create")
        TesCase_LogResult(**data["testcase_result"]["todo"]["add_categories"]["fail"])





    
    