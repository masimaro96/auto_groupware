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


def todo(domain_name):
    driver.get(domain_name + "todo/setting/setting/")

    Logging(" ")
    Logging('============ Menu To-Do ============')

    ''' Access to page To-Do '''
    Logging("- Access menu")
    time.sleep(5)
    setting_execution()
    try:
        time.sleep(5)
        admin = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, data["todo"]["admin"])))
        if admin.is_displayed():
            Logging("- Account admin")
            admin.click()
            time.sleep(3)
            admin_execution()
    except:
        Logging("=> Account user")

def add_folder():
    name_folder = data["todo"]["name_folder_1"] + str(n)

    ''' Add folder '''
    Wait10s_ClickElement(data["todo"]["add_button"])
    Logging("- Select button add folder")
    time.sleep(5)
    Wait10s_EnterElement(data["todo"]["add_folder_name"], name_folder)
    Logging("- Input name folder")
    time.sleep(5)

    ''' Check button save have work '''
    try:
        Wait10s_ClickElement(data["todo"]["save_button"])
        time.sleep(5)
        Logging("=> Save success")

        ''' Check folder to-do have create '''
        Logging("** Check folder have save yet!!")
        folder_todo = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.todo.list.setting']//span//a[contains(., '" + name_folder + "')]")))
        if folder_todo.is_displayed():
            Logging("=> Folder have add successfully")
            TesCase_LogResult(**data["testcase_result"]["todo"]["add_folder"]["pass"])
        else:
            Logging("=> Folder have add fail")
            TesCase_LogResult(**data["testcase_result"]["todo"]["add_folder"]["fail"])
            ValidateFailResultAndSystem("<div>[Todo]Folder have add fail </div>")
    except WebDriverException:
        Logging("=> Save fail")
    time.sleep(5)

    return name_folder

def delete_folder(name_folder):   
    try:
        Wait10s_ClickElement("//*[@id='ngw.todo.list.setting']//span//a[contains(., '" + name_folder + "')]")
        Logging("- Select folder")
        Wait10s_ClickElement(data["todo"]["del_button"])
        time.sleep(5)
        Wait10s_ClickElement(data["todo"]["del_button_1"])
        Logging("=> Delete folder")
        TesCase_LogResult(**data["testcase_result"]["todo"]["delete_folder"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["todo"]["delete_folder"]["fail"])
        pass

def setting_execution():
    Wait10s_ClickElement(data["todo"]["settings_todo"])
    time.sleep(5)

    page_title = driver.find_element_by_xpath("//*[@id='ngw.todo.list.setting']//span")
    if page_title.text == 'Settings':
        Logging("- Click settings success")
    else:
        Logging("- Click settings fail")
    time.sleep(5)
    Logging(" ")
    Logging("============ Test case settings To-Do ============")

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
        Logging("=> Add folder todo fail")
        TesCase_LogResult(**data["testcase_result"]["todo"]["add_folder"]["fail"])

def categories():
    categories_name = data["todo"]["manage_categories_name"] + str(n)

    Wait10s_ClickElement(data["todo"]["manage_categories"])
    Wait10s_ClickElement(data["todo"]["manage_categories_add"])
    time.sleep(2)
    Wait10s_EnterElement(data["todo"]["manage_categories_input"], categories_name)
    time.sleep(2)
    Wait10s_ClickElement(data["todo"]["manage_categories_save"])
    Logging("=> Add manage categories")
    time.sleep(5)

    Logging("** Check categories have been create!!")
    categories_todo = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.todo.category']//table//tr[contains(., '" + categories_name + "')]")))
    if categories_todo.is_displayed():
        Logging("=> Categories have been create")
        TesCase_LogResult(**data["testcase_result"]["todo"]["add_categories"]["pass"])
    else:
        Logging("=> Categories create fail")
        TesCase_LogResult(**data["testcase_result"]["todo"]["add_categories"]["fail"])
        ValidateFailResultAndSystem("<div>[Todo]Categories have been create fail</div>")
    
    return categories_name

def edit_categories(categories_name):
    categories_name_edit = data["todo"]["manage_categories_name_edit"] + str(m)

    Wait10s_ClickElement(data["todo"]["edit"])
    time.sleep(5)
    Wait10s_Clear_InputElement(data["todo"]["manage_categories_input"], categories_name_edit)
    time.sleep(2)
    Wait10s_ClickElement(data["todo"]["manage_categories_save"])
    Logging("=> Edit manage categories")
    time.sleep(2)

    Logging("** Check categories have been create!!")
    categories_todo_edit = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.todo.category']//table//tr[contains(., '" + categories_name_edit + "')]")))
    if categories_todo_edit.is_displayed():
        Logging("=> Categories have been edit")
        TesCase_LogResult(**data["testcase_result"]["todo"]["edit_categories"]["pass"])
    else:
        Logging("=> Categories edit fail")
        TesCase_LogResult(**data["testcase_result"]["todo"]["edit_categories"]["fail"])
        ValidateFailResultAndSystem("<div>[Todo]Categories have been edit fail </div>")

    return categories_name_edit

def search():
    search_key = data["todo"]["name_search"]
    try:
        Wait10s_EnterElement(data["todo"]["search"], search_key)
        Logging("=> Search manage categories")
        time.sleep(3)
        TesCase_LogResult(**data["testcase_result"]["todo"]["search_categories"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["todo"]["search_categories"]["fail"])
        pass

def delete(categories_name_edit):
    try:
        Wait10s_ElementLoaded("//*[@id='ngw.todo.category']//table//tr[contains(., '" + categories_name_edit + "')]")
        
        Wait10s_ClickElement(data["todo"]["del_category"])
        time.sleep(2)
        Wait10s_ClickElement(data["todo"]["button_ok"])
        TesCase_LogResult(**data["testcase_result"]["todo"]["delete_categories"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["todo"]["delete_categories"]["fail"])
        pass

    
def admin_execution():
    Logging(" ")
    Logging("============ Test case settings admin To-Do ============")
    Logging("- Admin to-do")

    try:
        categories_name = categories()
    except:
        categories_name = None
    
    if bool(categories_name) == True:
        try:
            categories_name_edit = edit_categories(categories_name)
            search()
            
            if bool(categories_name_edit) == True:
                try:
                    delete(categories_name_edit)
                except:
                    Logging(">> Can't continue execution")
                    pass   
            else: 
                Logging("=> Categories can't delete")
                TesCase_LogResult(**data["testcase_result"]["todo"]["delete_categories"]["fail"])
        except:
            categories_name_edit = None
    else:
        Logging("=> Categories can't create")
        TesCase_LogResult(**data["testcase_result"]["todo"]["add_categories"]["fail"])





    
    