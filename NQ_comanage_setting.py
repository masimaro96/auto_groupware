import time, json, random, openpyxl
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from datetime import datetime
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from random import choice
from random import randint
from openpyxl import Workbook
import re
from sys import exit
from selenium.webdriver import ActionChains
from selenium.webdriver.common.action_chains import ActionChains

import pathlib
from pathlib import Path
import os
from sys import platform
import NQ_function

from NQ_login_function import driver, data, ValidateFailResultAndSystem, Logging, TesCase_LogResult#, TestlinkResult_Fail, TestlinkResult_Pass

#chrome_path = os.path.dirname(Path(__file__).absolute())+"\\chromedriver.exe"

n = random.randint(1,3000)
m = random.randint(3000,6000)
now = datetime.now()

def co_manage(domain_name):
    driver.get(domain_name + "projectnew/projects")
    Logging(" ")
    Logging('============ Menu Comanage ============')
    

    ''' Access menu -> click admin settings '''
    # driver.find_element_by_xpath(data["co-manage"]["comanage"]).click()
    Logging("- Access menu")
    time.sleep(10)
    #driver.execute_script("window.scrollTo(-2000, -502)") 
    

    try:
        element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["pull_the_scroll_bar"])))
        element.location_once_scrolled_into_view
        time.sleep(3)
        element_1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["pull_the_scroll_bar_1"])))
        element_1.location_once_scrolled_into_view
        time.sleep(2)
        admin = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["pull_the_scroll_bar"])))
        if admin.is_displayed():
            Logging("- Account admin")
            admin.click()
            Logging("- Click setting admin")
            admin_execution()
    except:
        Logging("- Account user")

def work_type():
    text_worktype = data["co-manage"]["admin"]["worktype"]["input_text"] + str(n)

    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["worktype"]["work_type"]))).click()
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["worktype"]["button_add"]))).click()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["worktype"]["input"]))).send_keys(text_worktype)
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["worktype"]["button_save"]))).click()
        Logging("=> Create work type")
        time.sleep(3)
    except:
        pass

    Logging("** Check work type have been create")
    work_type = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.projectnew.adminWorkType']//table//tr[contains(., '" + text_worktype + "')]")))
    if work_type.is_displayed():
        Logging("=> Work type have been create")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["work_type"]["pass"])
    else:
        Logging("=> Work type have been create fail")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["work_type"]["fail"])
        ValidateFailResultAndSystem("<div>[Comanage]Work type have been create fail </div>")
    time.sleep(5)

    return text_worktype

def delete_worktype(text_worktype):
    try:
        Logging("** Delete work type")
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.projectnew.adminWorkType']//table//tr[contains(., '" + text_worktype + "')]//following-sibling::td//following-sibling::a//i"))).click()
        Logging("- Select work type to delete")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["worktype"]["button_OK"]))).click()
        Logging("=> Delete work type success")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_work_type"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_work_type"]["fail"])
        pass

def status_manage():
    text_status = data["co-manage"]["admin"]["status"]["name"] + str(n)
    description = data["co-manage"]["admin"]["status"]["description"]

    try:
        f = driver.find_element_by_xpath(data["co-manage"]["admin"]["status"]["manage_status"])
        driver.execute_script("arguments[0].scrollIntoView();",f)    

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["status"]["manage_status"]))).click()
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["status"]["button_add"]))).click()
        Logging("- Add new status")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["status"]["input_name"]))).send_keys(text_status)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["status"]["input_description"]))).send_keys(description)
        Logging("- Input name - description")

        ''' random option list '''
        options_list = ["To-Do", "In Progress", "Done"]

        sel = Select(driver.find_element_by_xpath(data["co-manage"]["admin"]["status"]["select_option"]))
        sel.select_by_visible_text(random.choice(options_list))
        Logging("- Select category in random option list")
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["status"]["button_save"]))).click()
        Logging("=> Create status")
        time.sleep(5)
    except:
        pass
    
    try:
        Logging("** Check status have been create")
        status = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.projectnew.adminStatus']//table//tr[contains(., '" + text_status + "')]")))
        if status.is_displayed():
            Logging("=> Status have been create")
            TesCase_LogResult(**data["testcase_result"]["comanage"]["status"]["pass"])
        else:
            Logging("=> Status have been create fail")
            TesCase_LogResult(**data["testcase_result"]["comanage"]["status"]["fail"])
            ValidateFailResultAndSystem("<div>[Comanage]Status have been create fail </div>")
        time.sleep(5)
    except:
        Logging("=> Status have been create fail")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["status"]["fail"])
        pass

    return text_status

def delete_status(text_status):
    try:
        Logging("** Delete status")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.projectnew.adminStatus']//table//tr[contains(., '" + text_status + "')]//following-sibling::td//following-sibling::a//i"))).click()
        Logging("- Select status to delete")
        time.sleep(4)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["worktype"]["button_OK"]))).click()
        Logging("=> Delete status success")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_status"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_status"]["fail"])
        pass

def manage_folders():
    name = data["co-manage"]["admin"]["folder"]["name_folder"] + str(n)

    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["folder"]["manage_folders"]))).click()
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["folder"]["input"]))).send_keys(name)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["folder"]["button_save"]))).click()
        Logging("=> Create folder")
        time.sleep(2)

        ''' Check folder have create '''
        Logging("** Check folder have create **")
        manage_folders = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='project_setting_form']//li//a[contains(., '" + name + "')]")))
        if manage_folders.is_displayed:
            Logging("=> Folder have create success")
            TesCase_LogResult(**data["testcase_result"]["comanage"]["folder"]["pass"])
        else:
            Logging("=> Folder have create fail")
            TesCase_LogResult(**data["testcase_result"]["comanage"]["folder"]["fail"])
            ValidateFailResultAndSystem("<div>[Comanage]Folder have create fail </div>")
        time.sleep(5)
    except:
        Logging("=> Folder have create fail")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["folder"]["fail"])
        pass
    return name

def sub_folder(name):
    subname = data["co-manage"]["admin"]["folder"]["name_subfolder"] + str(m)

    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["folder"]["parent_folder"]))).click()
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='project-folder-setting-down']//li//a[contains(., '" + name + "')]"))).click()
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["folder"]["input"]))).send_keys(subname)
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["folder"]["button_save"]))).click()
        Logging("=> Create Sub-folder")
        time.sleep(5)

        ''' Check sub-folder have create '''
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='project_setting_form']//li//a[contains(., '" + name + "')]"))).click()
        time.sleep(2)

        Logging("** Check sub-folder have create **")
        subfolder = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='project_setting_form']//li//a[contains(., '" + subname + "')]")))    
        if subfolder.is_displayed:
            Logging("=> Sub-Folder have create success")
            TesCase_LogResult(**data["testcase_result"]["comanage"]["subfolder"]["pass"])
        else:
            Logging("=> Sub-Folder have create fail")
            TesCase_LogResult(**data["testcase_result"]["comanage"]["subfolder"]["fail"])
            ValidateFailResultAndSystem("<div>[Comanage]Sub-Folder have create fail </div>")
        time.sleep(5)
    except:
        Logging("=> Sub-Folder have create fail")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["subfolder"]["fail"])
        pass
    
    return subname

def delete_folder(name):
    subname = data["co-manage"]["admin"]["folder"]["name_subfolder"] + str(m)
    f = driver.find_element_by_xpath(data["co-manage"]["admin"]["status"]["manage_status"])
    driver.execute_script("arguments[0].scrollIntoView();",f) 
    time.sleep(5)
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["folder"]["manage_folders"]))).click()
        Logging("** Delete folder")
        WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.XPATH, "//*[@id='project_setting_form']//span//a[contains(., '" + name + "')]"))).click()
        Logging("- Select folder")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='project_setting_form']//span//a[contains(., '" + subname + "')]"))).click()
        Logging("- Select sub folder")
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["folder"]["button_delete"]))).click()
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["folder"]["button_del"]))).click()
        Logging("=> Delete sub folder success")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_subfolder"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_subfolder"]["fail"])
        pass

    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='project_setting_form']//span//a[contains(., '" + name + "')]"))).click()
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["folder"]["button_delete"]))).click()
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["folder"]["button_del"]))).click()
        Logging("=> Delete folder success")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_folder"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_folder"]["fail"])
        pass

def create_project():
    name_project = data["co-manage"]["admin"]["project_list"]["name_text"] + str(n)

    a = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["list"])))
    a.location_once_scrolled_into_view
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["list"]))).click()
    time.sleep(1)
    try:
        Logging("** Create new project")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["list"]))).click()
        Logging("- Select list project")
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["button_add"]))).click()
        Logging("- Select add button")
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["name_input"]))).send_keys(name_project)
        Logging("- Input name project")
        time.sleep(3)

        ''' Select project '''
        ''' Kanban Project '''
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["kanban"]))).click()
        Logging("- Select Kanban Project")
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["button_save"]))).click()
        Logging("- Save kanban project")
        time.sleep(3)
    except:
        pass

    ''' Check project template '''
    Logging("** Check project template")
    project_type = driver.find_element_by_xpath("//*[@id='ngw.projectnew.project']//strong[contains(.,'Kanban')]")
    if project_type.text == 'Kanban':
        Logging("=> Create kanban project success")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["project"]["pass"])
    else:
        Logging("=> Create kanban project fail")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["project"]["fail"])
    time.sleep(5)

    '''edit = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.projectnew.project']//strong[contains(., '" + name_project + "')]"))).click()
    edit.clear()
    time.sleep(3)
    edit.send_keys(name_project_edit)
    Logging("- Edit name project")
    time.sleep(3)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["button_check_save"]))).click()
    Logging("- Save name edit")
    time.sleep(3)'''
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["select_org_leader"]))).click()
        Logging("- Select leader")
        time.sleep(3)
        leader = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["input_org_leader"])))
        time.sleep(3)
        leader.send_keys(data["co-manage"]["admin"]["project_list"]["name_org"])
        time.sleep(3)
        leader.send_keys(Keys.ENTER)
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["org_leader"]))).click()
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["add_leader"]))).click()
        Logging("- Select add leader")
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["save_leader"]))).click()
        Logging("- Save leader")
        time.sleep(3)
    except:
        pass

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["select_org_user"]))).click()
    Logging("- Select Participant(s)")
    time.sleep(3)
    leader = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["input_org_user"])))
    time.sleep(3)
    leader.send_keys(data["co-manage"]["admin"]["project_list"]["name_org"])
    time.sleep(3)
    leader.send_keys(Keys.ENTER)
    time.sleep(3)

    try:
        user = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["org_user_1"])))
        if user.is_displayed():
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["org_user_1"]))).click()
            time.sleep(3)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["org_user_2"]))).click()
            time.sleep(3)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["org_user_3"]))).click()
            time.sleep(3)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["add_user"]))).click()
            Logging("- Select add Participant(s)")
            time.sleep(3)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["save_user"]))).click()
            Logging("- Save Participant(s)")
            time.sleep(3)
        else:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["select_org_user"]))).click()
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["org_user_1"]))).click()
            time.sleep(3)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["org_user_2"]))).click()
            time.sleep(3)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["org_user_3"]))).click()
            time.sleep(3)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["add_user"]))).click()
            Logging("- Select add Participant(s)")
            time.sleep(3)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["save_user"]))).click()
            Logging("- Save Participant(s)")
            time.sleep(3)
    except:
        Logging("=> Can't select Participant(s)")
        pass
    
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["select_org_cc"]))).click()
        Logging("- Select CC")
        time.sleep(3)
        leader = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["input_org_cc"])))
        time.sleep(3)
        leader.send_keys(data["co-manage"]["admin"]["project_list"]["name_org"])
        time.sleep(3)
        leader.send_keys(Keys.ENTER)
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["org_cc"]))).click()
        time.sleep(3)
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["add_cc"]))).click()
        Logging("- Select add CC")
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["save_cc"]))).click()
        Logging("- Save CC")
        time.sleep(3)
    except:
        pass

    try:
        Logging("** Change stauts of project")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["status"]))).click()
        Logging("- Select status")
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["status_dropdown"]))).click()
        Logging("- Click dropdown")
        time.sleep(3)

        options_list = ["On Hold", "In Progress", "Complete"]

        sel = driver.find_element_by_xpath(data["co-manage"]["admin"]["project_list"]["input_dropdown"])
        sel.send_keys(random.choice(options_list))
        sel.send_keys(Keys.ENTER)
        
        Logging("=> Change stauts of project: " + str(random.choice(options_list)))
        TesCase_LogResult(**data["testcase_result"]["comanage"]["change_status"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["comanage"]["change_status"]["fail"])
        pass

    return name_project

def move_project(name):
    subname = data["co-manage"]["admin"]["folder"]["name_subfolder"] + str(m)
    try:
        Logging("** Move project")
        driver.find_element_by_xpath(data["co-manage"]["admin"]["project_list"]["move"]).click()
        Logging("- Select move project")
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//li//span//a[contains(., '" + name + "')]"))).click()
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//li//span//a[contains(., '" + subname + "')]"))).click()
        Logging("- Select folder to move project")
        time.sleep(3)
        driver.find_element_by_xpath(data["co-manage"]["admin"]["project_list"]["save_move"]).click()
        Logging("=> Move project success")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["move_project"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["comanage"]["move_project"]["fail"])
        pass

def delete_project():
    try:
        delete_button = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["delete_project"])))
        if delete_button.is_displayed():
            try:
                Logging("** Delete project")
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["delete_project"]))).click()
                Logging("- Delete project")
                time.sleep(3)
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["input_pass"]))).send_keys(data["co-manage"]["admin"]["project_list"]["pass"])
                Logging("- Input password")
                time.sleep(3)
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["button_delete"]))).click()
                Logging("=> Delete success")
                TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_project"]["pass"])
                time.sleep(5)
            except:
                TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_project"]["fail"])
                pass
        else:
            try:
                Logging("** Delete project")
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["delete_icon"]))).click()
                Logging("- Delete project")
                time.sleep(3)
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["input_pass"]))).send_keys(data["co-manage"]["admin"]["project_list"]["pass"])
                Logging("- Input password")
                time.sleep(3)
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["button_delete"]))).click()
                Logging("=> Delete success")
                TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_project"]["pass"])
                time.sleep(5)
            except:
                TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_project"]["fail"])
                pass
    except:
        pass
    time.sleep(5)

def admin_execution():
    name_project_edit = data["co-manage"]["admin"]["project_list"]["name_text_edit"] + str(m)

    Logging(" ")
    Logging("============ Test case settings admin co-manage ============")

    ''' Access manage status -> input random option list '''
    try:
        text_status = status_manage()
    except:
        text_status = None

    if bool(text_status) == True:
        try:
            delete_status(text_status)
        except:
            Logging(">> Can't continue execution")
            pass
    else:
        Logging("=> Status have been create fail")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["status"]["fail"])

    ''' Access work type '''
    try:
        text_worktype = work_type()
    except:
        text_worktype = None

    if bool(text_worktype) == True:
        try:
            delete_worktype(text_worktype)
        except:
            Logging(">> Can't continue execution")
            pass
    else:
        Logging("=> Work type have been create fail")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["work_type"]["fail"])
    
    ''' Access manage folders -> input name random -> creare folder'''
    try:
        manage_folder = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["folder"]["manage_folders"])))
        if manage_folder.is_displayed():
            try:
                name = manage_folders()
            except:
                name = None

            if bool(name) == True:
                try:
                    sub_folder(name)
                except:
                    Logging(">> Can't continue execution")
                    pass
            else:
                Logging("=> Folder have create fail")
        else:
            Logging("=> Domain don't have menu Manage folders")
    except:
        Logging("=> Domain don't have menu Manage folders")
    
    ''' Create project '''
    try:
        name_project = create_project()
    except:
        name_project = None

    if bool(name_project) == True:
        try:
            move_project(name)
        except:
            Logging(">> Can't continue execution")
            pass

        try:
            delete_project()
        except:
            Logging(">> Can't continue execution")
            pass

        try:
            delete_folder(name)
        except:
            Logging(">> Can't continue execution")
            pass
    else:
        Logging("=> Can't move folder")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["folder"]["fail"])

    

    

