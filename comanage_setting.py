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

from framework_sample import *
from NQ_login_function import local_path, driver, data, ValidateFailResultAndSystem, Logging, TesCase_LogResult#, TestlinkResult_Fail, TestlinkResult_Pass

#chrome_path = os.path.dirname(Path(__file__).absolute())+"\\chromedriver.exe"

n = random.randint(1,3000)
m = random.randint(3000,6000)
date_time = datetime.now().strftime("%Y/%m/%d, %H:%M:%S")

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
        element = Waits.Wait20s_ElementLoaded(data["co-manage"]["pull_the_scroll_bar"])
        element.location_once_scrolled_into_view
        time.sleep(3)
        element_1 = Waits.Wait20s_ElementLoaded(data["co-manage"]["pull_the_scroll_bar_1"])
        element_1.location_once_scrolled_into_view
        time.sleep(2)
        admin = Waits.Wait20s_ElementLoaded(data["co-manage"]["pull_the_scroll_bar"])
        if admin.is_displayed():
            Logging("- Account admin")
            admin.click()
            Logging("- Click setting admin")
            admin_execution()
    except:
        Logging("- Account user")

def work_type():
    text_worktype = data["title"] + date_time

    try:
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["worktype"]["work_type"])
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["worktype"]["button_add"])
        Commands.Wait20s_InputElement(data["co-manage"]["admin"]["worktype"]["input"], text_worktype)
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["worktype"]["button_save"])
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
        Commands.Wait20s_ClickElement("//*[@id='ngw.projectnew.adminWorkType']//table//tr[contains(., '" + text_worktype + "')]//following-sibling::td//following-sibling::a//i")
        Logging("- Select work type to delete")
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["worktype"]["button_OK"])
        Logging("=> Delete work type success")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_work_type"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_work_type"]["fail"])
        pass

def status_manage():
    text_status = data["title"] + date_time
    description = data["title"] + date_time

    try:
        f = driver.find_element_by_xpath(data["co-manage"]["admin"]["status"]["manage_status"])
        driver.execute_script("arguments[0].scrollIntoView();",f)    

        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["status"]["manage_status"])
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["status"]["button_add"])
        Logging("- Add new status")
        Commands.Wait20s_InputElement(data["co-manage"]["admin"]["status"]["input_name"], text_status)
        Commands.Wait20s_InputElement(data["co-manage"]["admin"]["status"]["input_description"], description)
        Logging("- Input name - description")

        ''' random option list '''
        options_list = ["To-Do", "In Progress", "Done"]

        sel = Select(driver.find_element_by_xpath(data["co-manage"]["admin"]["status"]["select_option"]))
        sel.select_by_visible_text(random.choice(options_list))
        Logging("- Select category in random option list")
        time.sleep(2)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["status"]["button_save"])
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
        Commands.Wait20s_ClickElement("//*[@id='ngw.projectnew.adminStatus']//table//tr[contains(., '" + text_status + "')]//following-sibling::td//following-sibling::a//i")
        Logging("- Select status to delete")
        time.sleep(4)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["worktype"]["button_OK"])
        Logging("=> Delete status success")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_status"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_status"]["fail"])
        pass

def manage_folders():
    name = data["co-manage"]["admin"]["folder"]["name_folder"] + str(n)

    try:
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["folder"]["manage_folders"])
        time.sleep(2)
        Commands.Wait20s_InputElement(data["co-manage"]["admin"]["folder"]["input"], name)

        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["folder"]["button_save"])
        Logging("=> Create folder")
        time.sleep(2)

        ''' Check folder have create '''
        Logging("** Check folder have create **")
        manage_folders = Waits.Wait20s_ElementLoaded("//*[@id='project_setting_form']//li//a[contains(., '" + name + "')]")
        if manage_folders.is_displayed():
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
    subname = data["co-manage"]["admin"]["folder"]["name_subfolder"] + str(n)

    try:
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["folder"]["parent_folder"])
        time.sleep(2)
        Commands.Wait20s_ClickElement("//*[@id='project-folder-setting-down']//li//a[contains(., '" + name + "')]")
        time.sleep(2)
        Commands.Wait20s_InputElement(data["co-manage"]["admin"]["folder"]["input"], subname)
        time.sleep(2)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["folder"]["button_save"])
        Logging("=> Create Sub-folder")
        time.sleep(5)

        ''' Check sub-folder have create '''
        Commands.Wait20s_ClickElement("//*[@id='project_setting_form']//li//a[contains(., '" + name + "')]")
        Logging("- Select parent folder")
        time.sleep(2)

        Logging("** Check sub-folder have create **")
        subfolder = Waits.Wait20s_ElementLoaded("//*[@id='project_setting_form']//li//span//a[contains(., '" + subname + "')]")
        if subfolder.is_displayed():
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

def delete_subfolder(subname):
    # subname = data["title"] + date_time
    f = driver.find_element_by_xpath(data["co-manage"]["admin"]["status"]["manage_status"])
    driver.execute_script("arguments[0].scrollIntoView();",f) 
    time.sleep(5)
    try:
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["folder"]["manage_folders"])
        Logging("** Delete sub folder")
        # Commands.Wait20s_ClickElement("//*[@id='project_setting_form']//span//a[contains(., '" + name + "')]")
        # Logging("- Select folder")
        Commands.Wait20s_ClickElement("//*[@id='project_setting_form']//li//span//a[contains(., '" + subname + "')]")
        Logging("- Select sub folder")
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["folder"]["button_delete"])
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["folder"]["button_del"])
        Logging("=> Delete sub folder success")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_subfolder"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_subfolder"]["fail"])
        pass

def delete_folder(name):
    try:
        Logging("** Delete folder")
        Commands.Wait20s_ClickElement("//*[@id='project_setting_form']//span//a[contains(., '" + name + "')]")
        time.sleep(2)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["folder"]["button_delete"])
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["folder"]["button_del"])
        Logging("=> Delete folder success")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_folder"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_folder"]["fail"])
        pass
    
    return name

def create_project():
    name_project = data["title"] + date_time

    a = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["co-manage"]["admin"]["project_list"]["list"])))
    a.location_once_scrolled_into_view
    Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["list"])
    time.sleep(1)
    try:
        Logging("** Create new project")
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["list"])
        Logging("- Select list project")
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["button_add"])
        Logging("- Select add button")
        time.sleep(3)
        Commands.Wait20s_InputElement(data["co-manage"]["admin"]["project_list"]["name_input"], name_project)
        Logging("- Input name project")
        time.sleep(3)

        ''' Select project '''
        ''' Kanban Project '''
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["kanban"])
        Logging("- Select Kanban Project")
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["button_save"])
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

    '''edit = Commands.Wait20s_ClickElement("//*[@id='ngw.projectnew.project']//strong[contains(., '" + name_project + "')]")
    edit.clear()
    time.sleep(3)
    edit.send_keys(name_project_edit)
    Logging("- Edit name project")
    time.sleep(3)
    Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["button_check_save"])
    Logging("- Save name edit")
    time.sleep(3)'''
    try:
        leader_name = data["co-manage"]["admin"]["project_list"]["name_org"]

        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["select_org_leader"])
        Logging("- Select leader")
        time.sleep(3)
        Commands.Wait20s_EnterElement(data["co-manage"]["admin"]["project_list"]["input_org_leader"], leader_name)
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["org_leader"])
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["add_leader"])
        Logging("- Select add leader")
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["save_leader"])
        Logging("- Save leader")
        time.sleep(3)
    except:
        pass

    Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["select_org_user"])
    Logging("- Select Participant(s)")
    time.sleep(3)
    Commands.Wait20s_EnterElement(data["co-manage"]["admin"]["project_list"]["input_org_user"], leader_name)
    time.sleep(3)

    try:
        user = Waits.Wait20s_ElementLoaded(data["co-manage"]["admin"]["project_list"]["org_user_1"])
        if user.is_displayed():
            Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["org_user_1"])
            time.sleep(3)
            Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["org_user_2"])
            time.sleep(3)
            Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["org_user_3"])
            time.sleep(3)
            Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["add_user"])
            Logging("- Select add Participant(s)")
            time.sleep(3)
            Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["save_user"])
            Logging("- Save Participant(s)")
            time.sleep(3)
        else:
            Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["select_org_user"])
            Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["org_user_1"])
            time.sleep(3)
            Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["org_user_2"])
            time.sleep(3)
            Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["org_user_3"])
            time.sleep(3)
            Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["add_user"])
            Logging("- Select add Participant(s)")
            time.sleep(3)
            Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["save_user"])
            Logging("- Save Participant(s)")
            time.sleep(3)
    except:
        Logging("=> Can't select Participant(s)")
        pass
    
    try:
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["select_org_cc"])
        Logging("- Select CC")
        time.sleep(3)
        Commands.Wait20s_EnterElement(data["co-manage"]["admin"]["project_list"]["input_org_cc"], leader_name)
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["org_cc"])
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["add_cc"])
        Logging("- Select add CC")
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["save_cc"])
        Logging("- Save CC")
        time.sleep(3)
    except:
        pass

    try:
        Logging("** Change stauts of project")
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["status"])
        Logging("- Select status")
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["status_dropdown"])
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
    # subname = data["title"] + date_time
    try:
        Logging("** Move project")
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["move"])
        Logging("- Select move project")
        time.sleep(3)
        Commands.Wait20s_ClickElement("//li//span//a[contains(., '" + name + "')]")
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["save_move"])
        Logging("=> Move project success")
        TesCase_LogResult(**data["testcase_result"]["comanage"]["move_project"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["comanage"]["move_project"]["fail"])
        pass

def delete_project():
    password_key = data["co-manage"]["admin"]["project_list"]["pass"]
    try:
        delete_button = Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["delete_project"])
        if delete_button.is_displayed():
            try:
                Logging("** Delete project")
                Commands.Wait20s_InputElement(data["co-manage"]["admin"]["project_list"]["input_pass"], password_key)
                Logging("- Input password")
                time.sleep(3)
                Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["button_delete"])
                Logging("=> Delete success")
                TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_project"]["pass"])
                time.sleep(5)
            except:
                TesCase_LogResult(**data["testcase_result"]["comanage"]["delete_project"]["fail"])
                pass
        else:
            try:
                Logging("** Delete project")
                Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["delete_icon"])
                Logging("- Delete project")
                time.sleep(3)
                Commands.Wait20s_InputElement(data["co-manage"]["admin"]["project_list"]["input_pass"], password_key)
                Logging("- Input password")
                time.sleep(3)
                Commands.Wait20s_ClickElement(data["co-manage"]["admin"]["project_list"]["button_delete"])
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
    # name_project_edit = data["title"] + date_time

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
        f = driver.find_element_by_xpath(data["co-manage"]["admin"]["status"]["manage_status"])
        driver.execute_script("arguments[0].scrollIntoView();",f) 
        
        manage_folder = Waits.Wait20s_ElementLoaded(data["co-manage"]["admin"]["folder"]["manage_folders"])
        if manage_folder.is_displayed():
            try:
                name = manage_folders()
            except:
                name = None

            if bool(name) == True:
                try:
                    subname = sub_folder(name)
                except:
                    Logging(">> Can't continue execution")
                    pass

                try:
                    delete_subfolder(subname)
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

    

    

