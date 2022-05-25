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
from NQ_login_function import driver, data, ValidateFailResultAndSystem, Logging, TesCase_LogResult, WaitElementLoaded

n = random.randint(1,3000)
m = random.randint(3000,6000)

def archive(domain_name):
    driver.get(domain_name + "archive/search/detail/")

    Logging(" ")
    Logging('============ Menu Archive ============')
    settings_execution()
    try:
        admin_user = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["settings_admin"])))
        if admin_user.is_displayed():
            Logging("- Account admin")
            time.sleep(5)
            admin_user.click()
            Logging("- Admin archive")
            admin_execution()
    except WebDriverException:
        Logging("=> Account user") 
    

def settings_execution():
    ''' Access menu -> click settings '''
    # driver.find_element_by_xpath(data["archive"]["setting"]["archive"]).click()
    Logging("- Access menu")
    time.sleep(2)
    Wait10s_ClickElement(data["archive"]["setting"]["settings"])
    Logging("- Click setting")
    time.sleep(5)

    Logging(" ")
    Logging("============ Test case settings Archive ============")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["setting"]["manage_myarchive"]))).click()
    Logging("- Click manage my archive")
    time.sleep(2)

    try:
        password = input_password()
    except:
        password = None

    try:
        name_folder = add_folder()
    except:
        name_folder = None

    if bool(name_folder) == True:
        try:
            name_folder_edit = edit_folder(name_folder)
        except:
            Logging(">> Can't countinue execution")
            pass
        
        delete_folder(name_folder_edit)
       
    else:
        Logging("=> Create folder fail")
        TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder"]["fail"])

def input_password():
    password = data["archive"]["setting"]["input_pass"]
    try:
        driver.find_element_by_xpath(data["archive"]["setting"]["pass"]).send_keys(password)
        driver.find_element_by_xpath(data["archive"]["setting"]["button_submit"]).click()
        Logging("** Input password **")
        time.sleep(2)

        Logging("** Check input password success")
        title_page = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "set-my-arch")))
        if title_page.is_displayed():
            Logging("=> Input password success")
        else:
            Logging("=> Input password fail")
            ValidateFailResultAndSystem("<div>[Archive]Input password fail </div>")
        time.sleep(5)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "set-my-arch")))
        time.sleep(2)
    except:
        False
    return password

def add_folder():
    name_folder = data["archive"]["setting"]["folder_name"] + str(n)
    ''' Create folder in my archive '''
    try:
        Logging(" ")
        Logging("** Create folder in my archive")
        driver.find_element_by_xpath(data["archive"]["setting"]["folder"]).click()
        Logging("- Select folder")
        time.sleep(4)
        driver.find_element_by_xpath(data["archive"]["setting"]["add_folder"]).click()
        Logging("- Select my archive")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["setting"]["enabled"]).click()
        Logging("- Select Permission")
        time.sleep(2)
        namefolder_input = driver.find_element_by_xpath(data["archive"]["setting"]["input_folder"])
        namefolder_input.send_keys(name_folder)
        Logging("- Input name folder")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["setting"]["save"]).click()
        Logging("=> Create folder in my archive success")
        time.sleep(5)
    except:
        pass
    
    try:
        Logging(" ")
        Logging("** Check folder have save **")
        check_folder = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[@id='set-my-arch']//span[contains(., '" + name_folder + "')]")))
        if check_folder.is_displayed:
            Logging("=> Create folder success")
            TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder"]["pass"])
        else:
            Logging("=> Create folder fail")
            TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder"]["fail"])
            ValidateFailResultAndSystem("<div>[Archive]Create folder in settings fail </div>")
        time.sleep(5)
    except:
        Logging("=> Create folder fail")
        TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder"]["fail"])
        pass
    return name_folder

def edit_folder(name_folder):
    name_folder_edit = data["archive"]["setting"]["folder_name_edit"] + str(n)
    try:
        Logging(" ")
        Logging("** Edit folder - name - permission")
        folder = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[@id='set-my-arch']//span[contains(., '" + name_folder + "')]")))
        folder.click()
        driver.find_element_by_xpath(data["archive"]["setting"]["modify"]).click()
        Logging("- Select edit folder")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["setting"]["disabled"]).click()
        Logging("- Edit Permission")
        time.sleep(2)
        namefolder_input = driver.find_element_by_xpath(data["archive"]["setting"]["input_folder"])
        namefolder_input.clear()
        time.sleep(2)
        namefolder_input.send_keys(name_folder_edit)
        Logging("- Input edit name")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["setting"]["save"]).click()
        Logging("=> Edit folder in my archive")
        TesCase_LogResult(**data["testcase_result"]["archive"]["edit_folder"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["edit_folder"]["fail"])
        pass
    
    return name_folder_edit

def delete_folder(name_folder_edit):
    try:
        Logging(" ")
        Logging("** Delete folder have edit")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='set-my-arch']//span[contains(., '" + name_folder_edit + "')]"))).click()
        driver.find_element_by_xpath(data["archive"]["setting"]["delete"]).click()
        time.sleep(2)
        Logging("- Select folder to delete - Click button delete")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["setting"]["button_ok"]).click()
        Logging("=> Delete folder success")
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_folder"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_folder"]["fail"])
        pass

def admin_execution():
    Logging("- Click Admin archive")

    Logging(" ")
    Logging("============ Test case Admin Archive ============")

    Logging(" ")
    Logging("** Manage Company Archive") 
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["ManageCompany"]["Manage_Company"]))).click()
    Logging("- Click Manage Company")
    time.sleep(5)

    ''' Add folder public first 
    -> Create sub folder in folder public
    -> Add authorized for sub folder
    -> Delete sub folder
    -> Delete public folder'''
    try:
        name_folder_admin = folder_public()
    except:
        name_folder_admin = None

    if bool(name_folder_admin) == True:
        try:
            name_subfolder = sub_folder(name_folder_admin)
        except:
            Logging(">> Can't countinue execution")
            pass

        try:
            add_authorized_sub_folder(name_subfolder)
        except:
            Logging(">> Can't countinue execution")
            pass

        try:
            delete_sub_folder(name_subfolder)
        except:
            Logging(">> Can't countinue execution")
            pass

        try:
            delete_public_folder(name_folder_admin)
        except:
            Logging(">> Can't countinue execution")
            pass
    else:
        Logging("=> Create folder public fail")

    ''' Add folder private first 
    -> Delete private folder'''
    try:
        name_folder_private = folder_private()
    except:
        name_folder_private = None

    if bool(name_folder_private) == True:
        try:
            delete_private_folder(name_folder_private)
        except:
            Logging(">> Can't countinue execution")
            pass
    else:
        Logging("=> Create folder private fail")

    ''' Add Archive Manager 
    -> Delete Archive Manager'''
    try:
       archive_manager()
    except:
        pass

    try:  
        delete_manager()
    except:
        pass
        Logging("=> Can't add Archive Manager")

    try:
        backup()
    except:
        Logging(">> Can't countinue execution")
        pass

    try:
        transfer_company_approval()
    except:
        Logging(">> Can't countinue execution")
        pass

    try:
        transfer_company_board()
    except:
        Logging(">> Can't countinue execution")
        pass

    try:
        transfer_company_task()
    except:
        Logging(">> Can't countinue execution")
        pass

def folder_public():
    name_folder_admin = data["archive"]["admin"]["ManageCompany"]["name_folder"] + str(n)
    ''' create folder public '''
    try:
        Logging(" ")
        Logging("-> Create folder public")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["loading_dialog"])))

        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["parent_folder"]).click()
        Logging("- Select folder to create")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["add_folder_admin"]).click()
        Logging("- Create folder")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["input_name_folder"]).send_keys(name_folder_admin)
        Logging("- Public folder")
        Logging("- Input name folder")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["save_folder"]).click()
        Logging("- Save public folder")
        time.sleep(3)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["button_close"]).click()
        Logging("=> Confirm create folder public")
        time.sleep(5)
    except:
        pass

    ''' Check folder public have create '''
    try:
        Logging(" ")
        Logging("** Check folder public have save **")
        check_folder_public = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[@id='archive-tab-content']//span[contains(., '" + name_folder_admin + "')]")))
        if check_folder_public.is_displayed:
            Logging("=> Create folder public success")
            TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder_public"]["pass"])
        else:
            Logging("=> Create folder public fail")
            TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder_public"]["fail"])
            ValidateFailResultAndSystem("<div>[Archive]Create folder public fail </div>")
        time.sleep(5)
    except:
        pass
    
    return name_folder_admin

def folder_private():
    name_folder_private = data["archive"]["admin"]["ManageCompany"]["name_folderprivate"] + str(m)

    ''' Create folder private -> Permission Disable '''
    try:
        Logging(" ")
        Logging("-> Create folder private")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["loading_dialog"])))
        
        time.sleep(2)

        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["parent_folder"]).click()
        Logging("- Select folder to create")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["add_folder_admin"]).click()
        Logging("- Create folder")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["private"]).click()
        Logging("- Private folder")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["disabled"]).click()
        Logging("- Permission: disabled")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["input_name_folder"]).send_keys(name_folder_private)
        Logging("- Input name folder")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["save_folder"]).click()
        Logging("- Save private folder")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["button_close"]).click()
        Logging("Create folder private")
        time.sleep(5)
    except:
        pass
    
    ''' Check folder private have create '''
    try:
        Logging(" ")
        Logging("**Check folder private have save**")
        check_folder_private = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[@id='archive-tab-content']//span[contains(., '" + name_folder_private + "')]")))
        if check_folder_private.is_displayed:
            Logging("=> Create folder private success")
            TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder_privte"]["pass"])
        else:
            Logging("=> Create folder private fail")
            TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder_privte"]["fail"])
            ValidateFailResultAndSystem("<div>[Archive]Create folder private fail </div>")
        time.sleep(5)
    except:
        pass
    
    return name_folder_private
    
def sub_folder(name_folder_admin):
    name_subfolder = data["archive"]["admin"]["ManageCompany"]["subfolder"] + str(m)

    ''' Create sub-folder '''
    try:
        Logging(" ")
        Logging("** Create sub folder in public folder")
        driver.find_element_by_tag_name("body").send_keys(Keys.END)
        time.sleep(5)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='archive-tab-content']//span[contains(., '" + name_folder_admin + "')]"))).click()
        Logging("- Select puclic to create sub folder")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["add_subfolder"]).click()
        Logging("- Add sub folder")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["input_subfolder"]).send_keys(name_subfolder)
        Logging("- Input name sub folder")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["save_subfolder"]).click()
        Logging("- Save sub folder")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["button_close"]).click()
        Logging("=> Create sub folder admin")
        time.sleep(5)
    except:
        pass

    ''' Check sub folder have create '''
    try:
        Logging(" ")
        Logging("** Check sub folder have create")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='archive-tab-content']//span[contains(., '" + name_folder_admin + "')]"))).click()
        time.sleep(2)
        Logging("**Check subfolder have save**")
        sub_folder = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[@id='archive-tab-content']//span[contains(., '" + name_subfolder + "')]")))
        if sub_folder.is_displayed:
            Logging("=> Create subfolder success")
            TesCase_LogResult(**data["testcase_result"]["archive"]["add_subfolder"]["pass"])
        else:
            Logging("=> Create subfolder fail")
            TesCase_LogResult(**data["testcase_result"]["archive"]["add_subfolder"]["fail"])
            ValidateFailResultAndSystem("<div>[Archive]Create sub folder in folder public fail </div>")
        time.sleep(5)
    except:
        pass

    return name_subfolder

def add_authorized_sub_folder(name_subfolder):
    ''' Add authorized_Dept for subfoder '''
    try:
        Logging(" ")
        Logging("** Add authorized Dept for subfoder")
        Logging("(Selected Dept. Only + Selected Folders Only)")
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[@id='archive-tab-content']//span[contains(., '" + name_subfolder + "')]"))).click()
        Logging("- Select sub folder")
        time.sleep(5)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["Select_authorized_Dept."]).click()
        Logging("- Select authorized Dept")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["Dept."]).click()
        Logging("- Select Dept.")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["add_Dept."]).click()
        Logging("=> Save settings")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["button_close"]).click()
        Logging("=> Confirm close popup")
        TesCase_LogResult(**data["testcase_result"]["archive"]["authorized_Dept"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["authorized_Dept"]["fail"])
        pass

    try:
        Logging(" ")
        Logging("(Include All Sub-Dept.(s) + Include All Sub-Folders)")
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["Include_Sub_Dept."]).click()
        Logging("- Select Dept.settings: Include All Sub-Dept.(s)")
        time.sleep(5)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='archive-tab-content']//span[contains(., '" + name_subfolder + "')]"))).click()
        time.sleep(3)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["Select_authorized_Dept."]).click()
        Logging("- Select Authorized Dept.")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["Dept."]).click()
        Logging("- Select Dept.")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["Include_Sub_folder."]).click()
        Logging("- Select Dept. Settings: Include All Sub-Folders")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["add_Dept."]).click()
        Logging("=> Save settings")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["button_close"]).click()
        Logging("=> Confirm close popup")
        TesCase_LogResult(**data["testcase_result"]["archive"]["Dept_settings"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["Dept_settings"]["fail"])
        pass

def delete_sub_folder(name_subfolder):
    ''' Delete sub folder '''
    try:
        Logging(" ")
        Logging("** Delete sub folder")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='archive-tab-content']//span[contains(., '" + name_subfolder + "')]"))).click()
        time.sleep(3)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["check_all"]).click()
        time.sleep(3)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["del_all"]).click()
        time.sleep(3)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["button_ok"]).click()
        time.sleep(3)
        Logging("- Select sub folder")
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["del_folder"]).click()
        time.sleep(3)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["button_OK"]).click()
        Logging("=> Delete sub folder")
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_subfolder"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_subfolder"]["fail"])
        pass

def delete_public_folder(name_folder_admin):
    try:
        ''' Delete folder '''
        Logging(" ")
        Logging("** Delete public folder")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='archive-tab-content']//span[contains(., '" + name_folder_admin + "')]"))).click()
        time.sleep(3)
        Logging("- Select parent folder")
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["del_folder"]).click()
        time.sleep(3)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["button_OK"]).click()
        Logging("=> Delete public folder")
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_folder_public"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_folder_public"]["fail"])
        pass

def delete_private_folder(name_folder_private):
    try:
        Logging("** Delete private folder")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='archive-tab-content']//span[contains(., '" + name_folder_private + "')]"))).click()
        time.sleep(3)
        Logging("- Select parent folder")
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["del_folder"]).click()
        time.sleep(3)
        driver.find_element_by_xpath(data["archive"]["admin"]["ManageCompany"]["button_OK"]).click()
        Logging("=> Delete private folder")
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_folder_privte"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_folder_privte"]["fail"])
        pass

def archive_manager():
    name_manager = data["archive"]["admin"]["ArchiveManager"]["Name"]
    ''' Add Archive Manager '''
    Logging(" ")
    Logging("** Add Archive Manager")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["ArchiveManager"]["Company_Archive"]))).click()
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["ArchiveManager"]["Archive_Manager"]))).click()
    Logging("Access Archive_Manager")
    time.sleep(5)

    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["ArchiveManager"]["button_add"]))).click()
        Logging("- Add manager archive")
        time.sleep(5)

        select_user = driver.find_element_by_xpath(data["archive"]["admin"]["ArchiveManager"]["input_name"])
        time.sleep(2)
        select_user.send_keys(name_manager)
        time.sleep(2)
        select_user.send_keys(Keys.ENTER)
        Logging("- Input name manager archive")
        time.sleep(2)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["ArchiveManager"]["select_user"]))).click()
        Logging("- Select user")
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ArchiveManager"]["button_save"]).click()
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["ArchiveManager"]["button_OK."]).click()
        Logging("=> Add Archive Manager")
        TesCase_LogResult(**data["testcase_result"]["archive"]["archive_manager"]["pass"])
        time.sleep(5)
    except:
        Logging("=> Can't add Archive Manager")
        TesCase_LogResult(**data["testcase_result"]["archive"]["archive_manager"]["fail"])
        pass

def delete_manager():
    ''' Delete user manager '''
    try:
        time.sleep(5)
        Logging("** Delete user manager")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["ArchiveManager"]["pick_user"]))).click()
        Logging("- Select user to delete")
        time.sleep(5)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["ArchiveManager"]["button_delete"]))).click()
        Logging("- Click button delete")
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["ArchiveManager"]["close_popup"]))).click()
        Logging("=> Delete Archive Manager")
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_archive_manager"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_archive_manager"]["fail"])
        pass

def backup():
    name_backup = data["archive"]["admin"]["backup_folder"]["name"] + str(n)
    ''' Backup folder '''
    Logging(" ")
    Logging("** Backup folder")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["backup_folder"]["menu"]))).click()
    Logging("- Access Backup folder")
    
    ''' Back up file need edit date backup '''    
    try:    
        driver.find_element_by_xpath(data["archive"]["admin"]["backup_folder"]["button_add"]).click()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["backup_folder"]["input_name"]))).send_keys(name_backup)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["backup_folder"]["selectfolder"]))).click()
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["backup_folder"]["choosefolder"]))).click()
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["backup_folder"]["click_date_range_from"]))).click()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["backup_folder"]["select_date_range_from"]))).click()
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["backup_folder"]["click_date_range_to"]))).click()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["backup_folder"]["select_date_range_to"]))).click()
        time.sleep(2)
        driver.find_element_by_xpath(data["archive"]["admin"]["backup_folder"]["save_backup"]).click()   
        time.sleep(2)   
        driver.find_element_by_xpath(data["archive"]["admin"]["backup_folder"]["close_popup"]).click()   
        Logging("Backup folder success")
        TesCase_LogResult(**data["testcase_result"]["archive"]["backup"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["backup"]["fail"])
        pass

def transfer_company_approval():
    ''' Transfer compnay -> Approval '''
    Logging(" ")
    Logging("** Transfer compnay -> Approval")
    driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["transfer_company"]).click()
    Logging("- Transfer company")
    time.sleep(5)
    try:
        ''' select approval menu '''
        driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_approval"]["approval"]).click()
        Logging("- Transfer company - Approval")
        time.sleep(5)
        sel = Select(driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_approval"]["folders_option"]))
        sel.select_by_visible_text('Selenium')        
        Logging("- Select Dept to transfer")
        time.sleep(2)

        driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_approval"]["location_option"]).click()
        time.sleep(3)
        driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_approval"]["location"]).click()
        Logging("- Select location to transfer")
        time.sleep(2)

        ''' select language '''       
        sel = Select(driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_approval"]["language_option"]))
        sel.select_by_visible_text('English')
        Logging("- Select language")
        time.sleep(2)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["transfer"]["menu_approval"]["button_OK"]))).click()
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["transfer"]["menu_approval"]["close_popup"]))).click()
        Logging("=> Transfer company - Approval")
        time.sleep(5)
        TesCase_LogResult(**data["testcase_result"]["archive"]["transfer_approval"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["transfer_approval"]["fail"])
        pass

def transfer_company_board():
    ''' Transfer compnay -> Company Board ''' 
    try:
        Logging(" ")
        Logging("** Transfer compnay -> Company Board")
        driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["transfer_company"]).click()
        time.sleep(2)

        ''' select Company Board menu '''
        driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_board"]["company_board"]).click()
        Logging("- Transfer company - Company Board")
        time.sleep(2)
        sel = Select(driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_board"]["board_option"]))
        sel.select_by_visible_text('Company Board')        
        Logging("- Select Company Board transfer")

        time.sleep(2)

        driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_board"]["location_option"]).click()
        time.sleep(3)
        driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_board"]["location"]).click()
        Logging("- Select location to transfer")
        time.sleep(2)

        ''' select language '''       
        sel = Select(driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_board"]["language_option"]))
        sel.select_by_visible_text('English')
        Logging("- Select language")
        time.sleep(2)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["transfer"]["menu_board"]["button_OK"]))).click()
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["transfer"]["menu_board"]["close_popup"]))).click()
        Logging("=> Transfer company - Company Board")
        TesCase_LogResult(**data["testcase_result"]["archive"]["transfer_board"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["transfer_board"]["fail"])
        pass

def transfer_company_task():
    ''' Transfer compnay -> work dairy '''
    try:
        Logging(" ")
        Logging("** Transfer compnay -> work dairy")

        sel = Select(driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_task"]["folders_option"]))
        sel.select_by_visible_text('Selenium')        
        Logging("- Select Dept to transfer")
        time.sleep(2)

        driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_task"]["location_option"]).click()
        time.sleep(3)
        driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_task"]["location"]).click()
        Logging("- Select location to transfer")
        time.sleep(2)

        ''' select language '''       
        sel = Select(driver.find_element_by_xpath(data["archive"]["admin"]["transfer"]["menu_task"]["language_option"]))
        sel.select_by_visible_text('English')
        Logging("- Select language")
        time.sleep(2)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["transfer"]["menu_task"]["button_OK"]))).click()
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["transfer"]["menu_task"]["close_popup"]))).click()
        Logging("=> Transfer company - work dairy")
        time.sleep(5)
        TesCase_LogResult(**data["testcase_result"]["archive"]["transfer_work_dairy"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["transfer_work_dairy"]["fail"])
        pass
