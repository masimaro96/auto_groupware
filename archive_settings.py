import time, json, openpyxl
import random
from selenium import webdriver
from datetime import datetime
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
date_time = datetime.now().strftime("%Y/%m/%d %H:%M:%S")

def archive(domain_name):
    driver.get(domain_name + "archive/search/detail/")

    Logging(" ")
    PrintGreen('============ Menu Archive ============')
    try:
        admin_user = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["settings_admin"])))
        if admin_user.is_displayed():
            Logging("- Account admin")
            settings_execution()
            
            admin_user.click()
            Logging("- Admin archive")
            admin_execution()
    except WebDriverException:
        Logging("=> Account user") 
        settings_execution()
    

def settings_execution():
    ''' Access menu -> click settings '''
    # driver.find_element_by_xpath(data["archive"]["setting"]["archive"]).click()
    Logging("- Access menu")
    
    Commands.Wait20s_ClickElement(data["archive"]["setting"]["settings"])
    Logging("- Click setting")
    

    Logging(" ")
    PrintGreen("============ Test case settings Archive ============")
    Commands.Wait20s_ClickElement(data["archive"]["setting"]["manage_myarchive"])
    Logging("- Click manage my archive")
    
    try:
        password_key = input_password()
    except:
        password_key = None

    if bool(password_key) == True:
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
    password_key = data["archive"]["setting"]["input_pass"]
    try:
        time.sleep(5)
        Commands.Wait20s_Clear_Click_InputElement(data["archive"]["setting"]["pass"], password_key)
        Commands.Wait20s_ClickElement(data["archive"]["setting"]["button_submit"])
        Logging("** Input password **")

        PrintYellow("** Check input password success")
        title_page = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "set-my-arch")))
        if title_page.is_displayed():
            Logging("=> Input password success")
        else:
            Logging("=> Input password fail")

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "set-my-arch")))
    except:
        False

    return password_key

def add_folder():
    name_folder = data["title"] + date_time
    ''' Create folder in my archive '''
    try:
        Logging(" ")
        PrintYellow("** Create folder in my archive")
        Commands.Wait20s_ClickElement(data["archive"]["setting"]["folder"])
        Logging("- Select folder")
        
        Commands.Wait20s_ClickElement(data["archive"]["setting"]["add_folder"])
        Logging("- Select my archive")
        
        Commands.Wait20s_ClickElement(data["archive"]["setting"]["enabled"])
        Logging("- Select Permission")
        
        Commands.Wait20s_InputElement(data["archive"]["setting"]["input_folder"], name_folder)
        Logging("- Input name folder")
        
        Commands.Wait20s_ClickElement(data["archive"]["setting"]["save"])
        Logging("=> Create folder in my archive success")
        
    except:
        pass
    
    try:
        Logging(" ")
        PrintYellow("** Check folder have save **")
        check_folder = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[@id='set-my-arch']//span[contains(., '" + name_folder + "')]")))
        if check_folder.is_displayed:
            Logging("=> Create folder success")
            TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder"]["pass"])
        else:
            Logging("=> Create folder fail")
            TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder"]["fail"])
            ValidateFailResultAndSystem("<div>[Archive]Create folder in settings fail </div>")
        
    except:
        Logging("=> Create folder fail")
        TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder"]["fail"])
        pass
    return name_folder

def edit_folder(name_folder):
    name_folder_edit = data["title"] + " " + str(n)
    try:
        Logging(" ")
        PrintYellow("** Edit folder - name - permission")
        Commands.Wait20s_ClickElement("//*[@id='set-my-arch']//span[contains(., '" + name_folder + "')]")
        Commands.Wait20s_ClickElement(data["archive"]["setting"]["modify"])
        Logging("- Select edit folder")
        
        Commands.Wait20s_ClickElement(data["archive"]["setting"]["disabled"])
        Logging("- Edit Permission")
        
        Commands.Wait20s_Clear_InputElement(data["archive"]["setting"]["input_folder"], name_folder_edit)
        Logging("- Input edit name")
        
        Commands.Wait20s_ClickElement(data["archive"]["setting"]["save"])
        Logging("=> Edit folder in my archive")
        TesCase_LogResult(**data["testcase_result"]["archive"]["edit_folder"]["pass"])
        
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["edit_folder"]["fail"])
        pass
    
    return name_folder_edit

def delete_folder(name_folder_edit):
    try:
        Logging(" ")
        PrintYellow("** Delete folder have edit")
        time.sleep(5)
        Commands.Wait20s_ClickElement("//*[@id='set-my-arch']//span[contains(., '" + name_folder_edit + "')]")
        Logging("- Select folder to delete")
        Commands.Wait20s_ClickElement(data["archive"]["setting"]["delete"])
        Logging("- Click button delete")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["archive"]["setting"]["button_ok"])
        Logging("=> Delete folder success")
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_folder"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_folder"]["fail"])
        pass

def admin_execution():
    Logging("- Click Admin archive")

    Logging(" ")
    PrintGreen("============ Test case Admin Archive ============")

    Logging(" ")
    PrintYellow("** Manage Company Archive") 
    Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["Manage_Company"])
    Logging("- Click Manage Company")
    

    ''' Add folder public first 
    -> Create sub folder in folder public
    -> Add authorized for sub folder
    -> Delete sub folder
    -> Delete public folder'''
    try:
        add_folder_test()
    except:
        pass

    try:
        name_folder_admin = folder_public()
    except:
        name_folder_admin = None

    if bool(name_folder_admin) == True:
        try:
            name_subfolder = sub_folder(name_folder_admin)
        except:
            PrintRed(">> Can't countinue execution")
            pass

        try:
            add_authorized_sub_folder(name_subfolder)
        except:
            PrintRed(">> Can't countinue execution")
            pass

        try:
            delete_public_folder(name_folder_admin)
        except:
            PrintRed(">> Can't countinue execution")
            pass

        # try:
        #     delete_public_folder(name_folder_admin)
        # except:
        #     PrintRed(">> Can't countinue execution")
        #     pass
    else:
        PrintRed("=> Create folder public fail")

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
            PrintRed(">> Can't countinue execution")
            pass
    else:
        PrintRed("=> Create folder private fail")

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
        PrintRed("=> Can't add Archive Manager")

    try:
        backup()
    except:
        PrintRed(">> Can't countinue execution")
        pass

    try:
        transfer_company_approval()
    except:
        PrintRed(">> Can't countinue execution")
        pass

    try:
        transfer_company_board()
    except:
        PrintRed(">> Can't countinue execution")
        pass

    try:
        transfer_company_task()
    except:
        PrintRed(">> Can't countinue execution")
        pass

def add_folder_test():
    folder_test = data["archive"]["admin"]["name_folder_test"]

    Logging("- Check domain have folder test")
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='archive-tab-content']//div[contains(@class, 'col-lg-5')]//a[contains(.,'NQuynh_folder')]")))
        Logging("- Domain have folder test")
    except:
        Logging("- Domain don't have folder test")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["parent_folder"])
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["add_folder_admin"])
        Logging("- Create folder test")
        
        Commands.Wait20s_InputElement(data["archive"]["admin"]["ManageCompany"]["input_name_folder"], folder_test)
        Logging("- Input name folder")
        
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["save_folder"])
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["button_close"])
        Logging("- Save folder")

    return folder_test

def folder_public():
    name_folder_admin = data["title"] + date_time + " " + str(n)
    ''' create folder public '''
    try:
        Logging(" ")
        PrintYellow("-> Create folder public")
        Waits.Wait20s_ElementLoaded( data["loading_dialog"])
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["parent_folder"])
        Logging("- Select folder to create")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["add_folder_admin"])
        Logging("- Create folder - Public folder")
        Commands.Wait20s_InputElement(data["archive"]["admin"]["ManageCompany"]["input_name_folder"], name_folder_admin)
        Logging("- Input name folder")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["save_folder"])
        Logging("- Save public folder")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["button_close"])
        Logging("=> Confirm create folder public")
    except:
        pass

    ''' Check folder public have create '''
    try:
        Logging(" ")
        PrintYellow(" ** Check folder public have save ** ")
        folder_public = Waits.Wait20s_ElementLoaded(data["archive"]["admin"]["ManageCompany"]["name_folder"] % str(name_folder_admin))
        if folder_public.is_displayed:
            Logging("=> Create folder public success")
            TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder_public"]["pass"])
        else:
            Logging("=> Create folder public fail")
            TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder_public"]["fail"])
        
    except:
        pass
    
    return name_folder_admin

def folder_private():
    name_folder_private = data["title"] + date_time + str(" private")

    ''' Create folder private -> Permission Disable '''
    try:
        Logging(" ")
        PrintYellow("-> Create folder private")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["parent_folder"])
        Logging("- Select folder to create")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["add_folder_admin"])
        Logging("- Create folder")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["private"])
        Logging("- Private folder")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["disabled"])
        Logging("- Permission: disabled") 
        Commands.Wait20s_InputElement(data["archive"]["admin"]["ManageCompany"]["input_name_folder"], name_folder_private)
        Logging("- Input name folder") 
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["save_folder"])
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["button_close"])
        Logging("- Save private folder") 
    except:
        pass
    
    ''' Check folder private have create '''
    try:
        Logging(" ")
        PrintYellow("**Check folder private have save**")
        time.sleep(5)
        Waits.Wait20s_ElementLoaded(data["archive"]["admin"]["ManageCompany"]["name_folder"] % str(name_folder_private))
        Logging("=> Create folder private success")
        TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder_privte"]["pass"])       
    except:
        Logging("=> Create folder private fail")
        TesCase_LogResult(**data["testcase_result"]["archive"]["add_folder_privte"]["fail"]) 
    
    return name_folder_private
    
def sub_folder(name_folder_admin):
    name_subfolder = data["title"] + str(" sub folder")

    ''' Create sub-folder '''
    try:
        Logging(" ")
        PrintYellow("** Create sub folder in public folder")
        driver.find_element_by_tag_name("body").send_keys(Keys.END)
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["name_folder"] % str(name_folder_admin))
        Logging("- Select puclic to create sub folder")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["add_subfolder"])
        Logging("- Add sub folder")
        Commands.Wait20s_InputElement(data["archive"]["admin"]["ManageCompany"]["input_subfolder"], name_subfolder)
        Logging("- Input name sub folder")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["save_subfolder"])
        Logging("- Save sub folder")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["button_close"])
        Logging("=> Create sub folder admin")
        
    except:
        pass

    ''' Check sub folder have create '''
    try:
        Logging(" ")
        PrintYellow("** Check sub folder have create")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["name_folder"] % str(name_folder_admin))
        PrintYellow("- Select parent folder")
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, data["archive"]["admin"]["ManageCompany"]["name_folder"] % str(name_subfolder))))
        Logging("=> Create subfolder success")
        TesCase_LogResult(**data["testcase_result"]["archive"]["add_subfolder"]["pass"])
    except:
        Logging("=> Create subfolder fail")
        TesCase_LogResult(**data["testcase_result"]["archive"]["add_subfolder"]["fail"])

    return name_subfolder

def add_authorized_sub_folder(name_subfolder):
    ''' Add authorized_Dept for subfoder '''
    try:
        Logging(" ")
        PrintYellow("** Add authorized Dept for subfoder")
        PrintYellow("(Selected Dept. Only + Selected Folders Only)")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["name_folder"] % str(name_subfolder))
        Logging("- Select sub folder")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["Select_authorized_Dept."])
        Logging("- Select authorized Dept")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["Dept."])
        Logging("- Select Dept.")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["add_Dept."])
        Logging("=> Save settings")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["button_close"])
        Logging("=> Confirm close popup")
        TesCase_LogResult(**data["testcase_result"]["archive"]["authorized_Dept"]["pass"])
        
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["authorized_Dept"]["fail"])
        pass

    try:
        Logging(" ")
        PrintYellow("(Include All Sub-Dept.(s) + Include All Sub-Folders)")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["Include_Sub_Dept."])
        Logging("- Select Dept.settings: Include All Sub-Dept.(s)")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["name_folder"] % str(name_subfolder))
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["Select_authorized_Dept."])
        Logging("- Select Authorized Dept.")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["Dept."])
        Logging("- Select Dept.")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["Include_Sub_folder."])
        Logging("- Select Dept. Settings: Include All Sub-Folders")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["add_Dept."])
        Logging("=> Save settings")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["button_close"])
        Logging("=> Confirm close popup")
        TesCase_LogResult(**data["testcase_result"]["archive"]["Dept_settings"]["pass"])
        
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["Dept_settings"]["fail"])
        pass

def delete_public_folder(name_folder_admin):
    ''' Delete public folder '''
    try:
        Logging(" ")
        PrintYellow("** Delete public folder")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["check_all"])
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["del_all"])
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["button_ok"])
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["name_folder"] % str(name_folder_admin))
        Logging("- Select folder")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["del_folder"])
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["button_OK"])
        Logging("=> Delete public folder")
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_subfolder"]["pass"])
        
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_subfolder"]["fail"])
        pass

# def delete_public_folder(name_folder_admin):
#     try:
#         ''' Delete folder '''
#         Logging(" ")
#         PrintYellow("** Delete public folder")
#         Commands.Wait20s_ClickElement("//*[@id='archive-tab-content']//span[contains(., '" + name_folder_admin + "')]")
#                     
#         Logging("- Select parent folder")
#         Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["del_folder"])
#         
#         Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["button_OK"])
#         Logging("=> Delete public folder")
#         TesCase_LogResult(**data["testcase_result"]["archive"]["delete_folder_public"]["pass"])
#         
#     except:
#         TesCase_LogResult(**data["testcase_result"]["archive"]["delete_folder_public"]["fail"])
#         pass

def delete_private_folder(name_folder_private):
    try:
        PrintYellow("** Delete private folder")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["name_folder"] % str(name_folder_private))
        Logging("- Select parent folder")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["del_folder"])
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ManageCompany"]["button_OK"])
        Logging("=> Delete private folder")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["company_archive"])
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_folder_privte"]["pass"])
        
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_folder_privte"]["fail"])

def archive_manager():
    name_manager = data["archive"]["admin"]["ArchiveManager"]["name"]

    ''' Add Archive Manager '''
    Logging(" ")
    PrintYellow("** Add Archive Manager")
    # Commands.Wait20s_ClickElement(data["archive"]["admin"]["ArchiveManager"]["Company_Archive"])
    # 
    Commands.Wait20s_ClickElement(data["archive"]["admin"]["ArchiveManager"]["Archive_Manager"])
    Logging("Access Archive_Manager")
    
    try:
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ArchiveManager"]["button_add"])
        Logging("- Add manager archive")
        Commands.Wait20s_EnterElement(data["archive"]["admin"]["ArchiveManager"]["input_name"], name_manager)
        Logging("- Input name manager archive")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ArchiveManager"]["select_user"])
        Logging("- Select user")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ArchiveManager"]["button_save"])
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ArchiveManager"]["button_OK."])
        Logging("=> Add Archive Manager")
        TesCase_LogResult(**data["testcase_result"]["archive"]["archive_manager"]["pass"])
    except:
        Logging("=> Can't add Archive Manager")
        TesCase_LogResult(**data["testcase_result"]["archive"]["archive_manager"]["fail"])

def delete_manager():
    ''' Delete user manager '''
    try:
        
        PrintYellow("** Delete user manager")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ArchiveManager"]["pick_user"])
        Logging("- Select user to delete")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ArchiveManager"]["button_delete"])
        Logging("- Click button delete")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["ArchiveManager"]["close_popup"])
        Logging("=> Delete Archive Manager")
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_archive_manager"]["pass"])
        
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["delete_archive_manager"]["fail"])
        

def backup():
    name_backup = data["title"] + date_time
    ''' Backup folder '''
    Logging(" ")
    PrintYellow("** Backup folder")
    Commands.Wait20s_ClickElement(data["archive"]["admin"]["backup_folder"]["menu"])
    Logging("- Access Backup folder")
    
    ''' Back up file need edit date backup '''    
    try:    
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["backup_folder"]["button_add"])
        Commands.Wait20s_InputElement(data["archive"]["admin"]["backup_folder"]["input_name"], name_backup)
        Logging("- Input name backup")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["backup_folder"]["selectfolder"])
        
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["backup_folder"]["choosefolder"])
        
        Logging("- Select folder to backup")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["backup_folder"]["click_date_range_from"])
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["backup_folder"]["select_date_range_from"])
        
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["backup_folder"]["click_date_range_to"])
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["backup_folder"]["select_date_range_to"])
        
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["backup_folder"]["save_backup"])
        Logging("- Save backup") 
           
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["backup_folder"]["close_popup"])
        Logging("- Backup folder success")
        TesCase_LogResult(**data["testcase_result"]["archive"]["backup"]["pass"])
        
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["backup"]["fail"])
        

def transfer_company_approval():
    ''' Transfer compnay -> Approval '''
    dept_key = data["archive"]["admin"]["transfer"]["dept"]
    language_key = data["archive"]["admin"]["transfer"]["language"]

    Logging(" ")
    PrintYellow("** Transfer compnay -> Approval")
    Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["transfer_company"])
    Logging("- Transfer company")
    
    try:
        ''' select approval menu '''
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_approval"]["approval"])
        Logging("- Transfer company - Approval")
        Commands.Selectbox_By_Visible_Text(data["archive"]["admin"]["transfer"]["menu_approval"]["folders_option"], dept_key)
        Logging("- Select Dept to transfer")

        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_approval"]["location_option"])
        
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_approval"]["location"])
        Logging("- Select location to transfer")
    
        ''' select language '''       
        Commands.Selectbox_By_Visible_Text(data["archive"]["admin"]["transfer"]["menu_task"]["language_option"], language_key) 
        Logging("- Select language")
    
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_approval"]["button_OK"])
        
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_approval"]["close_popup"])
        Logging("=> Transfer company - Approval")
        
        TesCase_LogResult(**data["testcase_result"]["archive"]["transfer_approval"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["transfer_approval"]["fail"])
        

def transfer_company_board():
    ''' Transfer compnay -> Company Board ''' 
    language_key = data["archive"]["admin"]["transfer"]["language"]
    board_key = data["archive"]["admin"]["transfer"]["board"]

    try:
        Logging(" ")
        PrintYellow("** Transfer compnay -> Company Board")
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["transfer_company"])

        ''' select Company Board menu '''
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_board"]["company_board"])
        Logging("- Transfer company - Company Board")
        
        Commands.Selectbox_By_Visible_Text(data["archive"]["admin"]["transfer"]["menu_board"]["board_option"], board_key)          
        Logging("- Select Company Board transfer")

        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_board"]["location_option"])
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_board"]["location"])
        Logging("- Select location to transfer")
    
        ''' select language '''       
        Commands.Selectbox_By_Visible_Text(data["archive"]["admin"]["transfer"]["menu_task"]["language_option"], language_key) 
        Logging("- Select language")
        
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_board"]["button_OK"])
        
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_board"]["close_popup"])
        Logging("=> Transfer company - Company Board")
        TesCase_LogResult(**data["testcase_result"]["archive"]["transfer_board"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["transfer_board"]["fail"])
        
def transfer_company_task():
    ''' Transfer compnay -> work dairy '''
    dept_key = data["archive"]["admin"]["transfer"]["dept"]
    language_key = data["archive"]["admin"]["transfer"]["language"]

    try:
        Logging(" ")
        PrintYellow("** Transfer compnay -> work dairy")

        Commands.Selectbox_By_Visible_Text(data["archive"]["admin"]["transfer"]["menu_task"]["folders_option"], dept_key)     
        Logging("- Select Dept to transfer")
    
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_task"]["location_option"])
        
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_task"]["location"])
        Logging("- Select location to transfer")
        
        ''' select language '''   
        Commands.Selectbox_By_Visible_Text(data["archive"]["admin"]["transfer"]["menu_task"]["language_option"], language_key) 
        Logging("- Select language")
        
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_task"]["button_OK"])
        
        Commands.Wait20s_ClickElement(data["archive"]["admin"]["transfer"]["menu_task"]["close_popup"])
        Logging("=> Transfer company - work dairy")
        
        TesCase_LogResult(**data["testcase_result"]["archive"]["transfer_work_dairy"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["archive"]["transfer_work_dairy"]["fail"])
        
