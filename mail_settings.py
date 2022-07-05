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
import NQ_login_function

from framework_sample import *
from NQ_login_function import local_path, driver, data, ValidateFailResultAndSystem, Logging, TesCase_LogResult#, #TestlinkResult_Fail, #TestlinkResult_Pass


#chrome_path = os.path.dirname(Path(__file__).absolute())+"\\chromedriver.exe"

n = random.randint(1,3000)
m = random.randint(3000,6000)
date_time = datetime.now().strftime("%Y/%m/%d, %H:%M:%S")

def mail(domain_name):
    driver.get(domain_name + "mail/list/all/")
    Logging('============ Menu Mail ============')
    # ''' Access to menu '''
    # settings()
    # element_admin = Waits.Wait20s_ElementLoaded(data["mail"]["settingsmail"])
    # element_admin.location_once_scrolled_into_view

    # Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["admin"])
    # time.sleep(2)
    # element_admin1 = Waits.Wait20s_ElementLoaded(data["mail"]["settingsmail_1"])
    # element_admin1.location_once_scrolled_into_view
    # time.sleep(2)
    try:
        admin_settings_user = Waits.Wait20s_ElementLoaded(data["mail"]["admin_mail"])
        if admin_settings_user.is_displayed():
            Logging("- Account admin")
            settings()
            Logging(" ")
            Logging("----- Admin mail -----")
            admin_settings()
    except WebDriverException:
        Logging("=> Account user") 
        settings()
  
def settings():
    element = Waits.Wait20s_ElementLoaded(data["mail"]["pull_the_scroll_bar"])
    element.location_once_scrolled_into_view
    time.sleep(1)
    Commands.Wait20s_ClickElement(data["mail"]["settings_mail"])
    time.sleep(1)
    element_1 = Waits.Wait20s_ElementLoaded(data["mail"]["pull_the_scroll_bar_1"])
    element_1.location_once_scrolled_into_view
    Commands.Wait20s_ClickElement(data["mail"]["click_fetching"])
    Logging("- Click setting")
    time.sleep(5)
    Logging(" ")
    Logging("============ Test case settings mail ============")

    try:
        add_signature()
    except:
        Logging(">>>> Cannot continue excution")
        pass

    try:
        delete_signature()
    except:
        Logging(">>>> Cannot continue excution")
        pass
    
    try:
        add_autosort()
    except:
        Logging(">>>> Cannot continue excution")
        pass
    
    try:
        delete_autosort()
    except:
        Logging(">>>> Cannot continue excution")
        pass
    
    try:
        vacation_auto_replies()
    except:
        Logging(">>>> Cannot continue excution")
        pass
    
    '''Add Block addressed setting'''
    try:
        addresses_block = add_block_address()
    except:
        addresses_block = None
    
    list_counter_number = block_total()

    if bool(addresses_block) == True:

        delete_block()

    else:
        if list_counter_number > 0:
            Commands.Wait20s_ClickElement(data["mail"]["settings"]["blockaddress"]["blocked_mail"])
            time.sleep(5)
            Commands.Wait20s_ClickElement(data["mail"]["settings"]["blockaddress"]["del_mail"])
            Logging("=> Delete Blocked Addresses")
            time.sleep(5)
            TesCase_LogResult(**data["testcase_result"]["mail"]["delete_block"]["pass"])
            #TestlinkResult_Pass("WUI-88")
        else:
           Logging(">> Can't continue execution")

    '''Add while list adressed settings'''
    try:
        whilelist = add_while_list()
    except:
        whilelist = None

    list_data = whilelist_total()

    if bool(whilelist) == True:
        
        delete_while_list(whilelist)
    
    else:
        if list_data["list_counter_number_update"] > list_data["list_counter_number"]:
            Commands.Wait20s_ClickElement(data["mail"]["settings"]["whitelist"]["select_list"])
            Logging("- Select list to delete")
            time.sleep(5)
            Commands.Wait20s_ClickElement(data["mail"]["settings"]["whitelist"]["del_list"])
            Logging("- Delete list")
            time.sleep(5)
            TesCase_LogResult(**data["testcase_result"]["mail"]["delete_whilelist"]["pass"])
        else:
            Logging(">> Can't continue execution")
            pass
    # else:
    #     Logging("=> Add while list fail")
    #     TesCase_LogResult(**data["testcase_result"]["mail"]["add_whilelist"]["fail"])

    ''' Add folder settings '''
    try:
        name_folders = add_folder()
    except:
        name_folders = None
    
    if bool(name_folders) == True:
        try:
            share_folder(name_folders)
            delete_folder()
        except:
            Logging(">> Can't continue execution")
            pass
    else:
        Logging("=> Create folders fail")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_folder"]["fail"])

def block_total():
    list_counter = Waits.Wait20s_ElementLoaded(data["mail"]["settings"]["blockaddress"]["total_list"])
    list_counter_number = int(list_counter.text.split(" ")[1])

    return list_counter_number

def whilelist_total():
    list_counter = Waits.Wait20s_ElementLoaded(data["mail"]["settings"]["whitelist"]["total_list"])
    list_counter_number = int(list_counter.text.split(" ")[1])

    list_counter_update = Waits.Wait20s_ElementLoaded(data["mail"]["settings"]["whitelist"]["total_list"])
    list_counter_number_update = int(list_counter_update.text.split(" ")[1])

    list_data = {
        "list_counter_number": list_counter_number,
        "list_counter_number_update": list_counter_number_update
    }
    return list_data

def alias_total():
    list_counter = Waits.Wait20s_ElementLoaded(data["mail"]["settings_admin"]["aliasaccount"]["total_list"])
    Logging("=> Total list number: " + list_counter.text)
    list_counter_number = int(list_counter.text.split(" ")[1])

    return list_counter_number

def add_signature():
    input_text = data["title"] + date_time
    try:
        Logging("** Create signature")
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["signature"]["signature_access"])
        Logging("- Click signature")
        time.sleep(5)

        Logging("- Add signature")
        ''' text '''
        Logging("- Add text signature")
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["signature"]["signature_add"])
        Logging("- Click add signature")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["signature"]["text1"])
        Logging("- Select text to signature")
        time.sleep(5)
        frame_task = driver.find_element_by_class_name("tox-edit-area__iframe")
        driver.switch_to.frame(frame_task)
        Commands.Wait20s_Clear_InputElement(data["mail"]["settings"]["signature"]["text2"], input_text)
        Logging("- Add text")
        time.sleep(5)
        driver.switch_to.default_content()
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["signature"]["text_save"])
        Logging("=> Add signature text success")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_signature"]["pass"])
        #TestlinkResult_Pass("WUI-83")
        time.sleep(10)
    except WebDriverException:
        Logging("- Add signature text fail")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_signature"]["fail"])
        #TestlinkResult_Fail("WUI-83")
        pass

def delete_signature():
    ''' Delete signature '''
    try:
        Logging("** Delete signature") 
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["signature"]["delete_signature1"])
        Logging("- Select signature")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["signature"]["delete_signature2"])
        Logging("=> Delete signature")
        time.sleep(5)
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_signature"]["pass"])
        #TestlinkResult_Pass("WUI-84")
    except WebDriverException:
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_signature"]["fail"])
        Logging("Delete signature fail")
        #TestlinkResult_Fail("WUI-84")

def add_autosort():
    from_key = data["mail"]["settings"]["auto_sort"]["input_from"]
    to_key = data["mail"]["settings"]["auto_sort"]["input_to"]
    subject_key = data["title"] + date_time
    mail_key = data["mail"]["settings"]["auto_sort"]["select_mailbox"]

    ''' Auto sort '''
    Logging(" ")
    Logging("** Auto sort")
    try:
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["auto_sort"]["autosort"])
        Logging("- Access Auto-Sort")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["auto_sort"]["addautosort"])
        Logging("- Click add auto sort")
        time.sleep(5)
        Commands.Wait20s_InputElement(data["mail"]["settings"]["auto_sort"]["input1"], from_key)
        Logging("- Input from")       
        time.sleep(5)
        Commands.Wait20s_InputElement(data["mail"]["settings"]["auto_sort"]["input2"], to_key)
        Logging("- Input to") 
        time.sleep(5)
        Commands.Wait20s_InputElement(data["mail"]["settings"]["auto_sort"]["input3"], subject_key)
        Logging("- Input subject") 
        time.sleep(5)
        Commands.Wait20s_InputElement(data["mail"]["settings"]["auto_sort"]["select1"], mail_key)
        Logging("- Select mail box") 
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["auto_sort"]["include_mail"])
        Logging("- Select Include existing mail") 
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["auto_sort"]["save_button"])
        Logging("=> Save auto sort") 
        #TestlinkResult_Pass("WUI-85")
        time.sleep(5)
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_auto_sort"]["pass"])
    except WebDriverException:
        Logging("Auto-Sort fail")  
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_auto_sort"]["fail"]) 
        #TestlinkResult_Fail("WUI-85")
    
    '''Logging("** Check auto sort create success")
    autosort = Commands.Wait20s_ClickElement("//*[@id='ngw.mail.autosort']//table/tbody/tr/td")
    if autosort.text == 'FROM : quynh2, TO : quynh1, SUBJECT : test':
        Logging("=> Auto-sort create success")
    else:
        Logging("=> Auto-sort create fail")
        ValidateFailResultAndSystem("<div>[Mail]Auto-sort create fail </div>")'''

def delete_autosort():
    ''' Delete auto sort '''
    Logging(" ")
    try:
        Logging("** Delete auto sort")
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["auto_sort"]["autosort_delete"])
        Logging("=> Delete Auto-Sort")
        time.sleep(5)
        #TestlinkResult_Pass("WUI-86")
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_auto_sort"]["pass"])
        time.sleep(5)
    except WebDriverException:
        pass
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_auto_sort"]["fail"])
        #TestlinkResult_Fail("WUI-86")

def vacation_auto_replies():
    text_key = data["title"] + date_time
    ''' Vacation auto replies '''
    Logging(" ")
    Logging("** Vacation auto replies")
    Commands.Wait20s_ClickElement(data["mail"]["settings"]["auto_replies"]["autoreplies"])
    Logging("- Access vacation auto replies")
    time.sleep(5)
    Commands.Wait20s_ClickElement(data["mail"]["settings"]["auto_replies"]["on/off_autoreplies"])
    Logging("- Click turn on vacation auto replies")
    time.sleep(5)

    Commands.Wait20s_ClickElement(data["mail"]["settings"]["auto_replies"]["date_end"])
    time.sleep(2)
    driver.find_element_by_css_selector(data["mail"]["settings"]["auto_replies"]["select_enddate"]) 
    Logging("- Set date end of auto reply")
    time.sleep(5)
    Commands.Wait20s_InputElement(data["mail"]["settings"]["auto_replies"]["input_text"], text_key)
    Logging("- Input message")
    time.sleep(5)
    Commands.Wait20s_ClickElement(data["mail"]["settings"]["auto_replies"]["save_button"])   
    Logging("=> Save turn on vacation auto replies")     
    time.sleep(5)

    Logging("** Check auto reply have turn on")
    button = Waits.Wait20s_ElementLoaded(data["mail"]["settings"]["auto_replies"]["on/off_autoreplies"])
    if button.is_enabled():
        Logging("=> Create vacation auto replies success") 
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_auto_replies"]["pass"])
    else:
        Logging("=> Create vacation auto replies fail")
        ValidateFailResultAndSystem("<div>[Mail]Create vacation auto replies fail </div>")
    time.sleep(5)

    Commands.Wait20s_ClickElement(data["mail"]["settings"]["auto_replies"]["on/off_autoreplies"])
    Logging("=> Turn off auto reply")
    time.sleep(5)
    Commands.Wait20s_ClickElement(data["mail"]["settings"]["auto_replies"]["save_button"])
    Logging("=> Save turn off vacation auto replies") 

    time.sleep(5)

def add_block_address():
    addresses_block = data["mail"]["settings"]["blockaddress"]["input_address"]
    driver.find_element_by_tag_name("body").send_keys(Keys.END)
    time.sleep(5)
    Logging("** Add block")
    try:
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["blockaddress"]["block_address"])
        Logging("- Access Blocked Addresses")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["blockaddress"]["text_address"]).send_keys(addresses_block)
        Logging("- Input block addresses")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["blockaddress"]["add_mail"])
        Logging("=> Add Blocked Addresses")
        time.sleep(5)

        Logging("** Check add block addresses success")
        Waits.Wait20s_ElementLoaded("//*[@id='ngw.mail.blockaddress']//td[contains(., '" + addresses_block + "')]")
        Logging("=> Add Block success")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_block"]["pass"])
        time.sleep(5)
    except WebDriverException:
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_block"]["fail"])
        Logging("=> Block addresses fail")
        #TestlinkResult_Fail("WUI-87")
    return addresses_block
    
def delete_block():
    ''' Del block '''
    Logging("")
    try:
        Logging("** Del block")
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["blockaddress"]["blocked_mail"])
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["blockaddress"]["del_mail"])
        Logging("=> Delete Blocked Addresses")
        time.sleep(5)
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_block"]["pass"])
        #TestlinkResult_Pass("WUI-88")
    except WebDriverException:
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_block"]["fail"])
        #TestlinkResult_Fail("WUI-88")

def add_while_list():
    whilelist = data["mail"]["settings"]["whitelist"]["input_list_1"]
    ''' Add while List '''
    Commands.ScrollDown()
    try:
        Logging(" ")
        Logging("** Add while List")
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["whitelist"]["white_list"])
        Logging("- Access to white list")
        time.sleep(5)

        Commands.Wait20s_InputElement(data["mail"]["settings"]["whitelist"]["addlisst"], whilelist)
        Logging("- Add white list")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["whitelist"]["add_button"])
        Logging("=> Save white list")
        time.sleep(5)

        Logging("** Search while List")
        Commands.Wait20s_Clear_InputElement(data["mail"]["settings"]["whitelist"]["searchwhitelist"], whilelist)
        Logging("Input white list")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["whitelist"]["search"])
        Logging("Search white list")
        time.sleep(5)
    except WebDriverException:
        pass
    
    try:
        Logging("** Check add while list success")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.mail.whitelist']//td[contains(., '" + whilelist + "')]")))
        Logging("=> Add while list success")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_whilelist"]["pass"])
    except WebDriverException:
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_whilelist"]["fail"])
    return whilelist

def delete_while_list(whilelist):
    try:
        Logging("** Del while List")
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["whitelist"]["refesh"])
        Logging("Refresh white list")
        Commands.Wait20s_Clear_InputElement(data["mail"]["settings"]["whitelist"]["searchwhitelist"], whilelist)
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["whitelist"]["search"])
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["whitelist"]["select_list"])
        Logging("Select list to delete")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["whitelist"]["del_list"])
        Logging("Delete list")
        time.sleep(5)
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_whilelist"]["pass"])
    except WebDriverException:
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_whilelist"]["fail"])
        pass

def add_folder():
    name_folders = data["title"] + str(n)
    try:
        Logging(" ")
        Logging("** Folders")

        Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["st_folders"])
        Logging("- Access to Folders")
        time.sleep(2)
        Logging("- Create folder")
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["add_folder"])
        time.sleep(2)
        Commands.Wait20s_InputElement(data["mail"]["settings"]["folders"]["input_name"], name_folders)
        Logging("- Input name folder")
        time.sleep(2)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["button_save"][0])
        Logging("=> Create folders success")
        time.sleep(5)

        try:
            popup = Waits.Wait20s_ElementLoaded(data["mail"]["settings"]["folders"]["pop_up"])
            if popup.is_displayed():
                Logging("Folder have exits - Create new folder")
                time.sleep(3)
                Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["popup_exits"])
                time.sleep(2)
                Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["add_folder"])
                time.sleep(2)
                Commands.Wait20s_InputElement(data["mail"]["settings"]["folders"]["input_name"], name_folders)
                Logging("- Input name folder")
                time.sleep(2)
                Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["button_save"][0])
                Logging("=> Create new folders success")
            else:
                Logging("Pop up error duplicate not show")
        except WebDriverException:
            Logging("Pop up error duplicate not show")
        time.sleep(5)
    except WebDriverException:
        pass
    
    try:
        Logging("** Check folder have create")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='mail_setting_form']//li//a[contains(., '" + str(name_folders) + "')]")))
        Logging("=> Create folders success")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_folder"]["pass"])
    except:
        Logging("=> Create folders fail")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_folder"]["fail"])

    return name_folders

def share_folder(name_folders):
    org_key = data["mail"]["settings"]["folders"]["org_text"]
    try:
        Commands.Wait20s_ClickElement("//*[@id='mail_setting_form']//li//a[contains(., '" + name_folders + "')]")
        time.sleep(3)
        Logging("- Select share permission")
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["share"])
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["usingshare"])
        Logging("- Check using share")
        time.sleep(3)
        Commands.Wait20s_EnterElement(data["mail"]["settings"]["folders"]["org_input"], org_key)
        Logging("- Input name user")
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["org_select"])
        Logging("- Select user")
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["org_add"])
        Logging("- Add user success")
        time.sleep(3)

        options_list = ["Read/Share/Reply/Forward", "Read Mail", "Shared Mail", "Reply/Forward", "Read/Share", "Read/Reply/Forward", "Share/Reply/Forward"]

        sel = Select(Waits.Wait20s_ElementLoaded(data["mail"]["settings"]["folders"]["dropdown"]))
        time.sleep(2)
        sel.select_by_visible_text(random.choice(options_list))
        Logging("- Select permission for user")
        
        time.sleep(2)
        
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["button_save"][1])
        Logging("=> Share folder success")
        time.sleep(5)

        Logging("** Check Share folder")
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["share"])
        time.sleep(3)
        button_share = Waits.Wait20s_ElementLoaded(data["mail"]["settings"]["folders"]["usingshare"])
        if button_share.is_enabled():
            Logging("=> Share folder success")
        else:
            Logging("=> Share folder fail")
            ValidateFailResultAndSystem("<div>[Mail]Share folder fail </div>")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["button_close"])
        time.sleep(5)
    except:
        pass

def upload_eml(name_folders):
    driver.find_element_by_tag_name("body").send_keys(Keys.END)
    time.sleep(5)

    Commands.Wait20s_ClickElement("//*[@id='mail_setting_form']//li//a[contains(., '" + name_folders + "')]")

    Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["Eml"])
    time.sleep(2)
    Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["addfile"])
    time.sleep(2)
    Commands.Wait20s_InputElement(data["mail"]["settings"]["folders"]["getfile"], NQ_login_function.file_upload)
    Logging("- Select file to upload")
    time.sleep(2)
    Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["button_save"][2])

    time.sleep(15)
    Logging("=> Upload EML success")

def backup_mailbox():
    Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["backup"])
    time.sleep(2)
    Logging("=> Backup success")

def empty_mailbox():
    Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["empty"])
    time.sleep(2)
    Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["close_popup"])
    Logging("=> Empty folder success")
    time.sleep(5)

def delete_folder():
    ''' Delete folders '''
    try:
        Logging("** Delete folder")
        Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["delete"])
        time.sleep(2)
        Logging("=> Delete success")
        time.sleep(5)
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_folder"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_folder"]["fail"])
        pass

    driver.find_element_by_tag_name("body").send_keys(Keys.HOME)
    time.sleep(5)

    '''Logging("** Download folder")
    Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["share_mailbox"])
    Commands.Wait20s_ClickElement(data["mail"]["settings"]["folders"]["downfile"])
    time.sleep(2)
    Logging("=> Download share mail box success")
    time.sleep(5)   '''

def forwarding():
    email = data["mail"]["settings"]["forwarding"]["text_inpiut"]
    Logging(" ")
   
    Commands.Wait20s_ClickElement("//a[contains(@data-defaulthref,'#/mail/forwarding/setting') and contains(.,' Forwarding')]")
    Logging("** Forwarding")
    Commands.Wait20s_ClickElement(data["mail"]["settings"]["forwarding"]["add_forwarding"])
    Commands.Wait20s_InputElement(data["mail"]["settings"]["forwarding"]["input_email"], email)
    Logging("- Input email")

    option_list = int(len(driver.find_elements_by_xpath(data["mail"]["settings_admin"]["forwarding"]["option_list"])))
    forwarding_list = []
    i = 0
    for i in range(option_list):
        i += 1
        mode_list = Waits.Wait20s_ElementLoaded(data["mail"]["settings_admin"]["forwarding"]["select_option"] + "[" + str(i) + "]")
        forwarding_list.append(mode_list.text)
    
    Logging("- Total of view mode list: " + str(len(forwarding_list)))
    x = random.choice(forwarding_list)
    time.sleep(2)
    select_mode_view = Waits.Wait20s_ElementLoaded(data["mail"]["settings_admin"]["forwarding"]["option_list"] + "[contains(.,'" + str(x) + "')]")
    select_mode_view
    Logging("- Select Forwarding option")

    Commands.Wait20s_ClickElement(data["mail"]["settings"]["forwarding"]["save_button"])
    Logging("- Save forwarding")

    Commands.Wait20s_ClickElement("//*[@id='ngw.mail.forwarding']//table//td[contains(., '" + email + "')]")
    Logging("- Select email")
    Commands.Wait20s_ClickElement(data["mail"]["settings"]["forwarding"]["delete_button"])
    Logging("- Delete email have set forwarding")

def approval_total():
    list_counter = Waits.Wait20s_ElementLoaded(data["mail"]["settings_admin"]["approval_mailbox"]["total_list"])
    list_counter_number = int(list_counter.text.split(" ")[1])

    return list_counter_number

def admin_settings():
    Logging("** Open settings admin mail")
    Logging('============ Test case settings admin mail ============')
    Commands.Wait20s_ClickElement(data["mail"]["settings_mail"])
    element_admin = Waits.Wait20s_ElementLoaded(data["mail"]["settingsmail"])
    element_admin.location_once_scrolled_into_view

    Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["admin"])
    time.sleep(2)
    element_admin1 = Waits.Wait20s_ElementLoaded(data["mail"]["settingsmail_1"])
    element_admin1.location_once_scrolled_into_view
    time.sleep(2)
    Logging("=> Click settings admin")
    time.sleep(5)

    # try:
    #     aliasdomain_name = add_alias_domain()
    # except:
    #     aliasdomain_name = None

    # if bool(aliasdomain_name) == True:
    #     try:
    #         delete_alias_domain(aliasdomain_name)
    #     except:
    #         Logging(">> Can't continue execution")
    #         pass
    # else:
    #     Logging("=> Add alias domain fail")
    #     TesCase_LogResult(**data["testcase_result"]["mail"]["alias_domain"]["fail"])

    try:
        approval_mailbox()
    except:
        Logging(">> Can't continue execution")
        pass
    
    try:
        company_signature()
    except:
        Logging(">> Can't continue execution")
        pass

    try:
        send_limit()
    except:
        Logging(">> Can't continue execution")
        pass
    try:
        alias_account()
    except:
        Logging("=> Can't create alias account")
        TesCase_LogResult(**data["testcase_result"]["mail"]["alias_account"]["fail"])
        pass

def add_alias_domain():
    aliasdomain_name = data["mail"]["settings_admin"]["alias_domian"]["aliasdomain_input"]
    try:
        Logging("** Add alias domain")

        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["alias_domian"]["aliasdomain"])
        Logging("- Access alias domain")
        time.sleep(5)

        Commands.Wait20s_InputElement(data["mail"]["settings_admin"]["alias_domian"]["input4"], aliasdomain_name)
        Logging("- Input domain")
        time.sleep(5)

        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["alias_domian"]["add_aliasdomain"])
        Logging("- Add alias domain")
        time.sleep(5)
    except:
        pass
    
    try:
        Logging("** Check alias domain have add")
        Waits.Wait20s_ElementLoaded("//*[@id='ngw.mail.adminalias_domain']//table[contains(., '" + aliasdomain_name + "')]")
        Logging("=> Add alias domain success")
        TesCase_LogResult(**data["testcase_result"]["mail"]["alias_domain"]["pass"])
        time.sleep(5) 
    except:
        Logging("=> Add alias domain fail")
        TesCase_LogResult(**data["testcase_result"]["mail"]["alias_domain"]["fail"])

    return aliasdomain_name

def delete_alias_domain(aliasdomain_name):
    Logging("** Delete alias domain")
    try:
        Commands.Wait20s_ClickElement("//*[@id='ngw.mail.adminalias_domain']//table[contains(., '" + aliasdomain_name + "')]//span")
        Logging("- Select alias domain")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["alias_domian"]["del_aliasdomain"])
        Logging("=> Delete alias domain")
        time.sleep(5)
        #TestlinkResult_Pass("WUI-131")
        TesCase_LogResult(**data["testcase_result"]["mail"]["del_alias_domain"]["pass"])
    except:
        TesCase_LogResult(**data["testcase_result"]["mail"]["del_alias_domain"]["fail"])
        #TestlinkResult_Fail("WUI-131")
        pass

def add_domain_user():
    text = data["mail"]["settings_admin"]["alias_domian"]["add_domain"]
    try:
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["alias_domian"]["add_domain"])
        Logging("- Add user (group) for domain")
        Commands.Wait20s_EnterElement("//*[@id='aliasDomainOrg']//input[contains(@type, 'text')]", text)
        Logging("- Input key user")
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["alias_domian"]["user_1"])
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["alias_domian"]["user_2"])
        Logging("- Select user")
        time.sleep(2)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["alias_domian"]["add_button"])
        Logging("- Add user")
        time.sleep(2)
        Commands.Wait20s_ClickElement("//*[@id='directive-domains']//label")
        Logging("- Select domain")
        time.sleep(2)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["alias_domian"]["save_button"])
        Logging("- Save")
    except:
        pass

def approval_mailbox():
    time.sleep(5)
    element_admin = Waits.Wait20s_ElementLoaded(data["mail"]["settingsmail"])
    element_admin.location_once_scrolled_into_view

    
    element_admin1 = Waits.Wait20s_ElementLoaded(data["mail"]["settingsmail_1"])
    element_admin1.location_once_scrolled_into_view

    Logging(" ")
    try:
        approval = Waits.Wait20s_ElementLoaded(data["mail"]["settings_admin"]["approval_mailbox"]["mailbox"])
        if approval.is_displayed():
            Logging("Add approval mailbox")
            Forced_approval()
            Selective_approval()
            delete_approval_mail_box()
    except WebDriverException:
        Logging("=> Domain don't have menu approval mail box")
        pass
    time.sleep(5)

def Forced_approval():
    Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["mailbox"])
    Logging("- Select approval mailbox")
    f = Waits.Wait20s_ElementLoaded(data["mail"]["settings_admin"]["approval_mailbox"]["policy"])
    driver.execute_script("arguments[0].scrollIntoView();",f)
    try:
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["policy"])
        Logging("- Select approval policy")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["add_policy"])
        Logging("- Add approval policy")
        time.sleep(5)
    #-----------------------------------------Forced approval----------------------------------#
        user_key = data["mail"]["settings_admin"]["approval_mailbox"]["approver_input_1"]
        forced_approval = data["mail"]["settings_admin"]["approval_mailbox"]["input_name_1"] + str(n)

        Commands.Wait20s_InputElement(data["mail"]["settings_admin"]["approval_mailbox"]["input5"], forced_approval)

        Logging("** Select type policy Forced approval")
        time.sleep(5)
        Logging("- Select user final approval")
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["select_approver"])
        Logging("- Select organization")
        time.sleep(5)
        Commands.Wait20s_EnterElement(data["mail"]["settings_admin"]["approval_mailbox"]["approver_input"], user_key)
        Logging("- Input user")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["approver_final_1"])
        Logging("=> Add approval")
        time.sleep(5)

        #-------Permission Recipient----------#
        org_key = data["mail"]["settings_admin"]["approval_mailbox"]["organization_input_1"]

        Logging("** Selective approval")
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["select_organization_1"])
        Logging("- Select organization of Permission Recipient**")
        time.sleep(5)
        Commands.Wait20s_EnterElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_input"], org_key)
        Logging("- Input user")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_select_1"])
        Logging("- Select user")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_add_1"])
        Logging("- Add user")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_save_1"])
        Logging("=> Save Permission Recipient approver")
        time.sleep(5)
        
        #-------------Mid-Approver----------#
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["select_organization_2"])
        time.sleep(5)
        Logging("** Select organization of mid approver**")
        
        Commands.Wait20s_EnterElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_input_2"], org_key)
        Logging("- Mid-Approver")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_select_2"])
        Logging("- Select user")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_add_2"])
        Logging("- Add user")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_save_2"])
        Logging("- Save user")
        time.sleep(5)

        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["save"])
        Logging("=> Save Approval Policy - Forced approval")
        TesCase_LogResult(**data["testcase_result"]["mail"]["approval_mail_box_forced_approval"]["pass"])
        time.sleep(5)
    except:
        Logging("Can't create forced approval")
        TesCase_LogResult(**data["testcase_result"]["mail"]["approval_mail_box_forced_approval"]["fail"])

def Selective_approval():
    try:
        selective_approval = data["mail"]["settings_admin"]["approval_mailbox"]["input_name_2"] + str(n)
        Logging("** Selective approval")
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["add_policy"])
        Logging("- Add approval policy")
        time.sleep(5)

        Commands.Wait20s_InputElement(data["mail"]["settings_admin"]["approval_mailbox"]["input5"], selective_approval)
        Logging("- Select basic policy")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["select_type"])
        Logging("=> Selective approval")
        time.sleep(5)

        #-------Final approval----------#
        final_user = data["mail"]["settings_admin"]["approval_mailbox"]["approver_input_1"]

        Logging("** Select final approval")
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["select_approver"])
        Logging("- Select organization")
        time.sleep(5)
        Commands.Wait20s_EnterElement(data["mail"]["settings_admin"]["approval_mailbox"]["approver_input"], final_user)
        Logging("- Input final approval")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["approver_final_2"])
        Logging("=> Select final approval")
        time.sleep(5)

        #-------Permission Recipient----------#
        org_key_1 = data["mail"]["settings_admin"]["approval_mailbox"]["organization_input_1"]
        Logging("** Select Permission Recipient")
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["select_organization_1"])
        Logging("- Select organization of Permission Recipient")
        time.sleep(5)
        Commands.Wait20s_EnterElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_input"], org_key_1)
        time.sleep(5)
        Logging("- Input user")
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_select_3"])
        Logging("- Select user")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_add_3"])
        Logging("- Add user")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_save_1"])
        Logging("=> Save Permission Recipient approver")
        time.sleep(5)

        #-------Mid-Approver----------#
        Logging("** Select Mid-Approver")
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["select_organization_2"])
        Logging("** Select organization of mid approver**")
        time.sleep(5)
        
        Commands.Wait20s_EnterElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_input_4"], org_key_1)
        Logging("- Input user")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_select_4"])
        Logging("- Select user")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_add_4"])
        Logging("- Add user")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["organization_save_2"])
        Logging("- Save Mid-Approver approver")
        time.sleep(5)

        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["save"])
        Logging("=> Save Approval Policy Selective approval")
        TesCase_LogResult(**data["testcase_result"]["mail"]["approval_mail_box_selective_approval"]["pass"])
        #TestlinkResult_Pass("WUI-128")
        time.sleep(5)
    except:
        Logging("Can't create selective approval")
        TesCase_LogResult(**data["testcase_result"]["mail"]["approval_mail_box_selective_approval"]["fail"])
        #TestlinkResult_Fail("WUI-128")
        pass

def delete_approval_mail_box():
    try:
        Logging("** Del approval mailbox")
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["select_policy_1"])
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["select_policy_2"])
        Logging("-  Select policy to delete")
        time.sleep(5)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["remove_1"])
        time.sleep(2)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["approval_mailbox"]["remove_2"])
        Logging("=> Delete policy")
        #TestlinkResult_Pass("WUI-129")
        TesCase_LogResult(**data["testcase_result"]["mail"]["del_approval_mail_box"]["pass"])
        time.sleep(5)
    except:
        TesCase_LogResult(**data["testcase_result"]["mail"]["del_approval_mail_box"]["fail"])
        #TestlinkResult_Fail("WUI-129")
        pass

def send_limit(): 
    Logging(" ")
    Logging("** Sent limit")
    text_domain = data["mail"]["settings_admin"]["sentlimit"]["input_text"][0]
    text_file = data["mail"]["settings_admin"]["sentlimit"]["input_text"][1]

    try:
        send_limit = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["sentlimit"]["send"])))
        if send_limit.is_displayed():
            send_limit
            time.sleep(2)
            Logging("-> Asset send limit")
            time.sleep(5)
            Commands.Wait20s_InputElement(data["mail"]["settings_admin"]["sentlimit"]["input"][0], text_domain)
            time.sleep(5)
            Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["sentlimit"]["button_add"][0])
            Logging("** Create domain")
            time.sleep(5)

            Logging("** Check domain limt create success")
            domain = Waits.Wait20s_ElementLoaded("//*[@id='domain-limit']//td[contains(., '" + text_domain + "')]")
            if domain.is_displayed():
                Logging("=> Create domain limit success")
                TesCase_LogResult(**data["testcase_result"]["mail"]["add_domain_limt"]["pass"])
            else:
                Logging("=> Create domain limit fail")
                TesCase_LogResult(**data["testcase_result"]["mail"]["add_domain_limt"]["fail"])
                ValidateFailResultAndSystem("<div>[Mail]Create domain limit fail </div>")
            time.sleep(5)

            try:
                Logging("** Delete domain")
                Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["sentlimit"]["select"][0])
                time.sleep(2)
                Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["sentlimit"]["button_delete"])
                time.sleep(2)
                Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["sentlimit"]["close_popup"])
                Logging("=> Delete domain")
                time.sleep(2)
                TesCase_LogResult(**data["testcase_result"]["mail"]["del_domain_limt"]["pass"])
            except:
                TesCase_LogResult(**data["testcase_result"]["mail"]["del_domain_limt"]["fail"])
                
            
            Logging("** Create file limit")
            Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["sentlimit"]["file"])
            time.sleep(2)
            Commands.Wait20s_InputElement(data["mail"]["settings_admin"]["sentlimit"]["input"][1], text_file)
            time.sleep(2)
            Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["sentlimit"]["button_add"][1])
            Logging("=> File limit have create")
            time.sleep(5)

            try:
                Logging("** Check file limit create success")
                Waits.Wait20s_ElementLoaded("//*[@id='file-limit']//td[contains(., '" + text_file + "')]")
                Logging("=> Create file limit success")
                TesCase_LogResult(**data["testcase_result"]["mail"]["add_file_limt"]["pass"])
            except:
                Logging("=> Create file limit fail")
                TesCase_LogResult(**data["testcase_result"]["mail"]["add_file_limt"]["fail"])
            time.sleep(5)

            try:
                Logging("** Delete file limit")
                Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["sentlimit"]["select"][1])
                time.sleep(2)
                Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["sentlimit"]["button_delete"])
                time.sleep(2)
                Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["sentlimit"]["close_popup"])
                Logging("=> Delete file limit")
                TesCase_LogResult(**data["testcase_result"]["mail"]["del_file_limit"]["pass"])
            except:
                TesCase_LogResult(**data["testcase_result"]["mail"]["del_file_limit"]["fail"])
               
        else:
            print("=> Domain don't have send limit")
    except:
        Logging("=> Domain don't have send limit")
        pass
    time.sleep(5)

def company_signature():
    ''' Company signature '''
    Logging(" ")
    Logging("** Company signature **")
   
    Commands.Wait20s_ClickElement("//a[contains(@data-defaulthref, '#/mail/admin/signature')]")
    
    time.sleep(5)

    force_list = int(len(driver.find_elements_by_xpath(data["mail"]["settings_admin"]["signature_company"]["force_list"])))
    force_apply_list = []
    i = 0
    for i in range(force_list):
        i += 1
        mode_list = Waits.Wait20s_ElementLoaded(data["mail"]["settings_admin"]["signature_company"]["force_select"] + "[" + str(i) + "]")
        force_apply_list.append(mode_list.text)
    
    Logging("- Total of view mode list: " + str(len(force_apply_list)))
    x = random.choice(force_apply_list)
    Logging(x)
    time.sleep(2)
    Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["signature_company"]["force_list"] + "[contains(.,'" + str(x) + "')]")
    Logging("- Select Force apply")
    
    if str(x) == "Forced-Appending":
        Logging("- Select mode view: Forced-Appending")
        time.sleep(2)

        position_list = int(len(driver.find_elements_by_xpath(data["mail"]["settings_admin"]["signature_company"]["signature_position"])))
        signature_position_list = []
        i = 0
        for i in range(position_list):
            i += 1
            signature_list = Waits.Wait20s_ElementLoaded(data["mail"]["settings_admin"]["signature_company"]["position"] + "[" + str(i) + "]")
            signature_position_list.append(signature_list.text)
        
        Logging("- Total of signature position: " + str(len(signature_position_list)))
        x = random.choice(signature_position_list)
        time.sleep(2)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["signature_company"]["signature_position"] + "[contains(.,'" + str(x) + "')]")
        Logging("- Select Signature position")

        input_text = data["mail"]["settings_admin"]["signature_company"]["input_signature"]
        frame_task = driver.find_element_by_class_name("tox-edit-area__iframe")
        driver.switch_to.frame(frame_task)
        text = driver.find_element_by_id("tinymce")
        text.clear()
        text.send_keys(input_text)
        Commands.SwitchToDefaultContent()
        time.sleep(2)
        driver.find_element_by_tag_name("body").send_keys(Keys.END)

        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["signature_company"]["button_save"])
        Logging("- Add company signature - Forced-Appending")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_signature_company"]["pass"])
    elif str(x) == "Force replace":
        input_text = data["mail"]["settings_admin"]["signature_company"]["input_signature"]
        frame_task = driver.find_element_by_class_name("tox-edit-area__iframe")
        driver.switch_to.frame(frame_task)
        text = driver.find_element_by_id("tinymce")
        text.clear()
        text.send_keys(input_text)
        Commands.SwitchToDefaultContent()
        time.sleep(2)
        driver.find_element_by_tag_name("body").send_keys(Keys.END)

        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["signature_company"]["button_save"])
        Logging("- Add company signature - Force replace")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_signature_company"]["pass"])
    elif str(x) == "None":
        input_text = data["mail"]["settings_admin"]["signature_company"]["input_signature"]
        frame_task = driver.find_element_by_class_name("tox-edit-area__iframe")
        driver.switch_to.frame(frame_task)
        text = driver.find_element_by_id("tinymce")
        text.clear()
        text.send_keys(input_text)
        Commands.SwitchToDefaultContent()
        time.sleep(2)
        driver.find_element_by_tag_name("body").send_keys(Keys.END)

        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["signature_company"]["button_save"])
        Logging("- Add company signature - None")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_signature_company"]["pass"])
    else:
        Logging("- Can't add signature company")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_signature_company"]["fail"])

def alias_account():
    name_account = data["title"] + date_time
    name_alias = data["mail"]["settings_admin"]["aliasaccount"]["input_nameallias"] + str(n)
    name_org = data["mail"]["settings_admin"]["aliasaccount"]["input_nameorg"]
    ''' Alias account '''
    try:
        Logging(" ")
        Logging("** Alias account")
        driver.find_element_by_tag_name("body").send_keys(Keys.END)
        time.sleep(3)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["aliasaccount"]["allias"])
        Logging("- Access alias account")
        time.sleep(5)
        Commands.Wait20s_InputElement(data["mail"]["settings_admin"]["aliasaccount"]["input_name"], name_account)
        Logging("- Input name account")
        time.sleep(2)
        Commands.Wait20s_InputElement(data["mail"]["settings_admin"]["aliasaccount"]["input_allias"], name_alias)
        Logging("- Input alias account")
        time.sleep(5)
        
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["aliasaccount"]["choose_org"])
        time.sleep(2)
        Commands.Wait20s_EnterElement(data["mail"]["settings_admin"]["aliasaccount"]["input_org"], name_org)
        Logging("- Input user")
        time.sleep(5)

        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["aliasaccount"]["select_org"])
        Logging("- Select user")
        time.sleep(2)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["aliasaccount"]["button_add_org"])
        Logging("- Add user")
        time.sleep(2)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["aliasaccount"]["button_save_org"])
        Logging("=> Save Organization")
        time.sleep(5)

        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["aliasaccount"]["button_save"])
        Logging("=> Save alias account")
        time.sleep(2)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["aliasaccount"]["close_popup"])
        time.sleep(5)
        TesCase_LogResult(**data["testcase_result"]["mail"]["alias_account"]["pass"])
    except:
        Logging("=> Can't create alias account")
        TesCase_LogResult(**data["testcase_result"]["mail"]["alias_account"]["fail"])
        pass

def log_analysis():
    try:
        ''' Log Analysis '''
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["log_analysis"]["loganalysis"])
        Logging("** Log Analysis **")
        time.sleep(2)
        #Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["log_analysis"]["filter_type"])
        Logging("- Select filter type")
        time.sleep(2)
        options_list = ["Received", "Send", "Spam", "Block", "Etc."]

        sel = Select(Waits.Wait20s_ElementLoaded(data["mail"]["settings_admin"]["log_analysis"]["filter_type"]))
        sel.select_by_visible_text(random.choice(options_list))
        Logging("=> Select type ")
        time.sleep(2)
        Commands.Wait20s_ClickElement(data["mail"]["settings_admin"]["log_analysis"]["search"])
        Logging("=> Search")
        time.sleep(2)

        type_select = Waits.Wait20s_ElementLoaded("//*[@id='mail_admin_log_analysis']//div[contains(@data-ng-show, 'isListView()')]")
        type_select_text = type_select.text
        text = type_select_text.split(" ")[1]
        if text == '0':
            print("=> No data")
        else:
            print("=> Search suceess")
        time.sleep(5)
    except:
        pass