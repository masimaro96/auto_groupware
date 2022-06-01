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

def mail(domain_name):
    driver.get(domain_name + "mail/list/all/")
    Logging('============ Menu Mail ============')
    # ''' Access to menu '''
    settings()
    element_admin = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settingsmail"])))
    element_admin.location_once_scrolled_into_view

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["admin"]))).click()
    time.sleep(2)
    element_admin1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settingsmail_1"])))
    element_admin1.location_once_scrolled_into_view
    time.sleep(2)
    try:
        admin_settings_user = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["click_menu"])))
        if admin_settings_user.is_displayed():
            Logging("- Account admin")
            admin_settings_user.click()
            Logging("- Admin mail")
            admin_settings()
    except WebDriverException:
        Logging("=> Account user") 
  
def settings():
    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["mail"]["pull_the_scroll_bar"])))
    element.location_once_scrolled_into_view
    time.sleep(1)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_mail"]))).click()
    time.sleep(1)
    element_1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["mail"]["pull_the_scroll_bar_1"])))
    element_1.location_once_scrolled_into_view
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["click_fetching"]))).click()
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
            driver.find_element_by_xpath(data["mail"]["settings"]["blockaddress"]["blocked_mail"]).click()
            time.sleep(5)
            driver.find_element_by_xpath(data["mail"]["settings"]["blockaddress"]["del_mail"]).click()
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
        
        delete_while_list()
    
    else:
        if list_data["list_counter_number_update"] > list_data["list_counter_number"]:
            driver.find_element_by_xpath(data["mail"]["settings"]["whitelist"]["select_list"]).click()
            Logging("- Select list to delete")
            time.sleep(5)
            driver.find_element_by_xpath(data["mail"]["settings"]["whitelist"]["del_list"]).click()
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
    list_counter = driver.find_element_by_xpath(data["mail"]["settings"]["blockaddress"]["total_list"])
    list_counter_number = int(list_counter.text.split(" ")[1])

    return list_counter_number

def whilelist_total():
    list_counter = driver.find_element_by_xpath(data["mail"]["settings"]["whitelist"]["total_list"])
    list_counter_number = int(list_counter.text.split(" ")[1])

    list_counter_update = driver.find_element_by_xpath(data["mail"]["settings"]["whitelist"]["total_list"])
    list_counter_number_update = int(list_counter_update.text.split(" ")[1])

    list_data = {
        "list_counter_number": list_counter_number,
        "list_counter_number_update": list_counter_number_update
    }
    return list_data

def alias_total():
    list_counter = driver.find_element_by_xpath(data["mail"]["settings_admin"]["aliasaccount"]["total_list"])
    Logging("=> Total list number: " + list_counter.text)
    list_counter_number = int(list_counter.text.split(" ")[1])

    return list_counter_number

def add_signature():
    input_text = data["mail"]["settings"]["signature"]["text3"]
    try:
        Logging("** Create signature")
        driver.find_element_by_xpath(data["mail"]["settings"]["signature"]["signature_access"]).click()
        Logging("- Click signature")
        time.sleep(5)

        Logging("- Add signature")
        ''' text '''
        Logging("- Add text signature")
        driver.find_element_by_xpath(data["mail"]["settings"]["signature"]["signature_add"]).click()
        Logging("- Click add signature")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["signature"]["text1"]).click()
        Logging("- Select text to signature")
        time.sleep(5)
        frame_task = driver.find_element_by_class_name("tox-edit-area__iframe")
        driver.switch_to.frame(frame_task)
        content = driver.find_element_by_xpath(data["mail"]["settings"]["signature"]["text2"])
        Logging("- Click content")
        content.clear()
        content.send_keys(input_text)
        Logging("- Add text")
        time.sleep(5)
        driver.switch_to.default_content()
        driver.find_element_by_xpath(data["mail"]["settings"]["signature"]["text_save"]).click()
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
        driver.find_element_by_xpath(data["mail"]["settings"]["signature"]["delete_signature1"]).click()
        Logging("- Select signature")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["signature"]["delete_signature2"]).click()
        Logging("=> Delete signature")
        time.sleep(5)
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_signature"]["pass"])
        #TestlinkResult_Pass("WUI-84")
    except WebDriverException:
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_signature"]["fail"])
        Logging("Delete signature fail")
        #TestlinkResult_Fail("WUI-84")

def add_autosort():
    ''' Auto sort '''
    Logging(" ")
    Logging("** Auto sort")
    try:
        driver.find_element_by_xpath(data["mail"]["settings"]["auto_sort"]["autosort"]).click()
        Logging("- Access Auto-Sort")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["auto_sort"]["addautosort"]).click()
        Logging("- Click add auto sort")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["auto_sort"]["input1"]).send_keys(data["mail"]["settings"]["auto_sort"]["input_from"])
        Logging("- Input from")       
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["auto_sort"]["input2"]).send_keys(data["mail"]["settings"]["auto_sort"]["input_to"])
        Logging("- Input to") 
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["auto_sort"]["input3"]).send_keys(data["mail"]["settings"]["auto_sort"]["input_subject"])
        Logging("- Input subject") 
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["auto_sort"]["select1"]).send_keys(data["mail"]["settings"]["auto_sort"]["select_mailbox"])
        Logging("- Select mail box") 
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["auto_sort"]["include_mail"]).click()
        Logging("- Select Include existing mail") 
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["auto_sort"]["save_button"]).click()
        Logging("=> Save auto sort") 
        #TestlinkResult_Pass("WUI-85")
        time.sleep(5)
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_auto_sort"]["pass"])
    except WebDriverException:
        Logging("Auto-Sort fail")  
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_auto_sort"]["fail"]) 
        #TestlinkResult_Fail("WUI-85")
    
    '''Logging("** Check auto sort create success")
    autosort = driver.find_element_by_xpath("//*[@id='ngw.mail.autosort']//table/tbody/tr/td")
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
        driver.find_element_by_xpath(data["mail"]["settings"]["auto_sort"]["autosort_delete"]).click()
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
    ''' Vacation auto replies '''
    Logging(" ")
    Logging("** Vacation auto replies")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["auto_replies"]["autoreplies"]))).click()
    Logging("- Access vacation auto replies")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["auto_replies"]["on/off_autoreplies"]))).click()
    Logging("- Click turn on vacation auto replies")
    time.sleep(5)

    driver.find_element_by_xpath(data["mail"]["settings"]["auto_replies"]["date_end"]).click()
    time.sleep(2)
    driver.find_element_by_css_selector(data["mail"]["settings"]["auto_replies"]["select_enddate"]).click() 
    Logging("- Set date end of auto reply")
    time.sleep(5)
    driver.find_element_by_xpath(data["mail"]["settings"]["auto_replies"]["input_text"]).send_keys(data["mail"]["settings"]["auto_replies"]["text_msg"])
    Logging("- Input message")
    time.sleep(5)
    driver.find_element_by_xpath(data["mail"]["settings"]["auto_replies"]["save_button"]).click()   
    Logging("=> Save turn on vacation auto replies")     
    time.sleep(5)

    Logging("** Check auto reply have turn on")
    button = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["auto_replies"]["on/off_autoreplies"])))
    if button.is_enabled():
        Logging("=> Create vacation auto replies success") 
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_auto_replies"]["pass"])
    else:
        Logging("=> Create vacation auto replies fail")
        ValidateFailResultAndSystem("<div>[Mail]Create vacation auto replies fail </div>")
    time.sleep(5)

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["auto_replies"]["on/off_autoreplies"]))).click()
    Logging("=> Turn off auto reply")
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["auto_replies"]["save_button"]))).click()
    Logging("=> Save turn off vacation auto replies") 

    time.sleep(5)

def add_block_address():
    addresses_block = data["mail"]["settings"]["blockaddress"]["input_address"]
    driver.find_element_by_tag_name("body").send_keys(Keys.END)
    time.sleep(5)
    Logging("** Add block")
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["blockaddress"]["block_address"]))).click()
        Logging("- Access Blocked Addresses")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["blockaddress"]["text_address"]).send_keys(addresses_block)
        Logging("- Input block addresses")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["blockaddress"]["add_mail"]).click()
        Logging("=> Add Blocked Addresses")
        time.sleep(5)

        Logging("** Check add block addresses success")
        block_name = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.mail.blockaddress']//td[contains(., '" + addresses_block + "')]")))
        if block_name.text == 'quynh2@qa1.hanbiro.net':
            Logging("=> Add Block success")
            TesCase_LogResult(**data["testcase_result"]["mail"]["add_block"]["pass"])
            #TestlinkResult_Pass("WUI-87")
        else:
            Logging("=> Block addresses fail")
            TesCase_LogResult(**data["testcase_result"]["mail"]["add_block"]["fail"])
            ValidateFailResultAndSystem("<div>Add block addresses fail </div>")
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
        driver.find_element_by_xpath(data["mail"]["settings"]["blockaddress"]["blocked_mail"]).click()
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["blockaddress"]["del_mail"]).click()
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
        driver.find_element_by_xpath((data["mail"]["settings"]["whitelist"]["white_list"])).click()
        Logging("- Access to white list")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["whitelist"]["addlisst"]).send_keys(whilelist)
        Logging("- Add white list")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["whitelist"]["add_button"]).click()
        Logging("=> Save white list")
        time.sleep(5)

        Logging("** Search while List")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["whitelist"]["searchwhitelist"]))).send_keys(data["mail"]["settings"]["whitelist"]["input_search_1"])
        Logging("Input white list")
        time.sleep(5)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["whitelist"]["search"]))).click()
        Logging("Search white list")
        time.sleep(5)
    except WebDriverException:
        pass
    
    try:
        Logging("** Check add while list succes")
        check_whilelist = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.mail.whitelist']//td[contains(., '" + whilelist + "')]")))
        if check_whilelist.is_displayed():
            Logging("=> Add while list success")
            TesCase_LogResult(**data["testcase_result"]["mail"]["add_whilelist"]["pass"])
        else:
            Logging("=> Add while list fail")
            TesCase_LogResult(**data["testcase_result"]["mail"]["add_whilelist"]["fail"])
            ValidateFailResultAndSystem("<div>[Mail]Add while list fail </div>")
        time.sleep(5)
    except WebDriverException:
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_whilelist"]["fail"])
        pass
    return whilelist

def delete_while_list():
    try:
        Logging("** Del while List")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["whitelist"]["refesh"]))).click()
        Logging("Refresh white list")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["whitelist"]["searchwhitelist"]))).send_keys(data["mail"]["settings"]["whitelist"]["input_search_1"])
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["whitelist"]["search"]).click()
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["whitelist"]["select_list"]).click()
        Logging("Select list to delete")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings"]["whitelist"]["del_list"]).click()
        Logging("Delete list")
        time.sleep(5)
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_whilelist"]["pass"])
    except WebDriverException:
        TesCase_LogResult(**data["testcase_result"]["mail"]["delete_whilelist"]["fail"])
        pass

def add_folder():
    name_folders = data["mail"]["settings"]["folders"]["name"] + str(n) 
    try:
        Logging(" ")
        Logging("** Folders")

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["st_folders"]))).click()
        Logging("- Access to Folders")
        time.sleep(2)
        Logging("- Create folder")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["add_folder"]))).click()
        time.sleep(2)
        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["input_name"]))).send_keys(name_folders)
        Logging("- Input name folder")
        time.sleep(2)
        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["button_save"][0]))).click()
        Logging("=> Create folders success")
        time.sleep(5)

        try:
            popup = driver.find_element_by_xpath(data["mail"]["settings"]["folders"]["pop_up"])
            if popup.is_displayed():
                Logging("Folder have exits - Create new folder")
                time.sleep(3)
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["popup_exits"]))).click()
                time.sleep(2)
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["add_folder"]))).click()
                time.sleep(2)
                WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["input_name"]))).send_keys(name_folders)
                Logging("- Input name folder")
                time.sleep(2)
                WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["button_save"][0]))).click()
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
        check_folder = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='mail_setting_form']//li//a[contains(., '" + name_folders + "')]")))
        if check_folder.is_displayed():
            Logging("=> Create folders success")
            TesCase_LogResult(**data["testcase_result"]["mail"]["add_folder"]["pass"])
        else:
            Logging("=> Create folders fail")
            TesCase_LogResult(**data["testcase_result"]["mail"]["add_folder"]["fail"])
            ValidateFailResultAndSystem("<div>[Mail]Create folders fail </div>")
        time.sleep(5)
    except WebDriverException:
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_folder"]["fail"])
        pass
    return name_folders

def share_folder(name_folders):
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='mail_setting_form']//li//a[contains(., '" + name_folders + "')]"))).click()
        time.sleep(3)
        Logging("- Select share permission")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["share"]))).click()
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["usingshare"]))).click()
        Logging("- Check using share")
        time.sleep(3)
        org = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["org_input"])))
        time.sleep(3)
        org.send_keys(data["mail"]["settings"]["folders"]["org_text"])
        time.sleep(3)
        org.send_keys(Keys.ENTER)
        Logging("- Input name user")
        time.sleep(3)
        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["org_select"]))).click()
        Logging("- Select user")
        time.sleep(3)
        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["org_add"]))).click()
        Logging("- Add user success")
        time.sleep(3)

        options_list = ["Read/Share/Reply/Forward", "Read Mail", "Shared Mail", "Reply/Forward", "Read/Share", "Read/Reply/Forward", "Share/Reply/Forward"]

        sel = Select(driver.find_element_by_xpath(data["mail"]["settings"]["folders"]["dropdown"]))
        time.sleep(2)
        sel.select_by_visible_text(random.choice(options_list))
        Logging("- Select permission for user")
        
        time.sleep(2)
        
        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["button_save"][1]))).click()
        Logging("=> Share folder success")
        time.sleep(5)

        Logging("** Check Share folder")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["share"]))).click()
        time.sleep(3)
        button_share = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["usingshare"])))
        if button_share.is_enabled():
            Logging("=> Share folder success")
        else:
            Logging("=> Share folder fail")
            ValidateFailResultAndSystem("<div>[Mail]Share folder fail </div>")
        time.sleep(5)
        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["button_close"]))).click()
        time.sleep(5)
    except:
        pass

def upload_eml(name_folders):
    driver.find_element_by_tag_name("body").send_keys(Keys.END)
    time.sleep(5)

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='mail_setting_form']//li//a[contains(., '" + name_folders + "')]"))).click()

    driver.find_element_by_xpath(data["mail"]["settings"]["folders"]["Eml"]).click()
    time.sleep(2)
    driver.find_element_by_xpath(data["mail"]["settings"]["folders"]["addfile"]).click()
    time.sleep(2)
    driver.find_element_by_xpath(data["mail"]["settings"]["folders"]["getfile"]).send_keys(NQ_login_function.file_upload)
    Logging("- Select file to upload")
    time.sleep(2)
    driver.find_element_by_xpath(data["mail"]["settings"]["folders"]["button_save"][2]).click()

    time.sleep(15)
    Logging("=> Upload EML success")

def backup_mailbox():
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["backup"]))).click()
    time.sleep(2)
    Logging("=> Backup success")

def empty_mailbox():
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["empty"]))).click()
    time.sleep(2)
    driver.find_element_by_xpath(data["mail"]["settings"]["folders"]["close_popup"]).click()
    Logging("=> Empty folder success")
    time.sleep(5)

def delete_folder():
    ''' Delete folders '''
    try:
        Logging("** Delete folder")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["delete"]))).click()
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
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["share_mailbox"]))).click()
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["folders"]["downfile"]))).click()
    time.sleep(2)
    Logging("=> Download share mail box success")
    time.sleep(5)   '''

def forwarding():
    email = data["mail"]["settings"]["forwarding"]["text_inpiut"]
    Logging(" ")
   
    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, "//a[contains(@data-defaulthref,'#/mail/forwarding/setting') and contains(.,' Forwarding')]"))).click()
    Logging("** Forwarding")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["forwarding"]["add_forwarding"]))).click()
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, data["mail"]["settings"]["forwarding"]["input_email"]))).send_keys(email)
    Logging("- Input email")

    option_list = int(len(driver.find_elements_by_xpath(data["mail"]["settings_admin"]["forwarding"]["option_list"])))
    forwarding_list = []
    i = 0
    for i in range(option_list):
        i += 1
        mode_list = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["forwarding"]["select_option"] + "[" + str(i) + "]")))
        forwarding_list.append(mode_list.text)
    
    Logging("- Total of view mode list: " + str(len(forwarding_list)))
    x = random.choice(forwarding_list)
    time.sleep(2)
    select_mode_view = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["forwarding"]["option_list"] + "[contains(.,'" + str(x) + "')]")))
    select_mode_view.click()
    Logging("- Select Forwarding option")

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["forwarding"]["save_button"]))).click()
    Logging("- Save forwarding")

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.mail.forwarding']//table//td[contains(., '" + email + "')]"))).click()
    Logging("- Select email")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings"]["forwarding"]["delete_button"]))).click()
    Logging("- Delete email have set forwarding")

def approval_total():
    list_counter = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["total_list"])
    list_counter_number = int(list_counter.text.split(" ")[1])

    return list_counter_number

def admin_settings():
    Logging(" ")
    Logging("** Open settings admin mail")
    Logging('============ Test case settings admin mail ============')

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
        delete_approval_mail_box()
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

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["alias_domian"]["aliasdomain"]))).click()
        Logging("- Access alias domain")
        time.sleep(5)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["alias_domian"]["input4"]))).send_keys(aliasdomain_name)
        Logging("- Input domain")
        time.sleep(5)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["alias_domian"]["add_aliasdomain"]))).click()
        Logging("- Add alias domain")
        time.sleep(5)
    except:
        pass
    
    try:
        Logging("** Check alias domain have add")
        name_domain = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.mail.adminalias_domain']//table[contains(., '" + aliasdomain_name + "')]")))
        if name_domain.is_displayed:
            Logging("=> Add alias domain success")
            TesCase_LogResult(**data["testcase_result"]["mail"]["alias_domain"]["pass"])
            #TestlinkResult_Pass("WUI-130")
        else:
            Logging("=> Add alias domain fail")
            TesCase_LogResult(**data["testcase_result"]["mail"]["alias_domain"]["fail"])
            #TestlinkResult_Fail("WUI-130")
            ValidateFailResultAndSystem("<div>[Mail]Add alias domain fail </div>")
        time.sleep(5) 
    except:
        Logging("=> Add alias domain fail")
        TesCase_LogResult(**data["testcase_result"]["mail"]["alias_domain"]["fail"])
        pass
    return aliasdomain_name

def delete_alias_domain(aliasdomain_name):
    Logging("** Delete alias domain")
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ngw.mail.adminalias_domain']//table[contains(., '" + aliasdomain_name + "')]//span"))).click()
        Logging("- Select alias domain")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["alias_domian"]["del_aliasdomain"]).click()
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
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["alias_domian"]["add_domain"]))).click()
        Logging("- Add user (group) for domain")
        add_user = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='aliasDomainOrg']//input[contains(@type, 'text')]")))
        add_user.send_keys(text)
        time.sleep(2)
        add_user.send_keys(Keys.ENTER)
        Logging("- Input key user")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["alias_domian"]["user_1"]))).click()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["alias_domian"]["user_2"]))).click()
        Logging("- Select user")
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["alias_domian"]["add_button"]))).click()
        Logging("- Add user")
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='directive-domains']//label"))).click()
        Logging("- Select domain")
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["alias_domian"]["save_button"]))).click()
        Logging("- Save")
    except:
        pass

def approval_mailbox():
    time.sleep(5)
    element_admin = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settingsmail"])))
    element_admin.location_once_scrolled_into_view

    
    element_admin1 = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settingsmail_1"])))
    element_admin1.location_once_scrolled_into_view

    Logging(" ")
    try:
        approval = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["mailbox"])
        if approval.is_displayed():
            Logging("Add approval mailbox")
            Forced_approval()
            Selective_approval()
    except WebDriverException:
        Logging("=> Domain don't have menu approval mail box")
        pass
    time.sleep(5)

def Forced_approval():
    driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["mailbox"]).click()
    Logging("- Select approval mailbox")
    f = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["policy"])
    driver.execute_script("arguments[0].scrollIntoView();",f)
    try:
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["policy"]).click()
        Logging("- Select approval policy")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["add_policy"]).click()
        Logging("- Add approval policy")
        time.sleep(5)
    #-----------------------------------------Forced approval----------------------------------#
        forced_approval = data["mail"]["settings_admin"]["approval_mailbox"]["input_name_1"] + str(n)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["input5"]).send_keys(forced_approval)

        Logging("** Select type policy Forced approval")
        time.sleep(5)
        Logging("- Select user final approval")
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["select_approver"]).click()
        Logging("- Select organization")
        time.sleep(5)
        final_approver = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["approver_input"])
        time.sleep(2)
        final_approver.send_keys(data["mail"]["settings_admin"]["approval_mailbox"]["approver_input_1"])
        time.sleep(2)
        final_approver.send_keys(Keys.ENTER)
        Logging("- Input user")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["approver_final_1"]).click()
        Logging("=> Add approval")
        time.sleep(5)

        #-------Permission Recipient----------#
        Logging("** Selective approval")
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["select_organization_1"]).click()
        Logging("- Select organization of Permission Recipient**")
        time.sleep(5)
        permission_recipient = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_input"])
        time.sleep(2)
        permission_recipient.send_keys(data["mail"]["settings_admin"]["approval_mailbox"]["organization_input_1"])
        time.sleep(2)
        permission_recipient.send_keys(Keys.ENTER)
        Logging("- Input user")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_select_1"]).click()
        Logging("- Select user")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_add_1"]).click()
        Logging("- Add user")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_save_1"]).click()
        Logging("=> Save Permission Recipient approver")
        time.sleep(5)
        
        #-------------Mid-Approver----------#
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["select_organization_2"]).click()
        time.sleep(5)
        Logging("** Select organization of mid approver**")
        
        mid_approver = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_input_2"])
        time.sleep(2)
        mid_approver.send_keys(data["mail"]["settings_admin"]["approval_mailbox"]["organization_input_1"])
        time.sleep(2)
        mid_approver.send_keys(Keys.ENTER)
        Logging("- Mid-Approver")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_select_2"]).click()
        Logging("- Select user")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_add_2"]).click()
        Logging("- Add user")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_save_2"]).click()
        Logging("- Save user")
        time.sleep(5)

        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["save"]).click()
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
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["add_policy"]).click()
        Logging("- Add approval policy")
        time.sleep(5)

        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["input5"]).send_keys(selective_approval)
        Logging("- Select basic policy")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["select_type"]).click()
        Logging("=> Selective approval")
        time.sleep(5)

        #-------Final approval----------#
        Logging("** Select final approval")
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["select_approver"]).click()
        Logging("- Select organization")
        time.sleep(5)
        final_approver = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["approver_input"])
        time.sleep(2)
        final_approver.send_keys(data["mail"]["settings_admin"]["approval_mailbox"]["approver_input_1"])
        time.sleep(2)
        final_approver.send_keys(Keys.ENTER)
        Logging("- Input final approval")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["approver_final_2"]).click()
        Logging("=> Select final approval")
        time.sleep(5)

        #-------Permission Recipient----------#
        Logging("** Select Permission Recipient")
        select_permission_recipient = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["select_organization_1"])
        select_permission_recipient.click()
        Logging("- Select organization of Permission Recipient")
        time.sleep(5)
        permission_recipient = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_input"])
        time.sleep(2)
        permission_recipient.send_keys(data["mail"]["settings_admin"]["approval_mailbox"]["organization_input_1"])
        time.sleep(2)
        permission_recipient.send_keys(Keys.ENTER)
        time.sleep(5)
        Logging("- Input user")
        approver_recipient = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_select_3"])
        approver_recipient.click()
        Logging("- Select user")
        time.sleep(5)
        add_approver_recipient = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_add_3"])
        add_approver_recipient.click()
        Logging("- Add user")
        time.sleep(5)
        save_approver_recipient = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_save_1"])
        save_approver_recipient.click()
        Logging("=> Save Permission Recipient approver")
        time.sleep(5)

        #-------Mid-Approver----------#
        Logging("** Select Mid-Approver")
        select_mid_approver = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["select_organization_2"])
        select_mid_approver.click()
        Logging("** Select organization of mid approver**")
        time.sleep(5)
        
        mid_approver = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_input_4"])
        time.sleep(2)
        mid_approver.send_keys(data["mail"]["settings_admin"]["approval_mailbox"]["organization_input_1"])
        time.sleep(2)
        mid_approver.send_keys(Keys.ENTER)
        Logging("- Input user")
        time.sleep(5)
        approver = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_select_4"])
        approver.click()
        Logging("- Select user")
        time.sleep(5)
        add_approver_recipient = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_add_4"])
        add_approver_recipient.click()
        Logging("- Add user")
        time.sleep(5)
        save_approver_recipient = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["organization_save_2"])
        save_approver_recipient.click()
        Logging("- Save Mid-Approver approver")
        time.sleep(5)

        save_approval_policy = driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["save"])
        save_approval_policy.click()
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
        driver.find_element_by_xpath((data["mail"]["settings_admin"]["approval_mailbox"]["select_policy_1"])).click()
        time.sleep(5)
        driver.find_element_by_xpath((data["mail"]["settings_admin"]["approval_mailbox"]["select_policy_2"])).click()
        Logging("-  Select policy to delete")
        time.sleep(5)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["remove_1"]).click()
        time.sleep(2)
        driver.find_element_by_xpath(data["mail"]["settings_admin"]["approval_mailbox"]["remove_2"]).click()
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
            send_limit.click()
            time.sleep(2)
            Logging("-> Asset send limit")
            time.sleep(5)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["sentlimit"]["input"][0]))).send_keys(text_domain)
            time.sleep(5)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["sentlimit"]["button_add"][0]))).click()
            Logging("** Create domain")
            time.sleep(5)

            Logging("** Check domain limt create success")
            domain = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='domain-limit']//td[contains(., '" + text_domain + "')]")))
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
                driver.find_element_by_xpath(data["mail"]["settings_admin"]["sentlimit"]["select"][0]).click()
                time.sleep(2)
                driver.find_element_by_xpath(data["mail"]["settings_admin"]["sentlimit"]["button_delete"]).click()
                time.sleep(2)
                driver.find_element_by_xpath(data["mail"]["settings_admin"]["sentlimit"]["close_popup"]).click()
                Logging("=> Delete domain")
                time.sleep(2)
                TesCase_LogResult(**data["testcase_result"]["mail"]["del_domain_limt"]["pass"])
            except:
                TesCase_LogResult(**data["testcase_result"]["mail"]["del_domain_limt"]["fail"])
                pass
            
            Logging("** Create file limit")
            driver.find_element_by_xpath(data["mail"]["settings_admin"]["sentlimit"]["file"]).click()
            time.sleep(2)
            driver.find_element_by_xpath(data["mail"]["settings_admin"]["sentlimit"]["input"][1]).send_keys(text_file)
            time.sleep(2)
            driver.find_element_by_xpath(data["mail"]["settings_admin"]["sentlimit"]["button_add"][1]).click()
            Logging("=> File limit have create")
            time.sleep(5)

            Logging("** Check file limit create success")
            file_limit = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='file-limit']//td[contains(., '" + text_file + "')]")))
            if file_limit.is_displayed():
                Logging("=> Create file limit success")
                TesCase_LogResult(**data["testcase_result"]["mail"]["add_file_limt"]["pass"])
            else:
                Logging("=> Create file limit fail")
                TesCase_LogResult(**data["testcase_result"]["mail"]["add_file_limt"]["fail"])
                ValidateFailResultAndSystem("<div>[Mail]Create file limit fail </div>")
            time.sleep(5)

            try:
                Logging("** Delete file limit")
                driver.find_element_by_xpath(data["mail"]["settings_admin"]["sentlimit"]["select"][1]).click()
                time.sleep(2)
                driver.find_element_by_xpath(data["mail"]["settings_admin"]["sentlimit"]["button_delete"]).click()
                time.sleep(2)
                driver.find_element_by_xpath(data["mail"]["settings_admin"]["sentlimit"]["close_popup"]).click()
                Logging("=> Delete file limit")
                TesCase_LogResult(**data["testcase_result"]["mail"]["del_file_limit"]["pass"])
            except:
                TesCase_LogResult(**data["testcase_result"]["mail"]["del_file_limit"]["fail"])
                pass
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
   
    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, "//a[contains(@data-defaulthref, '#/mail/admin/signature')]"))).click()
    
    time.sleep(5)

    force_list = int(len(driver.find_elements_by_xpath(data["mail"]["settings_admin"]["signature_company"]["force_list"])))
    force_apply_list = []
    i = 0
    for i in range(force_list):
        i += 1
        mode_list = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["signature_company"]["force_select"] + "[" + str(i) + "]")))
        force_apply_list.append(mode_list.text)
    
    Logging("- Total of view mode list: " + str(len(force_apply_list)))
    x = random.choice(force_apply_list)
    time.sleep(2)
    select_mode_view = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["signature_company"]["force_list"] + "[contains(.,'" + str(x) + "')]")))
    select_mode_view.click()
    Logging("- Select Force apply")
    
    if str(x) == "Forced-Appending":
        Logging("- Select mode view: Forced-Appending")
        time.sleep(2)

        position_list = int(len(driver.find_elements_by_xpath(data["mail"]["settings_admin"]["signature_company"]["signature_position"])))
        signature_position_list = []
        i = 0
        for i in range(position_list):
            i += 1
            signature_list = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["signature_company"]["position"] + "[" + str(i) + "]")))
            signature_position_list.append(signature_list.text)
        
        Logging("- Total of signature position: " + str(len(signature_position_list)))
        x = random.choice(signature_position_list)
        time.sleep(2)
        select_position = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["signature_company"]["signature_position"] + "[contains(.,'" + str(x) + "')]")))
        select_position.click()
        Logging("- Select Signature position")

        input_text = data["mail"]["settings_admin"]["signature_company"]["input_signature"]
        frame_task = driver.find_element_by_class_name("tox-edit-area__iframe")
        driver.switch_to.frame(frame_task)
        text = driver.find_element_by_id("tinymce")
        text.clear()
        text.send_keys(input_text)
        driver.switch_to.default_content()
        time.sleep(2)
        driver.find_element_by_tag_name("body").send_keys(Keys.END)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["signature_company"]["button_save"]))).click()
        Logging("- Add company signature")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_signature_company"]["pass"])
    elif str(x) == "Force replace":
        input_text = data["mail"]["settings_admin"]["signature_company"]["input_signature"]
        frame_task = driver.find_element_by_class_name("tox-edit-area__iframe")
        driver.switch_to.frame(frame_task)
        text = driver.find_element_by_id("tinymce")
        text.clear()
        text.send_keys(input_text)
        driver.switch_to.default_content()
        time.sleep(2)
        driver.find_element_by_tag_name("body").send_keys(Keys.END)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["signature_company"]["button_save"]))).click()
        Logging("- Add company signature")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_signature_company"]["pass"])
    elif str(x) == "None":
        input_text = data["mail"]["settings_admin"]["signature_company"]["input_signature"]
        frame_task = driver.find_element_by_class_name("tox-edit-area__iframe")
        driver.switch_to.frame(frame_task)
        text = driver.find_element_by_id("tinymce")
        text.clear()
        text.send_keys(input_text)
        driver.switch_to.default_content()
        time.sleep(2)
        driver.find_element_by_tag_name("body").send_keys(Keys.END)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["signature_company"]["button_save"]))).click()
        Logging("- Add company signature")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_signature_company"]["pass"])
    else:
        Logging("- Can't add signature company")
        TesCase_LogResult(**data["testcase_result"]["mail"]["add_signature_company"]["fail"])

def alias_account():
    name_account = data["mail"]["settings_admin"]["aliasaccount"]["input_nametext"] + str(n)
    name_alias = data["mail"]["settings_admin"]["aliasaccount"]["input_nameallias"] + str(n)
    ''' Alias account '''
    try:
        Logging(" ")
        Logging("** Alias account")
        driver.find_element_by_tag_name("body").send_keys(Keys.END)
        time.sleep(3)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["aliasaccount"]["allias"]))).click()
        Logging("- Access alias account")
        time.sleep(5)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["aliasaccount"]["input_name"]))).send_keys(name_account)
        Logging("- Input name account")
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["aliasaccount"]["input_allias"]))).send_keys(name_alias)
        Logging("- Input alias account")
        time.sleep(5)
        
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["aliasaccount"]["choose_org"]))).click()
        time.sleep(2)
        org = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["aliasaccount"]["input_org"])))
        time.sleep(2)
        org.send_keys(data["mail"]["settings_admin"]["aliasaccount"]["input_nameorg"])
        org.send_keys(Keys.ENTER)
        Logging("- Input user")
        time.sleep(5)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["aliasaccount"]["select_org"]))).click()
        Logging("- Select user")
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["aliasaccount"]["button_add_org"]))).click()
        Logging("- Add user")
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["aliasaccount"]["button_save_org"]))).click()
        Logging("=> Save Organization")
        time.sleep(5)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["aliasaccount"]["button_save"]))).click()
        Logging("=> Save alias account")
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["aliasaccount"]["close_popup"]))).click()
        time.sleep(5)
        TesCase_LogResult(**data["testcase_result"]["mail"]["alias_account"]["pass"])
    except:
        Logging("=> Can't create alias account")
        TesCase_LogResult(**data["testcase_result"]["mail"]["alias_account"]["fail"])
        pass

def log_analysis():
    try:
        ''' Log Analysis '''
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["log_analysis"]["loganalysis"]))).click()
        Logging("** Log Analysis **")
        time.sleep(2)
        #WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["log_analysis"]["filter_type"]))).click()
        Logging("- Select filter type")
        time.sleep(2)
        options_list = ["Received", "Send", "Spam", "Block", "Etc."]

        sel = Select(driver.find_element_by_xpath(data["mail"]["settings_admin"]["log_analysis"]["filter_type"]))
        sel.select_by_visible_text(random.choice(options_list))
        Logging("=> Select type ")
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["mail"]["settings_admin"]["log_analysis"]["search"]))).click()
        Logging("=> Search")
        time.sleep(2)

        type_select = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='mail_admin_log_analysis']//div[contains(@data-ng-show, 'isListView()')]")))
        type_select_text = type_select.text
        text = type_select_text.split(" ")[1]
        if text == '0':
            print("=> No data")
        else:
            print("=> Search suceess")
        time.sleep(5)
    except:
        pass