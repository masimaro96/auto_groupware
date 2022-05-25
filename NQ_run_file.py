import re, sys, json, openpyxl
import time, random, testlink
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, WebDriverException
from selenium.webdriver.support import expected_conditions as EC
from random import choice
from openpyxl import Workbook

from framework_sample import *
import NQ_login_function, NQ_todo_settings, NQ_mail_settings, NQ_archive_settings, NQ_calendar_settings, NQ_comanage_setting, NQ_new_comanage_settings
from NQ_login_function import execution_log, fail_log, error_log, Logging


def MyExecution(domain_name):
    error_menu = []

    try:
        NQ_login_function.access_qa(domain_name)
    except:
        Logging("Cannot continue execution")
        error_menu.append("NQ_login_function")

    try:
        NQ_mail_settings.mail(domain_name)
    except:
        Logging("Cannot continue execution")
        error_menu.append("NQ_mail_settings.mail")
    
    try:
        NQ_new_comanage_settings.new_co_manage(domain_name)
    except:
        Logging("Cannot continue execution")
        error_menu.append("NQ_comanage_setting.co_manage")

    try:
        NQ_new_comanage_settings.new_co_manage(domain_name)
    except:
        Logging("Cannot continue execution")
        error_menu.append("NQ_comanage_setting.new_co_manage")
    
    try:
        NQ_calendar_settings.calendar(domain_name)
    except:
        Logging("Cannot continue execution")
        error_menu.append("NQ_calendar_settings.calendar")
    
    try:
        NQ_archive_settings.archive(domain_name)
    except:
        Logging("Cannot continue execution")
        error_menu.append("NQ_archive_settings.archive")
    
    try:
        NQ_todo_settings.todo(domain_name)
    except:
        Logging("Cannot continue execution")
        error_menu.append("NQ_todo_settings.todo")
    
    nhuquynh_log = {
        "execution_log": execution_log,
        "fail_log": fail_log,
        "error_log": error_log,
        "error_menu": error_menu
    }

    return nhuquynh_log

def My_Execution(domain_name):
    NQ_login_function.access_qa(domain_name)
    # NQ_login_function.access_global3(domain_name)
    # NQ_new_comanage_settings.new_co_manage(domain_name)
    NQ_mail_settings.mail(domain_name)
    NQ_comanage_setting.co_manage(domain_name)
    NQ_calendar_settings.calendar(domain_name)
    NQ_archive_settings.archive(domain_name)
    NQ_todo_settings.todo(domain_name)

# My_Execution("http://qa.hanbiro.net/ngw/app/#")
My_Execution("http://qavn.hanbiro.net/ngw/app/#")
# My_Execution("http://gw.hanbirolinux.tk/ngw/app/#")
# My_Execution("http://global3.hanbiro.com/ngw/app/#")
# My_Execution("https://groupware57.hanbiro.net/ngw/app/#")
# My_Execution("http://global3.hanbiro.com/ncomanage/")

   
    
    
    
    


    
