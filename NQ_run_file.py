import re, sys, json, openpyxl
import time, random
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
import NQ_login_function, mail_settings #todo_settings, archive_settings, calendar_settings, comanage_setting, new_comanage_settings
from NQ_login_function import execution_log, Logging # , error_log #, fail_log


def MyExecution(domain_name):
    error_menu = []

    try:
        NQ_login_function.access_qa(domain_name)
    except:
        Logging("Cannot continue execution")
        error_menu.append("NQ_login_function")

    try:
        mail_settings.mail(domain_name)
    except:
        Logging("Cannot continue execution")
        error_menu.append("mail_settings.mail")
    
    # try:
    #     comanage_setting.co_manage(domain_name)
    # except:
    #     Logging("Cannot continue execution")
    #     error_menu.append("comanage_setting.co_manage")

    # try:
    #     new_comanage_settings.new_co_manage(domain_name)
    # except:
    #     Logging("Cannot continue execution")
    #     error_menu.append("new_comanage_settings.new_co_manage")

    # try:
    #     calendar_settings.calendar(domain_name)
    # except:
    #     Logging("Cannot continue execution")
    #     error_menu.append("calendar_settings.calendar")
    
    # try:
    #     archive_settings.archive(domain_name)
    # except:
    #     Logging("Cannot continue execution")
    #     error_menu.append("archive_settings.archive")
    
    # try:
    #     todo_settings.todo(domain_name)
    # except:
    #     Logging("Cannot continue execution")
    #     error_menu.append("todo_settings.todo")
    
    nhuquynh_log = {
        "execution_log": execution_log,
        # "fail_log": fail_log,
        # "error_log": error_log,
        # "error_menu": error_menu
    }

    return nhuquynh_log

def My_Execution(domain_name):
    NQ_login_function.access_qa(domain_name)
    # NQ_login_function.access_global3(domain_name)
    # new_comanage_settings.new_co_manage(domain_name)
    mail_settings.mail(domain_name)
    # comanage_setting.co_manage(domain_name)
    # calendar_settings.calendar(domain_name)
    # archive_settings.archive(domain_name)
    # todo_settings.todo(domain_name)

# My_Execution("http://qa.hanbiro.net/ngw/app/#")
# My_Execution("http://qavn.hanbiro.net/ngw/app/#")
# My_Execution("http://gw.hanbirolinux.tk/ngw/app/#")
# My_Execution("http://global3.hanbiro.com/ngw/app/#")
My_Execution("https://groupware57.hanbiro.net/ngw/app/#")
# My_Execution("http://global3.hanbiro.com/ncomanage/")
# My_Execution("http://myngoc.hanbiro.net/ngw/app/#")
