# -*- coding: utf-8 -*-
"""
Created on Fri Mar 11 15:36:47 2022

@author: tbednall
"""

import requests
from datetime import datetime

import modules.globals as globals

def log_action(action, object_name, object_id):
    if globals.user_data is None: return()
    try:
        result = requests.post(url = "https://www.iotimlabs.com/PEER/updatelog.php",
                               data = {"date":datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                       "user_id":globals.user_data.id,
                                       "action":action,
                                       "object_name":object_name,
                                       "object_id":object_id})
    except:
        pass
    

def user_login():
    if globals.user_data is None: return()
    try:
        result = requests.post(url = "https://www.iotimlabs.com/PEER/userlogin.php",
                               data = {"user_id":globals.user_data.id,
                                       "user_name":globals.user_data.name,
                                       "api_url":globals.config["API_URL"] + "users/" + str(globals.user_data.id),
                                       "accessed":datetime.now().strftime('%Y-%m-%d %H:%M:%S')})
        return(result.json())
    except:
        pass