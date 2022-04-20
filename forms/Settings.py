# -*- coding: utf-8 -*-
"""
Created on Fri Mar 11 17:43:56 2022

@author: tbednall
"""

import json
import tkinter as tk
from tkinter import ttk

from modules.functions import *
import modules.globals as globals
import forms.ToolTip as ToolTip

class Settings:
    def __init__(self, window):
        # Again, try to load the configuration file and overwrite defaults
        try:
            f = open("config.txt", "r")
            user_config = json.loads(f.read())
            f.close()
            for item in user_config: globals.config[item] = user_config[item]
        except:
            pass
        
        self.window = window
        window.title("Settings")
        window.resizable(0, 0)
        window.grab_set()
        
        tk.Label(window, text = "\nPlease enter your API Key from Canvas:", fg = "black").grid(row = 2, column = 1, sticky = "W")

        tk.Label(window, text = "Get Token:", fg = "black").grid(row = 3, column = 0, sticky = "E")
        self.access_canvas = tk.Button(window, text = "Get Token from Canvas", fg = "black", command = self.open_canvas, state = "disabled")
        self.access_canvas.grid(row = 3, column = 1, sticky="W", padx = 5, pady = 5)
        ToolTip.CreateToolTip(self.access_canvas, text = 
              'Click on this button to access Canvas in your web browser.\n'+
              'From your profile page, click on the "+New Access Token" button.\n'+
              'Generate the token and copy/paste it into the box below.')

        tk.Label(window, text = "\nPlease enter the URL for the Canvas API: (e.g., https://swinburne.instructure.com/)", fg = "black").grid(row = 0, column = 1, sticky = "W")
        tk.Label(window, text = "API URL:", fg = "black").grid(row = 1, column = 0, sticky = "E")
        self.settings_URL = tk.Entry(window, width = 50)
        self.settings_URL.bind("<Key>", self.check_status)
        self.settings_URL.grid(row = 1, column = 1, sticky="W", padx = 5)
        ToolTip.CreateToolTip(self.settings_URL, text = 
              'The base URL used to log into Canvas.')

        tk.Label(window, text = "Enter Token:", fg = "black").grid(row = 4, column = 0, sticky = "E")
        self.settings_Token = tk.Entry(window, width = 80)
        self.settings_Token.grid(row = 4, column = 1, sticky="W", padx = 5)
        ToolTip.CreateToolTip(self.settings_Token, text = 
              'The Token allows PEER to log into Canvas on your behalf\n' +
              'and access information about students, groups and courses.')
        
        tk.Label(window, text = "Survey Platform:", fg = "black").grid(row = 5, column = 0, sticky = "E")
        self.survey_platform = ttk.Combobox(window, width = 60, state="readonly", values = ["LimeSurvey", "Qualtrics"])
        self.survey_platform.grid(row = 5, column = 1, sticky = "W", padx = 5, pady = 5)
        ToolTip.CreateToolTip(self.survey_platform, text = 
              'Select the survey platform used to run the PEER evalution.')
        
        tk.Label(window, text = "Email format:", fg = "black").grid(row = 6, column = 0, sticky = "E")
        self.settings_email = tk.Entry(window, width = 50)
        self.settings_email.grid(row = 6, column = 1, sticky="W", padx = 5)
        ToolTip.CreateToolTip(self.settings_Token, text = 
              'If student email addresses are not provided by Canvas this,\n' +
              'setting can be used to construct them from other user data.')
        
        tk.Label(window, text = "You will only need to enter these settings once.", fg = "black").grid(row = 7, columnspan = 2)
        self.settings_openfolder = tk.Button(window, text = "Open App Folder", fg = "black", command = self.open_folder).grid(row = 8, columnspan = 2, pady = 5)
        self.settings_default = tk.Button(window, text = "Restore Default Settings", fg = "black", command = self.restore_defaults).grid(row = 9, columnspan = 2, pady = 5)
        self.settings_save = tk.Button(window, text = "Save Settings", fg = "black", command = self.save_settings).grid(row = 10, columnspan = 2, pady = 5)
        self.update_fields()
        
        self.check_status()

    def open_canvas(self):
        open_url(cleanupURL(self.settings_URL.get()) + "profile/settings")

    def open_folder(self):
        open_file(os.path.abspath(os.path.dirname(sys.argv[0])))
        self.window.destroy()

    def update_fields(self):
        self.settings_URL.delete(0, tk.END)
        self.settings_URL.insert(0, globals.config["API_URL"])
        self.settings_Token.delete(0, tk.END)
        self.settings_Token.insert(0, globals.config["API_TOKEN"])
        self.survey_platform.set(globals.config["SurveyPlatform"])
        self.settings_email.delete(0, tk.END)
        self.settings_email.insert(0, globals.config["EmailFormat"])
        
    def restore_defaults(self):
        globals.config = {}
        for item in defaults: globals.config[item] = defaults[item]
        self.update_fields()
        
    def save_settings(self):
        globals.config["API_URL"] = cleanupURL(self.settings_URL.get())
        globals.config["API_TOKEN"] = self.settings_Token.get().strip()
        globals.config["SurveyPlatform"] = self.survey_platform.get()
        globals.config["EmailFormat"] = self.settings_email.get()
        save_config(globals.config)
        self.window.destroy()
        
    def check_status(self, key_press = ""):
        txt = self.settings_URL.get().lower()
        if type(key_press) is not str: txt = txt + key_press.char.lower()
        if len(txt) >= 17 and txt[-17:] == ".instructure.com/" and (txt[0:7]=="http://" or txt[0:8]=="https://"):
            self.access_canvas["state"] = "normal"
        else:
            self.access_canvas["state"] = "disabled"

def Settings_call():
    settings_win = tk.Toplevel()
    set_win = Settings(settings_win)