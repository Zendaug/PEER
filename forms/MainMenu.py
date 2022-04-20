# -*- coding: utf-8 -*-
"""
Created on Fri Mar 11 17:41:12 2022

@author: tbednall
"""

import tkinter as tk
from tkinter import messagebox

import forms.ToolTip as ToolTip
import forms.Settings as Settings
import os
import sys
import canvasapi

import forms.BulkMail as BulkMail
import forms.ExportGroups as ExportGroups
import forms.PeerMark as PeerMark

from modules.functions import *
import modules.globals as globals
import modules.graphQL as graphQL

def upload_template():
    messagebox.showinfo("Create Survey", "To create the survey, upload the \"Peer_Evaluation_Qualtrics_Template.qsf\" template to Qualtrics. Read the \"Instructions.txt\" file for more information.")
    open_file(os.path.abspath(os.path.dirname(sys.argv[0])) + "\\survey")


class MainMenu:
    def __init__(self, master):
        self.master = master
        self.master.title("PEER")
        #self.window.geometry("300x250") # size of the window width:- 500, height:- 375
        self.master.resizable(0, 0) # this prevents from resizing the window
        self.master.bind("<FocusIn>", self.check_status)

        self.window = tk.Frame(self.master)
        self.window.pack(side="top", pady = 0, padx = 50)

        # create a pulldown menu, and add it to the menu bar
        self.menubar = tk.Menu(self.window)
        self.master.config(menu = self.menubar)

        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="Settings", command = Settings.Settings_call)
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Exit", command = self.exit_program)
        self.menubar.add_cascade(label="File", menu=self.filemenu)              

        self.toolmenu = tk.Menu(self.menubar, tearoff = 0)
        self.toolmenu.add_command(label="Send Bulk Message", command = BulkMail.BulkMail_call)
        self.toolmenu.add_command(label="Download Class List", command = ExportGroups.ExportGroups_classlist)
        self.menubar.add_cascade(label="Tools", menu=self.toolmenu)

        labelframe1 = tk.LabelFrame(self.window, text = "Group Formation", fg = "black")
        labelframe1.pack(pady = 10)

        self.button_cgm = tk.Button(labelframe1, text = "Confirm group\nmembership", fg = "black", command = BulkMail.BulkMail_confirm, width = 20)
        self.button_cgm.pack(pady = 10, padx = 10)# 'text' is used to write the text on the Button
        ToolTip.CreateToolTip(self.button_cgm, text = 
                      'Contact students via the Canvas bulk mailer to confirm\n'+
                      'their membership in each group, as well as their peers.')

        self.button_ng = tk.Button(labelframe1, text = "Contact students\nwithout groups", fg = "black", command = BulkMail.BulkMail_nogroup, width = 20)
        self.button_ng.pack(pady = 10, padx = 10)# 'text' is used to write the text on the Button
        ToolTip.CreateToolTip(self.button_ng, text = 
                      'Contact students via the Canvas bulk mailer\n'+
                      'who are not currently enrolled in groups.')

        labelframe2 = tk.LabelFrame(self.window, text = "Survey Launch", fg = "black")
        labelframe2.pack(pady = 10)
        
        self.button_tp = tk.Button(labelframe2, text = "Create " + globals.config["SurveyPlatform"] + " survey\nfrom template", fg = "black", command = upload_template, width = 20)
        self.button_tp.pack(pady = 10, padx = 10)# 'text' is used to write the text on the Button
        ToolTip.CreateToolTip(self.button_tp, text = 
                      'Upload the Peer Evaluation template\n'+
                      'to Qualtrics to create the survey.')
        
        self.button_dl = tk.Button(labelframe2, text = "Create Contacts list\nfrom group membership", fg = "black", command = ExportGroups.ExportGroups_call, width = 20)
        self.button_dl.pack(pady = 10, padx = 10)# 'text' is used to write the text on the Button
        ToolTip.CreateToolTip(self.button_dl, text = 
                      'Export group membership data from Canvas and create\n'+
                      'a Contacts list file to be uploaded to Qualtrics.')

        self.button_si = tk.Button(labelframe2, text = "Send survey invitations\nfrom distribution list", fg = "black", command = BulkMail.BulkMail_invitation, width = 20)
        self.button_si.pack(pady = 10, padx = 10)# 'text' is used to write the text on the Button
        ToolTip.CreateToolTip(self.button_si, text = 
                      'Send students a unique link to the Peer Evaluation from a\n'+ 
                      'distribution list using the Canvas bulk mailer.')

        labelframe3 = tk.LabelFrame(self.window, text = "Peer Evaluation Finalisation", fg = "black")
        labelframe3.pack(pady = 10)        
    
        self.button_sr = tk.Button(labelframe3, text = "Send survey reminders\nfrom distribution list", fg = "black", command = BulkMail.BulkMail_reminder, width = 20)
        self.button_sr.pack(pady = 10, padx = 10)# 'text' is used to write the text on the Button
        ToolTip.CreateToolTip(self.button_sr, text = 
                      'Send students who haven\'t completed the Peer Evaluation a reminder to\n'+
                      'complete it (from a Qualtrics distribution list via the Canvas bulk mailer).')

        self.button_pm = tk.Button(labelframe3, text = "Upload peer marks\nand comments", fg = "black", command = PeerMark.PeerMark_call, width = 20)
        self.button_pm.pack(pady = 10)# 'text' is used to write the text on the Button     
        ToolTip.CreateToolTip(self.button_pm, text = 
                      'Calculate peer marks from the Qualtrics survey data and upload\n'+ 
                      'these marks and peer feedback to the Canvas grade book.')

        self.check_status()
                
        if globals.config["API_TOKEN"] == "" or globals.config["API_URL"] == "":
            messagebox.showinfo("Warning", "No configuration found. Please go to File -> Settings to set up a link to Canvas.")            
    
    def check_status(self, event_type = ""):
        if globals.config["API_TOKEN"] == "" or globals.config["API_URL"] == "":
            self.button_dl["state"] ='disabled'
            self.button_pm["state"] ='disabled'
            self.button_ng["state"] = 'disabled'
            self.button_cgm["state"] = 'disabled'
            self.button_si["state"] = 'disabled'
            self.button_sr["state"] = 'disabled'
            self.button_pm["state"] = 'disabled'
            self.toolmenu.entryconfigure("Send Bulk Message", state = 'disabled')
            self.toolmenu.entryconfigure("Download Class List", state = 'disabled')
        else:
            globals.GQL = graphQL.GraphQL()
            if "course.names" not in globals.session:
                globals.session["course.names"] = []
                globals.session["course.ids"] = []
                courses = globals.GQL.courses()
                for course in courses:
                    globals.session["course.names"].append(courses[course]["name"])
                    globals.session["course.ids"].append(course)
            globals.canvas = canvasapi.Canvas(cleanupURL(globals.config["API_URL"]), globals.config["API_TOKEN"])
            globals.user_data = globals.canvas.get_user("self")
            self.button_dl["state"] ='normal'
            self.button_pm["state"] ='normal'
            self.button_ng["state"] = 'normal'
            self.button_cgm["state"] = 'normal'
            self.button_si["state"] = 'normal'
            self.button_sr["state"] = 'normal'
            self.button_pm["state"] = 'normal'
            self.toolmenu.entryconfigure("Send Bulk Message", state = 'normal')
            self.toolmenu.entryconfigure("Download Class List", state = 'normal')
            
    def exit_program(self):
        self.master.destroy()