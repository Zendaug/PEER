# -*- coding: utf-8 -*-
"""
Created on Fri Mar 11 17:41:12 2022

@author: tbednall
"""

import tkinter as tk
from tkinter import messagebox
from tkinter.scrolledtext import ScrolledText

import forms.ToolTip as ToolTip
import forms.Settings as Settings
import os
import sys
import canvasapi

import forms.BulkMail as BulkMail
import forms.ExportGroups as ExportGroups
import forms.PeerMark as PeerMark

import modules.functions as func
import modules.globals as globals
import modules.graphQL as graphQL

def upload_template():
    if globals.config["SurveyPlatform"] == "Qualtrics":
        messagebox.showinfo("Create Survey", "To create the survey, upload the \"Peer_Evaluation_Survey_Template.qsf\" template to Qualtrics. Read the \"Instructions.txt\" file for more information.")
        func.open_file(os.path.abspath(os.path.dirname(sys.argv[0])) + "\\survey\\Qualtrics")
    elif globals.config["SurveyPlatform"] == "LimeSurvey":
        messagebox.showinfo("Create Survey", "To create the survey, upload the \"Peer_Evaluation_Survey_Template.lss\" template to LimeSurvey. Read the \"Instructions.txt\" file for more information.")
        func.open_file(os.path.abspath(os.path.dirname(sys.argv[0])) + "\\survey\\LimeSurvey")

class MainMenu:
    def __init__(self, master):        
        self.master = master
        self.master.title("PEER v" + str(globals.version))
        self.master.geometry("800x600")
        #self.window.geometry("300x250") # size of the window width:- 500, height:- 375
        #self.master.resizable(True, True) # this prevents from resizing the window
        self.master.bind("<FocusIn>", self.check_status)
        self.master.borderwidth = 5

        self.window = tk.Frame(self.master)
        self.window.pack(side="top", padx = 10, pady = 10, expand=True, fill="both")

        # create a pulldown menu, and add it to the menu bar
        self.menubar = tk.Menu(self.window)
        self.master.config(menu = self.menubar)

        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="Settings", command = Settings.Settings_call)
        self.filemenu.add_command(label="Version History", command = self.VersionHistory)
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Exit", command = self.exit_program)
        self.menubar.add_cascade(label="File", menu=self.filemenu)

        self.peermenu = tk.Menu(self.menubar, tearoff=0)

        self.Stage1 = tk.Menu(self.peermenu, tearoff=0)
        self.Stage1.add_command(label = "Confirm group membership", command=BulkMail.BulkMail_confirm)
        self.Stage1.add_command(label = "Contact students without groups", command = BulkMail.BulkMail_nogroup)
        self.peermenu.add_cascade(label="Stage 1: Group Formation", menu=self.Stage1)
        
        self.Stage2 = tk.Menu(self.peermenu, tearoff=0)
        self.Stage2.add_command(label = "Create " + globals.config["SurveyPlatform"] + " survey from template", command = upload_template)
        self.Stage2.add_command(label = "Create Contacts list from group membership", command = ExportGroups.ExportGroups_call)
        self.Stage2.add_command(label = "Send survey invitations from distribution list", command = BulkMail.BulkMail_invitation)
        self.peermenu.add_cascade(label = "Stage 2: Survey Launch", menu=self.Stage2)

        self.Stage3 = tk.Menu(self.peermenu, tearoff=0)
        self.Stage3.add_command(label = "Send reminder to complete survey", command=BulkMail.BulkMail_reminder)
        self.Stage3.add_command(label = "Upload peer marks and comments", command = PeerMark.PeerMark_call)
        self.peermenu.add_cascade(label = "Stage 3: Finalisation", menu=self.Stage3)
        
        self.menubar.add_cascade(label="Peer Evaluation", menu=self.peermenu)

        self.toolmenu = tk.Menu(self.menubar, tearoff = 0)
        self.toolmenu.add_command(label="Send Bulk Message", command = BulkMail.BulkMail_call)
        self.toolmenu.add_command(label="Download Class List", command = ExportGroups.ExportGroups_classlist)
        self.menubar.add_cascade(label="Tools", menu=self.toolmenu)

        self.text_area = ScrolledText(self.window, state="normal")
        self.text_area.pack(side=tk.TOP, pady=10, expand=True, fill="both")
        
        self.check_status()
                
        sys.stdout = TextRedirector(self.text_area, "stdout")
        sys.stderr = TextRedirector(self.text_area, "stderr")

    def check_status(self, event_type = ""):
        if globals.config["API_TOKEN"] == "" or globals.config["API_URL"] == "":
            self.Stage1.entryconfigure("Confirm group membership", state = "disabled")
            self.Stage1.entryconfigure("Contact students without groups", state = "disabled")
            self.Stage2.entryconfigure("Create " + globals.config["SurveyPlatform"] + " survey from template", state = "disabled")
            self.Stage2.entryconfigure("Create Contacts list from group membership", state = "disabled")
            self.Stage2.entryconfigure("Send survey invitations from distribution list", state = "disabled")
            self.Stage3.entryconfigure("Send reminder to complete survey", state="disabled")
            self.Stage3.entryconfigure("Upload peer marks and comments", state = "disabled")
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
            globals.canvas = canvasapi.Canvas(func.cleanupURL(globals.config["API_URL"]), globals.config["API_TOKEN"])
            globals.user_data = globals.canvas.get_user("self")
            self.Stage1.entryconfigure("Confirm group membership", state = "normal")
            self.Stage1.entryconfigure("Contact students without groups", state = "normal")
            self.Stage2.entryconfigure("Create " + globals.config["SurveyPlatform"] + " survey from template", state = "normal")
            self.Stage2.entryconfigure("Create Contacts list from group membership", state = "normal")
            self.Stage2.entryconfigure("Send survey invitations from distribution list", state = "normal")
            self.Stage3.entryconfigure("Send reminder to complete survey", state="normal")
            self.Stage3.entryconfigure("Upload peer marks and comments", state = "normal")
            self.toolmenu.entryconfigure("Send Bulk Message", state = 'normal')
            self.toolmenu.entryconfigure("Download Class List", state = 'normal')
            
    def VersionHistory(self):
        try:
            f = open("Version History.txt", "r")
            updates = f.read()
            f.close()
            print(updates)
        except:
            print("Unable to locate Version History.")
    
    def exit_program(self):
        self.master.destroy()
                
class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        #self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.see("end")            
        #self.widget.configure(state="disabled")
        
    def flush(self):
        pass

    
    