# -*- coding: utf-8 -*-
"""
Created on Fri Mar 11 16:02:32 2022

@author: tbednall
"""

import tkinter as tk
from tkinter import ttk

from modules.functions import *
import modules.globals as globals
import forms.ToolTip as ToolTip

class Policies():
    def __init__(self, master):
        self.master = master
        self.master.title("Calculate Peer Marks")
        self.master.resizable(0, 0)
        self.master.grab_set()
        
        self.policies = tk.Frame(self.master)
        self.policies.pack()
        self.validation = self.policies.register(only_numbers)
        self.validation2 = self.policies.register(only_posfloat)
    
                
        labelframe1 = tk.LabelFrame(self.policies, text = "Peer Mark Calculation Policy", fg = "black")
        labelframe1.grid(row = 0, column = 0, pady = 5, padx = 5, sticky = "W")
        
        self.policies_lab1 = tk.Label(labelframe1, text = "Scoring method:", fg = "black").grid(row = 0, column = 0, sticky = "E", pady = 5)
        self.policies_scoring = ttk.Combobox(labelframe1, width = 40, values = ["Teamwork mark (average rating received)", "Moderated group mark (based on peer ratings)"], state="readonly")
        self.policies_scoring.grid(row = 0, column = 1, columnspan = 2, sticky = "W", padx = 5)
        self.policies_scoring.current(globals.config["ScoreType"]-1)
        ToolTip.CreateToolTip(self.policies_scoring, text =
                       '"Teamwork mark" means each student receives their average mark from their peers.\n' +
                       '"Moderated group mark" means the student receives their team mark, which is then adjusted based on the peer ratings they receive.')
        
        tk.Label(labelframe1, text = "Self-scoring policy:", fg = "black").grid(row = 1, column = 0, sticky = "E", pady = 5)
        self.policies_selfrat = ttk.Combobox(labelframe1, width = 40, values = ["Do not count self-rating", "Count self-rating", "Subsitute self-rating with average rating given", "Cap self-rating at average rating given"], state="readonly")
        self.policies_selfrat.grid(row = 1, column = 1, columnspan = 2, sticky = "W", padx = 5)
        self.policies_selfrat.current(globals.config["SelfVote"])
        ToolTip.CreateToolTip(self.policies_selfrat, text =
                       'This option is used to specify whether to count each student\'s self-rating\n'+
                       'as part of their peer mark.')
        
        tk.Label(labelframe1, text = "Minimum peer responses:", fg = "black").grid(row = 2, column = 0, sticky = "E", pady = 5)
        self.policies_minimum = tk.Entry(labelframe1, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.policies_minimum.grid(row = 2, column = 1, padx = 5, sticky = "W")
        self.policies_minimum.insert(0, str(globals.config["MinimumPeers"]))
        ToolTip.CreateToolTip(self.policies_minimum, text =
                       'Specifies the minimum number of peers required to calculate the peer rating.\n' +
                       'Under the "Adjusted score" method, the student receives their unadjusted team mark.')

        tk.Label(labelframe1, text = "If minimum is not reached:", fg = "black").grid(row = 3, column = 0, sticky = "E", pady = 5)
        self.policies_minpolicy = ttk.Combobox(labelframe1, width = 40, values = ["Set Peer Mark status to 'Excused'", "Substitute Peer Mark with shared team mark", "Assign Peer Mark of zero"], state="readonly")
        self.policies_minpolicy.grid(row = 3, column = 1, columnspan = 2, sticky = "W", padx = 5)
        self.policies_minpolicy.current(globals.config["MinimumPeersPolicy"]-1)
        ToolTip.CreateToolTip(self.policies_minimum, text =
                       'This option is used to control what happens if the minimum number of peers is not reached.\n' +
                       "The 'Excused' option means that the Peer Mark will not be counted towards students' final grade, but points are not deducted." +
                       "The 'shared team mark' option means that students will receive the (unadjusted) team mark for their Peer Mark." +
                       "The 'Assign Zero' option means that students receive a 0 for the Peer Mark.")

        self.feedback_only = tk.IntVar()
        self.feedback_only.set(int(globals.config["feedback_only"]))
        tk.Label(labelframe1, text = "Feedback only?", fg = "black").grid(row = 4, column = 0, sticky = "NE")
        self.fbo = ttk.Checkbutton(labelframe1, text = "Peer mark does not count towards final grade", variable = self.feedback_only)
        self.fbo.grid(row = 4, column = 1, sticky = "W")
        ToolTip.CreateToolTip(self.fbo, text =
                       'If this option is enabled, the peer mark will not count towards students\'s final grades.')

        labelframe2 = tk.LabelFrame(self.policies, text = "Penalties Policy", fg = "black")
        labelframe2.grid(row = 1, column = 0, pady = 5, padx = 5, sticky = "W")
        tk.Label(labelframe2, text = "Apply % penalty", fg = "black").grid(row = 0, column = 2, sticky = "W")
        
        tk.Label(labelframe2, text = "Students do not complete the peer evaluation", fg = "black").grid(row = 1, rowspan = 2, column = 0, sticky = "NE")
        tk.Label(labelframe2, text = "\n", fg = "black").grid(row = 1, rowspan = 2, column = 1, sticky = "NE")
        self.penalty_noncomplete = tk.Entry(labelframe2, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.penalty_noncomplete.grid(row = 1, column = 2, sticky="W", padx = 5, pady = 5)
        self.penalty_noncomplete.insert(0, globals.config["Penalty_NonComplete"])
        ToolTip.CreateToolTip(self.penalty_noncomplete, text =
                       'Automatically deduct % if the student does not answer any questions in the peer evaluation.')

        tk.Label(labelframe2, text = "Students only partially complete the peer evaluation", fg = "black").grid(row = 2, rowspan = 2, column = 0, sticky = "NE")
        tk.Label(labelframe2, text = "\n", fg = "black").grid(row = 2, rowspan = 2, column = 1, sticky = "NE")
        self.penalty_partialcomplete = tk.Entry(labelframe2, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.penalty_partialcomplete.grid(row = 2, column = 2, sticky="W", padx = 5, pady = 5)
        self.penalty_partialcomplete.insert(0, globals.config["Penalty_PartialComplete"])
        ToolTip.CreateToolTip(self.penalty_partialcomplete, text =
                       'Automatically deduct % if the student completes only some of the peer evaluation questions but does not finish the survey.')
        
        tk.Label(labelframe2, text = "Students submit the peer evaluation late", fg = "black").grid(row = 3, rowspan = 2, column = 0, sticky = "NE")
        self.custom_latepenalty = tk.IntVar()
        self.custom_latepenalty.set(int(globals.config["Penalty_Late_Custom"]))
        ttk.Radiobutton(labelframe2, text = "Use default Canvas penalty", variable = self.custom_latepenalty, value = 0, command = self.change_customlatepenalty).grid(row = 3, column = 1, sticky = "NW")
        ttk.Radiobutton(labelframe2, text = "Custom late penalty", variable = self.custom_latepenalty, value = 1, command = self.change_customlatepenalty).grid(row = 4, column = 1, sticky = "NW")
        tk.Label(labelframe2, text = "(per day)", fg = "black").grid(row = 4, column = 2, sticky = "NE")
        self.penalty_late = tk.Entry(labelframe2, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.penalty_late.grid(row = 4, column = 2, sticky="NW", padx = 5, pady = 5)
        self.penalty_late.insert(0, globals.config["Penalty_PerDayLate"])
        ToolTip.CreateToolTip(self.penalty_late, text =
                       'Automatically deduct % for each day the peer evaluation is submitted after the due date.')
        
        tk.Label(labelframe2, text = "Students give themselves perfect scores", fg = "black").grid(row = 5, rowspan = 2, column = 0, sticky = "NE")
        self.perfect_score = tk.IntVar()
        ttk.Radiobutton(labelframe2, text = "Retain scores", variable = self.perfect_score, value = 0).grid(row = 5, column = 1, sticky = "NW")
        ttk.Radiobutton(labelframe2, text = "Exclude scores", variable = self.perfect_score, value = 1).grid(row = 6, column = 1, sticky = "NW")
        self.perfect_score.set(globals.config["Exclude_SelfPerfect"])
        self.penalty_selfperfect = tk.Entry(labelframe2, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.penalty_selfperfect.grid(row = 5, column = 2, sticky="W", padx = 5)
        self.penalty_selfperfect.insert(0, globals.config["Penalty_SelfPerfect"])
        ToolTip.CreateToolTip(self.penalty_selfperfect, text =
                       'Automatically deduct % if the student gives themselves a perfect self-rating across all questions.')
        
        tk.Label(labelframe2, text = "Students give all peers the bottom score", fg = "black").grid(row = 7, rowspan = 2, column = 0, sticky = "NE")
        self.bottom_score = tk.IntVar()
        ttk.Radiobutton(labelframe2, text = "Retain scores", variable = self.bottom_score, value = 0).grid(row = 7, column = 1, sticky = "NW")
        ttk.Radiobutton(labelframe2, text = "Exclude scores", variable = self.bottom_score, value = 1).grid(row = 8, column = 1, sticky = "NW")
        self.bottom_score.set(globals.config["Exclude_PeersAllZero"])
        self.penalty_allzero = tk.Entry(labelframe2, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.penalty_allzero.grid(row = 7, column = 2, sticky="W", padx = 5)
        self.penalty_allzero.insert(0, globals.config["Penalty_PeersAllZero"])
        ToolTip.CreateToolTip(self.penalty_allzero, text =
                       'Automatically deduct % if the student gives all of their peers the lowest rating across all questions.')        
        
        labelframe3 = tk.LabelFrame(self.policies, text = "Additional Settings", fg = "black")
        labelframe3.grid(row = 2, column = 0, pady = 5, padx = 5, sticky = "W")

        tk.Label(labelframe3, text = "Peer mark online survey scale ranges from 1 to:", fg = "black").grid(row = 0, column = 0, sticky = "NE")
        self.points_possible = tk.Entry(labelframe3, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.points_possible.grid(row = 0, column = 1, sticky="W", padx = 5)
        self.points_possible.insert(0, globals.config["PointsPossible"])
        ToolTip.CreateToolTip(self.points_possible, text =
                       'Only change this setting if you modify the scale of the peer evaluation survey questions.\n'+
                       'For example: you change the questions from a 1 to 5 scale to a 1 to 7 scale.')

        tk.Label(labelframe3, text = "Only moderate Peer Mark if adjustment greater than (%):", fg = "black").grid(row = 1, column = 0, sticky = "NE")
        self.min_adj = tk.Entry(labelframe3, width = 5, validate="key", validatecommand=(self.validation2, '%S'))
        self.min_adj.grid(row = 1, column = 1, sticky="W", padx = 5)
        self.min_adj.insert(0, globals.config["MinimumAdjustment"])
        ToolTip.CreateToolTip(self.min_adj, text =
                       'For the moderated scoring method, this setting prevents an adjustment from being made,\nunless it exceeds a threshold.')                     
        if self.policies_scoring.current()==0: self.min_adj["state"] = "disabled"
        
        tk.Label(labelframe3, text = "Assignment name:", fg = "black").grid(row = 2, column = 0, sticky = "E", pady = 5)
        self.pm_name = tk.Entry(labelframe3, width = 40)
        self.pm_name.grid(row = 2, column = 1, padx = 5, sticky = "W")
        self.pm_name.insert(0, str(globals.config["PeerMark_Name"]))
        ToolTip.CreateToolTip(self.pm_name, text =
                       'The name that the Peer Mark will be given in the Canvas grade book.')

        self.publish_assignment = tk.IntVar()
        self.publish_assignment.set(int(globals.config["publish_assignment"]))
        tk.Label(labelframe3, text = "Publish Peer Mark column in Grade Book?", fg = "black").grid(row = 3, column = 0, sticky = "NE")
        self.pub_assn = ttk.Checkbutton(labelframe3, text = "Make the column visible but hide grades", variable = self.publish_assignment)
        self.pub_assn.grid(row = 3, column = 1, sticky = "W")
        ToolTip.CreateToolTip(self.pub_assn, text =
                       'Whether the Peer Evaluation will be published in Canvas.')

        self.weekdays_only = tk.IntVar()
        self.weekdays_only.set(int(globals.config["weekdays_only"]))
        tk.Label(labelframe3, text = "Exclude weekends?", fg = "black").grid(row = 4, column = 0, sticky = "NE")
        self.wko = ttk.Checkbutton(labelframe3, text = "Only count weekdays when calculating late penalty", variable = self.publish_assignment)
        self.wko.grid(row = 4, column = 1, sticky = "W")
        ToolTip.CreateToolTip(self.wko, text =
                       'When calculating penalties for late submission of the Peer Evaluation, only count weekdays.')

        self.export_data = tk.IntVar()
        self.export_data.set(int(globals.config["ExportData"]))
        tk.Label(labelframe3, text = "Export peer data?", fg = "black").grid(row = 5, column = 0, sticky = "NE")
        self.export = ttk.Checkbutton(labelframe3, text = "Export peer ratings to an XLSX file", variable = self.export_data)
        self.export.grid(row = 5, column = 1, sticky = "W")
        ToolTip.CreateToolTip(self.export, text =
                       'After the peer marks are calculated, export them to an Excel file.\n' +
                       'This format is useful for inspecting the data and performing data analyses.')
        
        self.upload_data = tk.IntVar()
        self.upload_data.set(int(globals.config["UploadData"]))
        tk.Label(labelframe3, text = "Upload peer data?", fg = "black").grid(row = 6, column = 0, sticky = "NE")
        self.upload = ttk.Checkbutton(labelframe3, text = "Upload peer marks and feedback to Canvas", variable = self.upload_data)
        self.upload.grid(row = 6, column = 1, sticky = "W")
        ToolTip.CreateToolTip(self.wko, text =
                       'Upload the peer marks and feedback to the Grade book in Canvas.\n' +
                       'This creates a new data column (it does not overwrite existing grades).')
        
        self.save_policies = tk.IntVar()
        self.save_policies.set(1)
        tk.Label(labelframe3, text = "Retain policies?", fg = "black").grid(row = 7, column = 0, sticky = "NE")
        ttk.Checkbutton(labelframe3, text = "Save policies for future assessments", variable = self.save_policies).grid(row = 7, column = 1, sticky = "W")
        
        tk.Button(self.policies, text = "Continue", fg = "black", command = self.save_settings).grid(row = 13, column = 0, columnspan = 3, pady = 5)
        
        self.policies_scoring.bind("<<ComboboxSelected>>", self.change_assnname)
        self.change_customlatepenalty(self)
    
    def change_assnname(self, event_list = ""):
        self.pm_name.delete(0, 'end')
        if self.policies_scoring.current()==0:
            self.pm_name.insert(0, "Teamwork Mark (average rating received)") # Average mark received
            self.policies_minpolicy.current(0)
            self.min_adj["state"] = "disabled"
        else:
            self.pm_name.insert(0, "Moderated Group Mark (based on peer ratings)")
            self.policies_minpolicy.current(1)
            self.min_adj["state"] = "normal"
            
    def change_customlatepenalty(self, event_list = ""):
        if self.custom_latepenalty.get() == 0:
            self.penalty_late["state"] = "disabled"
        else:
            self.penalty_late["state"] = "normal"
    
    def save_settings(self):
        globals.config["ScoreType"] = int(self.policies_scoring.current()+1)
        globals.config["SelfVote"] = int(self.policies_selfrat.current())
        globals.config["MinimumPeers"] = int(self.policies_minimum.get())
        globals.config["MinimumPeersPolicy"] = int(self.policies_minpolicy.current()+1)
        globals.config["PeerMark_Name"] = str(self.pm_name.get().strip())
        globals.config["Penalty_NonComplete"] = int(self.penalty_noncomplete.get())
        globals.config["Penalty_PartialComplete"] = int(self.penalty_partialcomplete.get())
        globals.config["Penalty_PerDayLate"] = int(self.penalty_late.get())
        globals.config["Penalty_SelfPerfect"] = int(self.penalty_selfperfect.get())
        globals.config["Penalty_PeersAllZero"] = int(self.penalty_allzero.get())
        globals.config["PointsPossible"] = int(self.points_possible.get())
        globals.config["MinimumAdjustment"] = cleanupNumber(self.min_adj.get())
        globals.config["publish_assignment"] = bool(self.publish_assignment.get())
        globals.config["weekdays_only"] = bool(self.weekdays_only.get())
        globals.config["ExportData"] = bool(self.export_data.get())
        globals.config["UploadData"] = bool(self.upload_data.get())
        globals.config["Penalty_Late_Custom"] = bool(self.custom_latepenalty.get())    
        globals.config["feedback_only"] = bool(self.feedback_only.get())
    
        globals.adj_dict = {}
        
        if self.save_policies.get() == 1: save_config(globals.config)
        self.master.destroy()
        
    

        

           
def Policies_call():
    policies = tk.Toplevel()
    pol = Policies(policies)

