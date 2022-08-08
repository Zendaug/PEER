# -*- coding: utf-8 -*-
"""
Created on Fri Mar 11 17:54:28 2022

@author: tbednall
"""

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading

from modules.functions import *
import modules.globals as globals
import modules.server as server

class ExportGroups():
    def __init__(self, master, preset = ""):
        self.preset = preset
        self.master = master
        self.exportgroups = tk.Frame(self.master)
        self.exportgroups.pack()
        
        if preset=="classlist":
            self.master.title("Download Class List")
            tk.Label(self.exportgroups, text = "\nDownload list of students and teachers", fg = "black").grid(row = 0, column = 0, columnspan = 2)            
        else:
            self.master.title("Create Contact List")
            tk.Label(self.exportgroups, text = "\nDownload student groups and group membership", fg = "black").grid(row = 0, column = 0, columnspan = 2)

        self.master.resizable(0, 0)
        self.master.grab_set()

        tk.Label(self.exportgroups, text = "Unit:", fg = "black").grid(row = 1, column = 0, sticky = "E", pady = 5)           
        self.exportgroups_unit = ttk.Combobox(self.exportgroups, width = 60, values = [], state="readonly")
        self.exportgroups_unit.bind("<<ComboboxSelected>>", self.group_sets)
        self.exportgroups_unit.grid(row = 1, column = 1, sticky = "W", padx = 5)

        tk.Label(self.exportgroups, text = "Group Set:", fg = "black").grid(row = 2, column = 0, sticky = "E", pady = 5)
        self.exportgroups_groupset = ttk.Combobox(self.exportgroups, width = 60, state="readonly")
        self.exportgroups_groupset.grid(row = 2, column = 1, sticky = "W", padx = 5)
        
        self.exportgroups_export = tk.Button(self.exportgroups, text = "Start Download", fg = "black", command = lambda : threading.Thread(target=self.begin_export, daemon = True).start()) #self.begin_export2)
        self.exportgroups_export.grid(row = 4, columnspan = 2, pady = 5)

        self.group_sets_id = []
        self.group_sets_name = []          
        
        self.exportgroups_unit["values"] = tuple(globals.session["course.names"])
        self.exportgroups_unit.current(0)
        self.group_sets()
                
        if preset!="classlist" and globals.config["firsttime_export"] == True:
            messagebox.showinfo("Reminder", "Before creating a Contacts list for Qualtrics, it is recommended you check the accuracy of the group membership data in Canvas. Use the \"Confirm group membership\" function in the main menu to contact students to confirm their membership in each team.")
            globals.config["firsttime_export"] = False
            save_config(globals.config)
    
    def group_sets(self, event_type = ""):
        self.group_sets_id = []
        self.group_sets_name = []
        globals.GQL.course_id = globals.session["course.ids"][self.exportgroups_unit.current()]
        for group_set in globals.GQL.group_sets():
            self.group_sets_id.append(group_set["_id"])
            self.group_sets_name.append(group_set["name"])
        if len(self.group_sets_id) > 0:
            self.exportgroups_groupset["values"] = tuple(self.group_sets_name)
            self.exportgroups_groupset.current(0)
        else:
            self.exportgroups_groupset["values"] = ()
            self.exportgroups_groupset.set('')
                    
    def begin_export(self, event_type = ""):
        if self.preset=="classlist":
            filename = tk.filedialog.asksaveasfilename(initialfile=globals.session["course.names"][self.exportgroups_unit.current()], defaultextension=".xlsx", title = "Save Exported Data As...", filetypes = (("Excel files","*.xlsx"),("all files","*.*")))
        else:
            filename = tk.filedialog.asksaveasfilename(initialfile=globals.session["course.names"][self.exportgroups_unit.current()] +" (Contacts List)", defaultextension=".csv", title = "Save Contacts List As...", filetypes = (("Comma separated values files","*.csv"),("all files","*.*")))

        if filename is None or filename == "": return
        
        # Log the action
        server.log_action("Export student list", globals.session["course.names"][self.exportgroups_unit.current()], globals.session["course.ids"][self.exportgroups_unit.current()])

        globals.GQL.course_id = globals.session["course.ids"][self.exportgroups_unit.current()]
        if self.exportgroups_groupset.current() >= 0:
            globals.GQL.groupset_id = self.group_sets_id[self.exportgroups_groupset.current()]
            globals.session["group_set"] = self.group_sets_id[self.exportgroups_groupset.current()]    
        globals.session["course.id"] = globals.session["course.ids"][self.exportgroups_unit.current()]
        globals.session["course.name"] = globals.session["course.names"][self.exportgroups_unit.current()]
        print("\nAccessing " + globals.session["course.name"] + "...")        
        students = globals.GQL.students_comprehensive()
        print("Found {} students...".format(len(students)))

        if self.preset=="classlist":
            sections = globals.GQL.sections()
            teachers = globals.GQL.teachers_comprehensive()
            print("Found {} teachers...".format(len(teachers)))
            
            student_list = {"id": [],
                         "name":[],
                         "first_name":[],
                         "last_name":[],
                         "email":[],
                         "group_name":[],
                         "group_url":[]}
            
            teacher_list = {"id": [],
                         "name":[],
                         "first_name":[],
                         "last_name":[],
                         "email":[]}
            
            for section in sections:
                student_list[sections[section]["name"]] = []
                teacher_list[sections[section]["name"]] = []
            for student in students:
                student_list["id"].append(student)
                student_list["name"].append(students[student]["name"])
                student_list["first_name"].append(students[student]["first_name"])
                student_list["last_name"].append(students[student]["last_name"])
                student_list["email"].append(students[student]["email"])
                if "group_name" in students[student]:
                    student_list["group_name"].append(students[student]["group_name"])
                    student_list["group_url"].append(globals.config["API_URL"] + "groups/" + str(students[student]["group_id"]) + "/")
                else:
                    student_list["group_name"].append("")
                    student_list["group_url"].append("")
                for section in sections:
                    if section in students[student]["sections"]:
                        student_list[sections[section]["name"]].append(True)
                    else:
                        student_list[sections[section]["name"]].append(False)
            
            for teacher in teachers:
                teacher_list["id"].append(teacher)
                teacher_list["name"].append(teachers[teacher]["name"])
                teacher_list["first_name"].append(teachers[teacher]["first_name"])
                teacher_list["last_name"].append(teachers[teacher]["last_name"])
                teacher_list["email"].append(teachers[teacher]["email"])                
                for section in sections:
                    if section in teachers[teacher]["sections"]:
                        teacher_list[sections[section]["name"]].append(True)
                    else:
                        teacher_list[sections[section]["name"]].append(False)
            
            # New code to write to XLSX
            writeXL(filename=filename, datasets=[student_list, teacher_list], sheetnames=["Student List", "Teacher List"])
            messagebox.showinfo("Download complete.", "Student list saved as {}.".format(filename))
            
        else:            
            warning_zeropeers = []
            if globals.config["SurveyPlatform"] == "Qualtrics":
                student_list = {"ExternalDataReference":[],
                                "Student_name":[],
                                "FirstName":[],
                                "LastName":[],
                                "Email":[],
                                "Group_id":[], "Group_name":[],
                                "Peer1_id":[], "Peer1_name":[],
                                "Peer2_id":[], "Peer2_name":[],
                                "Peer3_id":[], "Peer3_name":[],
                                "Peer4_id":[], "Peer4_name":[],
                                "Peer5_id":[], "Peer5_name":[],
                                "Peer6_id":[], "Peer6_name":[],
                                "Peer7_id":[], "Peer7_name":[],
                                "Peer8_id":[], "Peer8_name":[],
                                "Peer9_id":[], "Peer9_name":[]}
                
                print("\nProcessing student details...")

                for student in students:
                    if ("group_id" in students[student]):
                        student_list["ExternalDataReference"].append(student)
                        student_list["Student_name"].append(students[student]["name"])
                        student_list["FirstName"].append(students[student]["first_name"])
                        student_list["LastName"].append(students[student]["last_name"])
                        student_list["Email"].append(students[student]["email"])
                        student_list["Group_id"].append(students[student]["group_id"])            
                        student_list["Group_name"].append(students[student]["group_name"])
                        print("> {}: {} ({} peers)".format(students[student]["name"], students[student]["group_name"], len(students[student]["peers"])))
                        for j, peer in enumerate(students[student]["peers"],1):
                            student_list["Peer" + str(j) + "_id"].append(peer)
                            student_list["Peer" + str(j) + "_name"].append(students[peer]["name"])
                        for k in range(len(students[student]["peers"])+1, 10):
                            student_list["Peer" + str(k) + "_id"].append(None)
                            student_list["Peer" + str(k) + "_name"].append(None)
                        if len(students[student]["peers"]) == 0: warning_zeropeers.append(students[student]["group_name"])
                    else:
                        print("> {}: No group found".format(students[student]["name"]))

            if globals.config["SurveyPlatform"] == "LimeSurvey":
                student_list = {"firstname":[],
                                "lastname":[],
                                "email":[],
                                "usesleft":[],
                                "attribute_1 <Group_name>":[],
                                "attribute_2 <Peer1_name>":[],
                                "attribute_3 <Peer2_name>":[],
                                "attribute_4 <Peer3_name>":[],
                                "attribute_5 <Peer4_name>":[],
                                "attribute_6 <Peer5_name>":[],
                                "attribute_7 <Peer6_name>":[],
                                "attribute_8 <Peer7_name>":[],
                                "attribute_9 <Peer8_name>":[],
                                "attribute_10 <Peer9_name>":[],
                                "attribute_11 <Student_id>":[],
                                "attribute_12 <Group_id>":[],
                                "attribute_13 <Peer1_id>":[],
                                "attribute_14 <Peer2_id>":[],
                                "attribute_15 <Peer3_id>":[],
                                "attribute_16 <Peer4_id>":[],
                                "attribute_17 <Peer5_id>":[],
                                "attribute_18 <Peer6_id>":[],
                                "attribute_19 <Peer7_id>":[],
                                "attribute_20 <Peer8_id>":[],
                                "attribute_21 <Peer9_id>":[]}
            
                for student in students:
                    if ("group_id" in students[student]):
                        student_list["attribute_11 <Student_id>"].append(student)
                        student_list["firstname"].append(students[student]["first_name"])
                        student_list["lastname"].append(students[student]["last_name"])
                        student_list["email"].append(students[student]["email"])
                        student_list["usesleft"].append(10)
                        student_list["attribute_12 <Group_id>"].append(students[student]["group_id"])            
                        student_list["attribute_1 <Group_name>"].append(students[student]["group_name"])
                        print("> {}: {} ({} peers)".format(students[student]["name"], students[student]["group_name"], len(students[student]["peers"])))
                        for j, peer in enumerate(students[student]["peers"],1):
                            student_list["attribute_" + str(j+12) + " <Peer" + str(j) + "_id>"].append(peer)
                            student_list["attribute_" + str(j+1) + " <Peer" + str(j) + "_name>"].append(students[peer]["name"])
                        for k in range(len(students[student]["peers"])+1, 10):
                            student_list["attribute_" + str(k+12) + " <Peer" + str(k) + "_id>"].append(None)
                            student_list["attribute_" + str(k+1) + " <Peer" + str(k) + "_name>"].append(None)
                        if len(students[student]["peers"]) == 0: warning_zeropeers.append(students[student]["group_name"])
                    else:
                        print("> {}: No group found".format(students[student]["name"]))
                        
            if len(warning_zeropeers) > 0:
                print("\nWarning: The following groups have only a single person in them (and no peers):")
                for group in warning_zeropeers:
                    print("> {}".format(group))
                            
            # Write CSV of group members
            print("\nCreating Contacts file to be uploaded to " + globals.config["SurveyPlatform"] + ".")
            writeCSV(student_list, filename)
            messagebox.showinfo("Download complete.", "Contacts file saved as \"{}\".".format(filename))

        self.master.destroy()


def ExportGroups_call():
    exportgroups = tk.Toplevel()
    exp_grp = ExportGroups(exportgroups)

def ExportGroups_classlist():
    exportgroups = tk.Toplevel()
    exp_grp = ExportGroups(exportgroups, preset="classlist")