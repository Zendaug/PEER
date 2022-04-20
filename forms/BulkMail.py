# -*- coding: utf-8 -*-
"""
Created on Fri Mar 11 17:49:46 2022

@author: tbednall
"""

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

from modules.functions import *
import modules.globals as globals
import modules.server as server

class BulkMail:
    def __init__(self, master, preset = "", subject = "", message = ""):
        self.preset = preset
        self.master = master
        self.master.title("Send Bulk Message via Canvas")
        self.master.resizable(0, 0)
        
        self.master.bind("<FocusIn>", self.check_focus)
        
        self.window = tk.Frame(self.master)
        self.window.pack()
        
        tk.Label(self.window, text = "Unit:", fg = "black").grid(row = 1, column = 0, sticky = "E", pady = 5)
        self.window_unit = ttk.Combobox(self.window, width = 60, state="readonly")
        self.window_unit.grid(row = 1, column = 1, sticky = "W", padx = 5)
        self.window_unit["values"] = tuple(globals.session["course.names"])
        self.window_unit.current(0)

        self.sec_text = tk.StringVar()
        self.sec_text.set("Distribution list file")
        self.secondary_lab = tk.Label(self.window, textvariable = self.sec_text, fg = "black").grid(row = 3, column = 0, sticky = "E", pady = 5)        

        # Distribution list file
        self.window_distlist = tk.Button(self.window, text = "Select...", fg = "black", command = self.select_file)
        self.window_distlist.grid(row = 3, column = 1, pady = 5, padx = 5, sticky = "W")
        self.datafile = tk.Label(self.window, text = "", fg = "black")
        self.datafile.grid(row = 4, column = 1, pady = 5, sticky = "W")

        # Group sets
        self.window_groupset = ttk.Combobox(self.window, width = 60, state="readonly")
        self.group_sets(self)
        self.window_unit.bind("<<ComboboxSelected>>", self.group_sets)
        #self.groupset.grid(row = 3, column = 1, sticky = "W", padx = 5)

        tk.Label(self.window, text = "Send to (via):", fg = "black").grid(row = 0, column = 0, sticky = "E", pady = 5)
        self.window_sendto = ttk.Combobox(self.window, width = 60, state="disabled")
        self.window_sendto.grid(row = 0, column = 1, sticky = "W", pady = 5, padx = 5)
        self.window_sendto["values"] = ["All students (Canvas)",
                                        "Students in groups (Canvas)",
                                        "Students not in groups (Canvas)",
                                        "Groups (Canvas)",
                                        "All students (distribution list file)",
                                        "Students who completed the survey (distribution list file)",
                                        "Students who did not complete the survey (distribution list file)"]
        self.change_message(self)
        self.window_sendto.bind("<<ComboboxSelected>>", self.change_message)
        
        tk.Label(self.window, text = "Message:", fg = "black").grid(row = 5, column = 0, columnspan = 2, pady = 5, padx = 5)

        tk.Label(self.window, text = "Subject:", fg = "black").grid(row = 6, column = 0, sticky = "E", pady = 5)
        self.subject_text = tk.StringVar()
        self.subject = tk.Entry(self.window, textvariable = self.subject_text)
        self.subject.grid(row = 6, column = 1, pady = 5, padx = 5, sticky = "WE")
        self.subject_text.set(subject)

        self.message = tk.Text(self.window, width = 60, height = 15, wrap = tk.WORD)
        self.message.grid(row = 7, column = 0, columnspan = 2, sticky = "W", pady = 5, padx = 5)
        self.message.bind("<KeyRelease>", self.check_status)
        self.message.insert(tk.END, message)
        
        self.button_field = tk.Button(self.window, text = "Fields...", fg = "black", command = self.field_picker)
        self.button_field.grid(row = 8, column = 0, columnspan = 2, sticky = "W", padx = 5)
        
        self.sendmessage = tk.Button(self.window, text = "Send Bulk Message", fg = "black", state = "disabled", command = self.MessageSend)
        self.sendmessage.grid(row = 8, column = 0, columnspan = 2, pady = 5)
       
        self.distlist = {}
        globals.session["DistList"] = ""

        self.save_template = tk.IntVar()
        self.st_check = ttk.Checkbutton(self.window, text = "Save message as template", variable = self.save_template, state = "normal")
        self.st_check.grid(row = 8, column = 0, columnspan = 2, sticky = "E")
        
        # Check presets
        if preset=="confirm":
            self.window_sendto.current(3)
        elif preset=="nogroup":
            self.window_sendto.current(2)
        elif preset=="invitation":
            self.window_sendto.current(4)
        elif preset=="reminder":
            self.window_sendto.current(6)
        else:
            self.window_sendto.current(0)
            self.window_sendto["state"] = "readonly"
            self.st_check["state"] = "disabled"

        self.change_message(self)
        self.check_status(self)
    
    def check_focus(self, event_type = ""):
        self.master.grab_set()
    
    def distlist_keys(self, event_type = ""):
        keys = []
        if self.window_sendto.current() in [0,1,2]:
            keys = ["first_name", "last_name", "email"]
        if self.window_sendto.current() == 1:
            keys.append("group_name")
            keys.append("peers_list")
            keys.append("peers_bulletlist")
            keys.append("group_homepage")
        elif self.window_sendto.current() == 3:
            keys = ["firstnames_list", "members_list", "members_bulletlist", "name", "group_homepage"]        
        elif self.window_sendto.current() in [4,5,6]:
            self.create_distlist()
            if len(self.distlist) > 0:
                keys = list(self.distlist[next(iter(self.distlist))].keys())
        return(keys)
    
    def field_picker(self, event_type = ""):
        self.master.grab_release()
        field_window = tk.Toplevel()
        fp = FieldPicker(field_window, self.distlist_keys())
    
    def insert_field(self, event_type = "", field = ""):
        self.message.insert(tk.INSERT, field)
    
    def check_status(self, event_type = ""):
        if (len(self.message.get(1.0, tk.END).strip()) < 1) \
            or (self.window_sendto.current() in [4,5,6] and globals.session["DistList"] == "") \
            or (self.window_sendto.current() in [1,2,3] and len(self.group_sets_id) == 0):
            self.sendmessage["state"] = "disabled"
        else:
            self.sendmessage["state"] = "normal"
        if (self.window_sendto.current() in [4,5,6] and globals.session["DistList"] == ""):
            self.button_field["state"] = "disabled"
        else:
            self.button_field["state"] = "normal"            
        
    def select_file(self):
       globals.session["DistList"] = filedialog.askopenfilename(initialdir = ".", title = "Select file",filetypes = (("Compatible File Types",("*.csv")),("All Files","*.*")))
       self.datafile["text"] = globals.session["DistList"]
       self.check_status(self)
       
    def change_message(self, event_type = ""):
        self.distlist = {}
        if self.window_sendto.current() == 0: #Canvas all students
            self.sec_text.set("")
            self.window_distlist.grid_forget()
            self.window_groupset.grid_forget()
        elif self.window_sendto.current() in [1,2,3]:
            self.sec_text.set("Group set:")
            self.window_distlist.grid_forget()
            self.window_groupset.grid(row = 3, column = 1, sticky = "W", padx = 5)
        elif self.window_sendto.current() in [4,5,6]:
            self.sec_text.set("Distribution list file:")
            self.window_distlist.grid(row = 3, column = 1, pady = 5, padx = 5, sticky = "W")
            self.window_groupset.grid_forget()
            
    def group_sets(self, event_type = ""):
        self.group_sets_id = []
        self.group_sets_name = []
        globals.GQL.course_id = globals.session["course.ids"][self.window_unit.current()]
        for group_set in globals.GQL.group_sets():
            self.group_sets_id.append(group_set["_id"])
            self.group_sets_name.append(group_set["name"])
        if len(self.group_sets_id) > 0:
            self.window_groupset["values"] = tuple(self.group_sets_name)
            self.window_groupset.current(0)
        else:
            self.window_groupset["values"] = ()
            self.window_groupset.set('')
        globals.session["group_set"] = self.group_sets_id
        try: self.check_status(self)
        except: pass

    def create_distlist(self, event_type = ""):
        globals.GQL.course_id = globals.session["course.ids"][self.window_unit.current()] # Set the course ID to the one currently chosen

        # Log the action
        server.log_action("Send bulk message", globals.session["course.names"][self.window_unit.current()], globals.session["course.ids"][self.window_unit.current()])

        if self.window_sendto.current() in [1,2,3]: # Prepopulate the list of groups
            globals.GQL.groupset_id = self.group_sets_id[self.window_groupset.current()]        
        if self.window_sendto.current() in [0,1,2]: # Download students from Canvas
            students = globals.GQL.students_comprehensive()
        if self.window_sendto.current() == 0: # Simply copy the distribution list
            self.distlist = students
        if self.window_sendto.current() == 1: # Drop students without groups; get list of peers
            for student in students:
                if "group_id" in students[student]:
                    self.distlist[student] = students[student]
                    self.distlist[student]["peers_bulletlist"] = ""
                    self.distlist[student]["peers_list"] = ""
                    for n, peer in enumerate(self.distlist[student]["peers"], 1):
                        self.distlist[student]["peers_bulletlist"] += "* " + students[peer]["name"]
                        if (n < len(self.distlist[student]["peers"])):
                            self.distlist[student]["peers_bulletlist"] += "\n"
                            if (n > 1):
                                self.distlist[student]["peers_list"] += ", "
                        elif (n == len(self.distlist[student]["peers"])):
                            self.distlist[student]["peers_list"] += " and "
                        self.distlist[student]["peers_list"] += students[peer]["name"]
                    self.distlist[student]["group_homepage"] = globals.config["API_URL"] + "groups/" + str(students[student]["group_id"]) + "/"
        if self.window_sendto.current() == 2: # Drop students who have groups
            for student in students:
                if "group_id" not in students[student]:
                    self.distlist[student] = students[student]
        if self.window_sendto.current() == 3: # Get list of groups
            groups = globals.GQL.groups()
            students = globals.GQL.students_comprehensive()
            for group in groups:
                if len(groups[group]["users"]) > 0:
                    self.distlist[group] = groups[group]
                    self.distlist[group]["members_bulletlist"] = ""
                    self.distlist[group]["members_list"] = ""
                    self.distlist[group]["firstnames_list"] = ""
                    self.distlist[group]["group_homepage"] = globals.config["API_URL"] + "groups/" + str(group) + "/"
                    for n, member in enumerate(groups[group]["users"],1):
                        self.distlist[group]["members_bulletlist"] += "* " + students[member]["name"]
                        if (n < len(groups[group]["users"])):
                            self.distlist[group]["members_bulletlist"] += "\n"
                            if (n > 1):
                                self.distlist[group]["members_list"] += ", "
                                self.distlist[group]["firstnames_list"] += ", "
                        elif (n == len(groups[group]["users"])):
                            self.distlist[group]["members_list"] += " and "
                            self.distlist[group]["firstnames_list"] += " and "
                        self.distlist[group]["members_list"] += students[member]["name"]
                        self.distlist[group]["firstnames_list"] += students[member]["first_name"]
        if self.window_sendto.current() in [4,5,6]:
            print(globals.session["DistList"])
            if len(globals.session["DistList"]) > 0:
                try:
                    df = readTable(filename = globals.session["DistList"])
                    students = {}
                    globals.session["ID_column"] = None
                    if "Student_id" in df: globals.session["ID_column"] = "Student_id"
                    if "External Data Reference" in df: globals.session["ID_column"] = "External Data Reference"
                    if "ExternalDataReference" in df: globals.session["ID_column"] = "ExternalDataReference"
                    if "id" in df: globals.session["ID_column"] = "id"
                    if "ID" in df: globals.session["ID_column"] = "ID"
                    for column in df:
                        if "student_id" in column.lower():
                            globals.session["ID_column"] = column
                            break
                    if globals.session["ID_column"] is not None:
                        df = df_byRow(df, id_col = globals.session["ID_column"])
                        if df is None: 
                            self.distlist = None
                            return
                        for ID in df:
                            students[str(ID)] = {}
                            for col_name in df[ID]:
                                students[str(ID)][col_name] = df[ID][col_name]
                    else:
                        messagebox.showinfo("Error", "Cannot find user ID column in file. This column should be labelled 'id', 'Student_id' or 'External Data Reference'.")
                        self.master.destroy()
                except:
                    messagebox.showinfo("Error", "Trouble parsing file.")
                    self.master.destroy()
        if self.window_sendto.current() == 4:
            self.distlist = students
        if self.window_sendto.current() in [5,6] and "Status" not in df[list(df.keys())[0]]: # Look for "Status" in the first dictionary key
            messagebox.showinfo("Error", "Cannot find 'Status' column in file used to indicate whether a student has completed the survey.")
            self.master.destroy()
        if self.window_sendto.current() == 5:
            for student in students:
                if students[student]["Status"] == "Finished Survey":
                    self.distlist[student] = students[student]
        if self.window_sendto.current() == 6:
            for student in students:
                if students[student]["Status"] != "Finished Survey":
                    self.distlist[student] = students[student]
        
        # Delete all of the fields that are lists
        delete_dict = {}
        for case in self.distlist:
            for field in self.distlist[case]:
                if type(self.distlist[case][field]) in [list,tuple,dict]:
                    if case not in delete_dict: delete_dict[case] = []
                    delete_dict[case].append(field)
        for case in delete_dict:
            for field in delete_dict[case]:
                del self.distlist[case][field]
                
    def MessageSend(self):
        if self.preset!="" and self.save_template.get() == 1:
            print("Saving message as template...")
            globals.config["subject_" + self.preset] = self.subject_text.get()
            if self.preset == "invitation" or self.preset == "reminder":
                globals.config["message_" + self.preset][globals.config["SurveyPlatform"]] = self.message.get(1.0, tk.END)            
            else:
                globals.config["message_" + self.preset] = self.message.get(1.0, tk.END)
            save_config(globals.config)
            
        print("Preparing distribution list...")
        self.create_distlist(self)
        if self.distlist is None:
            self.master.destroy()
            return
        if len(self.distlist) == 0: 
            messagebox.showinfo("Error", "No students found.")
            self.master.destroy()
            return

        print("Preparing messages...")
        messagelist = {}
        for case in self.distlist:
            messagelist[case] = {}
            message = self.message.get(1.0, tk.END)
            subject = self.subject_text.get()
            for field in self.distlist[case]:
                subject = subject.replace("["+field+"]", str(self.distlist[case][field]))
                message = message.replace("["+field+"]", str(self.distlist[case][field]))
            messagelist[case]["subject"] = subject
            messagelist[case]["message"] = message
        
        for k, case in enumerate(messagelist, 1):
            print("Sending message {} of {}".format(k, len(messagelist)))
            if self.window_sendto.current() == 3: # Send group message
                bulk_message = globals.canvas.create_conversation(recipients = ['group_' + str(case)],
                                                     body = messagelist[case]["message"],
                                                     subject = messagelist[case]["subject"],
                                                     group_conversation = True,
                                                     context_code = "course_" + str(globals.GQL.course_id))
            else: # Send individual message
                try: # Try sending the course code information
                    bulk_message = globals.canvas.create_conversation(recipients = [str(case)],
                                                              body = messagelist[case]["message"],
                                                              subject = messagelist[case]["subject"],
                                                              force_new = True,
                                                              context_code = "course_" + str(globals.GQL.course_id))
                except: # Ignore the course code; don't send within context
                    try:
                        bulk_message = globals.canvas.create_conversation(recipients = [str(case)],
                                                                  body = messagelist[case]["message"],
                                                                  subject = messagelist[case]["subject"],
                                                                  force_new = True)
                    except: # Keep going even if the student has not been found, but report an error
                        print("Warning: Unable to send message to student ID {}.".format(case))
        messagebox.showinfo("Finished", "All messages have been sent.")
        self.master.destroy()


class FieldPicker:
    def __init__(self, master, fields = []):
        self.fields = fields
        self.master = master
        self.master.title("Fields")
        self.master.resizable(0,0)
        self.master.grab_set()
        
        self.window = tk.Frame(self.master)
        self.window.pack()
        
        self.field_list = tk.Listbox(self.window, width = 30, height = 20)
        self.field_list.pack(padx = 10, pady = 10)
        for field in fields:
            self.field_list.insert(tk.END, field)
            
        self.insert_field = tk.Button(self.window, text = "Insert Field", command = self.insert_text)
        self.insert_field.pack(padx = 10, pady = 10)
        
    def insert_text(self):
        self.master.grab_release()
        bm.insert_field(field = "[" + self.fields[self.field_list.curselection()[0]] + "]")
        self.master.destroy()
        

bm = None
def BulkMail_call(preset = ""):
    global bm
    window = tk.Toplevel()
    bm = BulkMail(window)
    
def BulkMail_confirm():
    global bm
    window = tk.Toplevel()
    bm = BulkMail(window,
                  preset = "confirm", 
                  subject = globals.config["subject_confirm"],
                  message = globals.config["message_confirm"])

def BulkMail_nogroup():
    global bm
    window = tk.Toplevel()
    bm = BulkMail(window,
                  preset = "nogroup", 
                  subject = globals.config["subject_nogroup"],
                  message = globals.config["message_nogroup"])
    
def BulkMail_invitation():
    global bm
    window = tk.Toplevel()
    bm = BulkMail(window,
                  preset = "invitation", 
                  subject = globals.config["subject_invitation"],
                  message = globals.config["message_invitation"][globals.config["SurveyPlatform"]])

def BulkMail_reminder():
    global bm
    window = tk.Toplevel()
    bm = BulkMail(window,
                  preset = "reminder", 
                  subject = globals.config["subject_reminder"],
                  message = globals.config["message_reminder"][globals.config["SurveyPlatform"]]) 