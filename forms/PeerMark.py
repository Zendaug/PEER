# -*- coding: utf-8 -*-
"""
Created on Fri Mar 11 18:01:37 2022

@author: tbednall
"""
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from tkcalendar import DateEntry
from datetime import datetime

from modules.functions import *
import modules.globals as globals
import forms.ToolTip as ToolTip
import forms.Policies as Policies
import modules.server as server
import modules.survey as survey_class
import modules.canvas as canvas

class PeerMark():
   def __init__(self, master):
       self.master = master
       self.master.title("Calculate Peer Marks")
       self.master.resizable(0, 0)
       
       self.peermark = tk.Frame(self.master)
       self.peermark.pack()
       self.validation = self.peermark.register(only_numbers)       

    
       self.UploadPeerMark = tk.BooleanVar()
       self.UploadPeerMark.set(True)
    
       tk.Label(self.peermark, text = "Unit:", fg = "black").grid(row = 0, column = 0, sticky = "E", pady = 5)
       self.peermark_unit = ttk.Combobox(self.peermark, width = 60, state="readonly")
       self.peermark_unit.bind("<<ComboboxSelected>>", self.group_assns)
       self.peermark_unit.grid(row = 0, column = 1, sticky = "W", padx = 5)
       ToolTip.CreateToolTip(self.peermark_unit, text =
                      'Select the unit that you want to upload the peer marks to.')      
    
       tk.Label(self.peermark, text = "Team assessment (shared mark):", fg = "black").grid(row = 1, column = 0, sticky = "E", pady = 5)
       self.peermark_assn = ttk.Combobox(self.peermark, width = 60, state="readonly")
       self.peermark_assn.grid(row = 1, column = 1, sticky = "W", padx = 5)
       ToolTip.CreateToolTip(self.peermark_assn, text =
                      'If you are using the moderated scoring method, select the team\n'+
                      'assignment used to calculate the student\'s initial peer mark.')
    
       tk.Label(self.peermark, text = globals.config["SurveyPlatform"] + " file:", fg = "black").grid(row = 2, column = 0, sticky = "E", pady = 5)
       self.peermark_peerdata = tk.Button(self.peermark, text = "Select...", fg = "black", command = self.select_file)
       self.peermark_peerdata.grid(row = 2, column = 1, padx = 5, pady = 5, sticky = "W")
       self.datafile = tk.Label(self.peermark, text = "", fg = "black")
       self.datafile.grid(row = 3, column = 1, pady = 5, sticky = "W")
       ToolTip.CreateToolTip(self.peermark_peerdata, text =
                      'Select a file containing the survey response data.')
    
       tk.Label(self.peermark, text = "Points:", fg = "black").grid(row = 4, column = 0, sticky = "E", pady = 5)
       self.peermark_points = tk.Entry(self.peermark, width = 5, validate="key", validatecommand=(self.validation, '%S'))
       self.peermark_points.grid(row = 4, column = 1, padx = 5, pady = 5, sticky = "W")
       self.peermark_points.insert(0, str(globals.config["RescaleTo"]))
       self.peermark_points.bind("<Key>", self.check_status)
       ToolTip.CreateToolTip(self.peermark_points, text =
                       'The number of points the Peer Mark is worth.\n' +
                       'If you adjust this setting, the Peer Mark will be rescaled.')       
    
       tk.Label(self.peermark, text = "Peer evaluation due date:", fg = "black").grid(row = 5, column = 0, sticky = "E")
       self.peermark_duedate = DateEntry(self.peermark, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern = "yyyy-mm-dd")
       self.peermark_duedate.grid(row = 5, column = 1, padx = 5, sticky = "W")
       ToolTip.CreateToolTip(self.peermark_duedate, text =
                      'The due date for the Peer Evaluation survey. Late penalties\n' +
                      'can be applied to students submitting after this date.')

       tk.Label(self.peermark, text = "Scoring policies:", fg = "black").grid(row = 6, column = 0, sticky = "E")    
       self.peermark_customise = tk.Button(self.peermark, text = "Review/modify", fg = "black", command = Policies.Policies_call)
       self.peermark_customise.grid(row = 6, column = 1, padx = 5, pady = 5, sticky = "W")
       ToolTip.CreateToolTip(self.peermark_customise, text =
                      'Review/modify the policies for scoring the peer mark.')
    
       self.peermark_calculate = tk.Button(self.peermark, text = "Calculate marks", fg = "black", command = self.start_calculate, state = 'disabled')
       self.peermark_calculate.grid(row = 7, column = 0, padx = 5, columnspan = 2, pady = 5)
       ToolTip.CreateToolTip(self.peermark_calculate, text =
                      'Calculate the peer marks and upload them to Canvas, along\n' +
                      'with peer feedback.')
       
       self.peermark_unit["values"] = tuple(globals.session["course.names"])
       self.peermark_unit.current(0)
       self.group_assns()
        
       self.status_text = tk.StringVar()
       self.statusbar = tk.Label(self.peermark, bd=1, relief=tk.SUNKEN, anchor=tk.W, textvariable = self.status_text)
       self.statusbar.grid(row = 8, columnspan = 2, sticky = "ew")       
       
       self.master.bind("<Enter>", self.check_status)
       
       if globals.config["firsttime_upload"] == True:
           messagebox.showinfo("Reminder", "Before uploading the peer evaluation marks to Canvas, it is recommended that you first review the scoring policies. Click on the \"Review/modify\" button to review these settings.")
           globals.config["firsttime_upload"] = False
           save_config(globals.config)

   def check_status(self, event_type):
       self.master.grab_set()
       
       # Check whether the user should have the option to select the assignment
       if globals.config["ScoreType"] == 2 or globals.config["MinimumPeersPolicy"] == 2:
           self.peermark_assn["state"] = "readonly"
       else:
           self.peermark_assn["state"] = "disabled"

       # Check whether the Calculate Peer Marks button should be enabled
       if len(self.datafile["text"]) > 0 and not ((globals.config["ScoreType"]==2 or globals.config["MinimumPeersPolicy"] == 2) and len(self.assn_id)==0) and not (self.peermark_points.get() == "" or int(self.peermark_points.get()) <= 0):
           self.peermark_calculate["state"] = 'normal'
       else:
           self.peermark_calculate["state"] = 'disabled'

   def select_file(self):
       globals.session["DataFile"] = filedialog.askopenfilename(initialdir = ".",title = "Select file",filetypes = (("Compatible File Types",("*.csv","*.tsv","*.xlsx")),("All Files","*.*")))
       self.datafile["text"] = globals.session["DataFile"]
       self.check_status
       
   def group_assns(self, event_type = ""):
        self.assn_id = []
        self.assn_name = []
        
        globals.GQL.course_id = globals.session["course.ids"][self.peermark_unit.current()]
        assns = globals.GQL.group_assignments()
        for assn in assns:
            self.assn_id.append(assn)
            self.assn_name.append(assns[assn]["name"])
        
        self.peermark_assn["values"] = tuple(self.assn_name)
        try: # Some courses do not have assignments (bizarrely enough!)
            self.peermark_assn.current(0)
        except:
            pass
    
   def status(self, stat_text):
       self.status_text.set(stat_text)
       
   def start_calculate(self):
       self.status("Loading, please be patient...")
       if globals.config["RescaleTo"] != int(self.peermark_points.get()):
           globals.config["RescaleTo"] = int(self.peermark_points.get())
           save_config(globals.config)
       if int(self.peermark_points.get()) <= 0:
           messagebox.showinfo("Error", "The Peer Mark must be worth more than 0 points.")
           return
       globals.session["DueDate"] = self.peermark_duedate.get_date()
       if (globals.config["ScoreType"]==2 or globals.config["MinimumPeersPolicy"]==2):
           globals.session["GroupAssignment_ID"] = self.assn_id[self.peermark_assn.current()]
       globals.GQL.course_id = globals.session["course.ids"][self.peermark_unit.current()]
       globals.GQL.groupset_id = 0
       globals.session["course.name"] = globals.session["course.names"][self.peermark_unit.current()]
       globals.session["course.id"] = globals.session["course.ids"][self.peermark_unit.current()]
       course = globals.session["course.ids"][self.peermark_unit.current()]
       calculate_marks(course, self)

def PeerMark_call():
    globals.pm = tk.Toplevel()
    window = PeerMark(globals.pm)
    

#%% Calculate marks   
def calculate_marks(course, pm):
   if globals.config["ExportData"]:
       filename = ""
       filename = tk.filedialog.asksaveasfilename(initialfile=globals.session["course.name"] + " (Peer Data Export)", defaultextension=".xlsx", title = "Save Exported Data As...", filetypes = (("Excel files","*.xlsx"),("all files","*.*")))

   server.log_action("Calculate/upload peer marks", globals.session["course.name"], globals.session["course.id"])
   print("\nCalculating peer marks... please be patient.")
   
   # Set up data frame
   canvas.data = globals.GQL.students_comprehensive()
   
   # Load the Qualtrics file
   survey = survey_class.survey(globals.session["DataFile"], globals.config["SurveyPlatform"])
   if survey.error: return
   
   # Set up the assignment
   new_assn = new_assn = 'courseId: {}, name: "{}", pointsPossible: {}, description: "This assignment is your Peer Mark for the team project.", state: published, omitFromFinalGrade: {}, dueAt: "{}"'.format(course, globals.config['PeerMark_Name'], globals.config['RescaleTo'], str(bool(globals.config["feedback_only"])).lower(), datetime.strftime(globals.session["DueDate"], '%Y-%m-%dT23:59:59'))
      
   # Get assignment group (if it exists):
   if (globals.config["ScoreType"]==2 or globals.config["MinimumPeersPolicy"]==2):
       assignment = globals.GQL.group_assignments()[globals.session["GroupAssignment_ID"]]
       #new_assn["assignment_group_id"] = assignment["assignmentGroup"]["_id"]
       new_assn = new_assn + ', assignmentGroupId: "{}"'.format(assignment["assignmentGroup"]["_id"])
         
   # Create new assignment
   if globals.config["UploadData"]:
       print("\nCreating new assignment for Peer Mark")
       try:
           new_assignment = globals.GQL.query("mutation MyMutation {__typename createAssignment(input: {" + new_assn + "})  {assignment {_id}}}")["data"]["createAssignment"]["assignment"]["_id"]
       except:
           messagebox.showinfo("Error", "Unable to create assignment. You may not be authorised to perform this action in {}.".format(globals.session["course.name"]))
           return
       print("New assignment created: " + globals.config['PeerMark_Name'] + " (" + str(new_assignment) + ")")
            
       globals.GQL.query('mutation MyMutation {\nsetAssignmentPostPolicy(input: {assignmentId: ' + str(new_assignment) + ', postManually: true}) {\npostPolicy {\npostManually\n}\n}}')

   # Get assignment details   
   if (globals.config["ScoreType"]==2 or globals.config["MinimumPeersPolicy"]==2):
       # Download student details and grades and submission score
       assignment = globals.GQL.group_assignments()[globals.session["GroupAssignment_ID"]]
       print("\nDownloading group marks for " + assignment["name"])
       globals.session["assn_points_possible"] = assignment["pointsPossible"]
       #assignment = course.get_assignment(globals.session["GroupAssignment_ID"])
       
       submission = 'query MyQuery {__typename assignment(id: "' + str(globals.session["GroupAssignment_ID"]) + '") {submissionsConnection {nodes {score user {_id}}}}}'
       submissions = globals.GQL.query(submission)["data"]["assignment"]["submissionsConnection"]["nodes"]
       
       #submissions = assignment.get_submissions()
       for submission in submissions:
           user = submission["user"]["_id"]
           if str(user) in canvas.data: # Leave out the Test Student
               print(" > {} ({}): {}".format(canvas.data[str(user)]["name"], user, submission["score"]))
               canvas.data[str(user)]["orig_grade"] = submission["score"]
   else:
       # Use this method if we are only calculating totals. Download all users, but not submission score.
       print("\nDownloading list of students...")
       for user in canvas.data:
           canvas.data[str(user)]["orig_grade"] = None
    
   # Upload marks to Canvas
   if globals.config["UploadData"]:
       print("\nUploading marks to Canvas...")
       course = globals.canvas.get_course(course)
       iter = 0
       tstamp = datetime.now().timestamp()
       globals.session["date_error"] = False
       while (iter==0 and (datetime.now().timestamp() - tstamp < 10)): # Keep looping until at least one iteration has been completed. Timeout after 10 seconds.
           #sub_query = 'query MyQuery {__typename assignment(id: "' + str(new_assignment) + '") {submissionsConnection {nodes {score user {_id}}}}}'
           assignment = course.get_assignment(new_assignment)
           submissions = assignment.get_submissions()
           #submissions = globals.GQL.query(sub_query)["data"]["assignment"]["submissionsConnection"]["nodes"]
           for submission in submissions:
               iter = iter + 1
               ID = submission.user_id #submission["user"]["_id"]
               adj = survey.adj_score(ID)
               try:
                   print(" > {}: {} / {}".format(canvas.data[str(ID)]["name"],
                                                 round(adj["score"],2),
                                                 globals.config["RescaleTo"]))
               except:
                   print( "No data available for: ID: {}, Name: {}".format(ID,
                                                           canvas.data[str(ID)]["name"] if str(ID) in canvas.data else "User not found"))
#                   pass

               sub_date = survey.submission_date(ID)
               sub_dict = {"posted_at": "null"}
 
               if adj["status"] == "normal" or adj["status"] == "original_score":
                   sub_dict["posted_grade"] = adj["score"]*adj["penalty"]
                   if (sub_date is not None and ID in survey.data):
                       sub_dict["submitted_at"] = datetime.strftime(sub_date, '%Y-%m-%dT%H:%M:%S')
                   if adj["seconds_late"] > 0 and globals.config["Penalty_Late_Custom"] == False:
                       sub_dict["late_policy_status"] = "late"
                       sub_dict["seconds_late_override"] = adj["seconds_late"]
               elif adj["status"] == "excused":
                   if adj["penalty"]==0: # If they have lost all of their marks because of penalties, give them a score of zero rather than excusing them
                       sub_dict["posted_grade"] = 0
                   else:
                       sub_dict["excuse"] = True
               elif adj["status"] == "zero":
                   sub_dict["posted_grade"] = 0
               elif adj["status"] == "error":
                   sub_dict["late_policy_status"] = "missing"
                
               submission.edit(submission=sub_dict)
               submission.edit(comment={'text_comment': survey.comments(ID)})
               
       if globals.session["date_error"]:
           print("\nWarning: Could not parse the 'EndDate' field. This error sometimes occurs when the Qualtrics file has been modified in Excel.")

       globals.GQL.query('mutation MyMutation {__typename updateAssignment(input: {id: ' + str(new_assignment) + ', state: ' + ("published" if globals.config["publish_assignment"] else 'unpublished') + 'unpublished})}')
       
       if iter==0: 
           messagebox.showinfo("Error", "Unable to upload marks to Canvas. Please try again.")
           pm.master.destroy()

    # Save the data; add all of the penalties information
   if globals.config["ExportData"] and filename != "":
        print("\nSaving data to Excel...")
        globals.session["sections"] = globals.GQL.sections()
        print(" > Rater/receiver data (long format)...")
        long_data = {}
        long_data["RaterName"] = []
        long_data["ReceiverName"] = []
        long_data["GroupName"] = []
        long_data["RaterID"] = []
        long_data["ReceiverID"] =  []
        long_data["GroupID"] = []
        for item in survey.tw_list: long_data[item] = []
        long_data["FeedbackShared"] = []
        long_data["FeedbackInstructor"] = []
        if (globals.config["ScoreType"] == 2):
            long_data["OriginalScore"] = []
            long_data["RatioAdjust"] = []
        long_data["PeerMark"] = []
        long_data["NonComplete"] = []
        for ID2 in canvas.data: # The list of students from Canvas
            ID = cleanupNumber(ID2)
            if ID in survey.data:
                long_data["NonComplete"].extend([0] * len(survey.rater_peerID(ID)))
            else:
                long_data["NonComplete"].extend([1] * len(survey.rater_peerID(ID)))
            for a, rater in enumerate(survey.rater_peerID(ID)):
                long_data["ReceiverName"].append(canvas.data[str(ID)]["name"])
                long_data["ReceiverID"].append(ID)
                if str(rater) in canvas.data: # Check that the person's peer is in the Canvas data...
                    long_data["RaterName"].append(canvas.data[str(rater)]["name"])
                elif rater in survey.data: # Otherwise, get the person's name from the Qualtrics data...
                    long_data["RaterName"].append(survey.data[rater]["Student_name"])
                else: # Otherwise, just leave blank
                    long_data["RaterName"].append(None)
                long_data["RaterID"].append(rater)
                if ID in survey.data: # Check that the person is in the Qualtrics data...
                    long_data["GroupName"].append(survey.data[ID]["Group_name"])
                    long_data["GroupID"].append(survey.data[ID]["Group_id"])
                else:
                    long_data["GroupName"].append(None)
                    long_data["GroupID"].append(None)                     
                for b, item in enumerate(survey.tw_list):
                    long_data[item].append(survey.ratings_received(ID)[b][a])
                if a > 0: # Provide the peer feedback
                    try:
                        long_data["FeedbackShared"].append(survey.feedback(ID)[0][a-1])
                    except:
                        long_data["FeedbackShared"].append("")
                    try:
                        long_data["FeedbackInstructor"].append(survey.feedback(ID)[1][a-1])
                    except:
                        long_data["FeedbackInstructor"].append("")
                else: # Provide the person's general feedback
                    if ID in survey.data:
                        try:
                            long_data["FeedbackShared"].append(survey.data[ID]["Feedback1"][-1])
                        except:
                            long_data["FeedbackShared"].append("")
                        try:
                            long_data["FeedbackInstructor"].append(survey.data[ID]["Feedback2"][-1])
                        except:
                            long_data["FeedbackInstructor"].append("")
                    else:
                        long_data["FeedbackShared"].append("")
                        long_data["FeedbackInstructor"].append("")                       
                if (globals.config["ScoreType"] == 2):
                    if "orig_grade" in canvas.data[ID2]:
                        long_data["OriginalScore"].append(cleanupNumber(canvas.data[ID2]["orig_grade"]))
                    else:
                        long_data["OriginalScore"].append(None)
                    long_data["RatioAdjust"].append(survey.ratee_ratio(ID)[a])
            long_data["PeerMark"].extend([survey.adj_score(ID)["score"]] * len(survey.rater_peerID(ID)))

    # Create dataset (wide form) if requested
        print(" > Student data (wide format)...")
        wide_data = {}
        wide_data["StudentID"] = []
        wide_data["StudentName"] = []
        for section in globals.session["sections"]: wide_data[globals.session["sections"][section]["name"]] = []
        if (globals.config["ScoreType"] == 2):
            wide_data["OriginalScore"] = []
        wide_data["PeerMark"] = []
        wide_data["PartialComplete"] = []
        wide_data["SelfPerfect"] = []
        wide_data["PeersAllZero"] = []
        wide_data["NonComplete"] = []
        for item in survey.other_list: wide_data[item] = []

        for ID2 in canvas.data:
            ID = int(ID2)
            if ID in survey.data:
                wide_data["NonComplete"].append(0)
                for a, item in enumerate(survey.other_list):
                    wide_data[item].append(survey.data[ID][item])
                wide_data["PartialComplete"].append(survey.PartialComplete(ID))
                wide_data["SelfPerfect"].append(survey.SelfPerfect(ID))
                wide_data["PeersAllZero"].append(survey.PeersAllZero(ID))
            else:
                wide_data["NonComplete"].append(1)
                wide_data["PartialComplete"].append(None)
                wide_data["SelfPerfect"].append(None)
                wide_data["PeersAllZero"].append(None)
                for a, item in enumerate(survey.other_list):
                    wide_data[item].append(None)
            for section in globals.session["sections"]: wide_data[globals.session["sections"][section]["name"]].append(str(section) in canvas.data[ID2]["sections"])
            wide_data["StudentName"].append(canvas.data[ID2]["name"])
            wide_data["StudentID"].append(ID)
            if (globals.config["ScoreType"] == 2):
                if "orig_grade" in canvas.data[ID2]:
                    wide_data["OriginalScore"].append(cleanupNumber(canvas.data[ID2]["orig_grade"]))
                else:
                    wide_data["OriginalScore"].append(None)
            wide_data["PeerMark"].append(survey.adj_score(ID)["score"])

   if globals.config["ExportData"] and globals.config["UploadData"]:
        messagebox.showinfo("Upload Complete", 'The peer marks have been successfully uploaded to Canvas but are hidden from students. When they are ready to be released to students, you should un-hide them in the Gradebook. Peer rater data saved to "' + filename + '".')
   elif globals.config["UploadData"]:
        messagebox.showinfo("Upload Complete", 'The peer marks have been successfully uploaded to Canvas but are hidden from students. When they are ready to be released to students, you should un-hide them in the Gradebook.')
   elif globals.config["ExportData"]:
        messagebox.showinfo("Export Complete", 'Peer rater data saved to "' + filename + '".')
       
   writeXL(filename = filename, datasets=[long_data, wide_data], sheetnames=["Rater Data", "Student Data"])
   pm.master.destroy()