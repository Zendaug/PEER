# -*- coding: utf-8 -*-
"""
Created on Thu Mar 17 08:22:16 2022

@author: tbednall
"""

from datetime import datetime, timedelta
from random import shuffle
from tkinter import messagebox
import re

from modules.functions import *
import modules.canvas as canvas
import modules.globals as globals

class survey():
    def __init__(self, file = "    ", platform = "Qualtrics"):
        self.file = file
        self.data = []
        self.orig_data = []      
        self.colnames = []
        self.tw_list = []
        self.tw_names = []
        self.tw_long = []
        self.other_list = []
        self.error = False
               
        # Load the survey file
        print("Loading the survey data file...")
        try:
            if file[-3:].upper()=="CSV":
                self.data = readTable(file)
            elif file[-3:].upper() in ["TSV","TXT"]:
                self.data = readTable(file, delimiter="\t")
            elif file[-3:].upper() == "XLS" or file[-3:].upper() in ["XLSX", "XLSM"]:
                self.data = readXL(file) #load_workbook(globals.session["DataFile"])
            else:
                messagebox.showinfo("Error", "Unable to import data from this file type.")
                self.error = True
                return
        except:
            messagebox.showinfo("Error", "Unable to parse data from this file.")
            self.error = True
            return
        
        self.orig_data = self.data
        
        if platform=="Qualtrics": self.cleanupQualtrics()
        if platform=="LimeSurvey": self.cleanupLimeSurvey()
        
    def cleanupQualtrics(self):
        print("\nCleaing up Qualtrics file...")
        rep_list = ["Please rate each team member in relation to: ", " - Myself"]
        for a in range(1,10): rep_list.append(" - [Field-Peer" + str(a) + "_name]")
        
        # Determine the teamwork items in the survey data file
        for column in self.data:
            if (column[0:8].upper() == "TEAMWORK"):
                t_name = self.data[column][0]
                t_name.replace("&shy", "\u00AD")
                if (column.find("_") != -1):
                    self.tw_list.append(column[0:column.rfind("_")])
                    if t_name.count("\u00AD") == 2:
                        str_loc = find_all(t_name, "\u00AD")
                        self.tw_names.append(t_name[str_loc[0]+1:str_loc[1]])
                    elif t_name.count("*") == 2:
                        str_loc = find_all(t_name, "*")
                        self.tw_names.append(t_name[str_loc[0]+1:str_loc[1]])
                    else:
                        self.tw_names.append(replace_multiple(t_name, rep_list, ""))           
            elif (column[0:8].upper() != "FEEDBACK") and replace_multiple(column.upper(), range(0,10)) not in ["PEER_NAME", "PEER_ID"] and column not in ["Student_name", "Student_id", "ExternalReference"]:
                self.other_list.append(column)
        self.tw_list = unique(self.tw_list)
        self.tw_names = unique(self.tw_names)
        
        if len(self.tw_list)==0:
            messagebox.showinfo("Error", "Cannot find the peer ratings in the survey data file. These columns should be labelled TeamWork1_1, TeamWork2_1, etc.")
            self.error = True
            return
        
        # Drop redundant rows
        first_num = first_numeric(self.data["ExternalReference"])
        for column in self.data:
            del self.data[column][0:first_num] # Drop the header rows
        
        # Convert the "Finished" column to Boolean values
        for b, case in enumerate(self.data["Finished"]):
            self.data["Finished"][b] = (case=="True")
        
        # It may be possible to remove this from the Qualtrics function...
        # Detect duplicate rows
        for b, case in enumerate(self.data["ExternalReference"]):
            if self.data["ExternalReference"].count(case) > 1:
                messagebox.showinfo("Error", "Duplicate rows detected. Problem with student ID" + str(case) + ". Please remove duplicate rows and try again.")
                self.error = True
                return
            else:
                self.data["ExternalReference"][b] = cleanupNumber(case) # Convert to an integer
        
        # Convert to a "records" data frame
        self.data = df_byRow(self.data, "ExternalReference")
         
        # Collapse the data about peers into a list   
        collapseList(self.data, "Peers", "Peer*_id")
        collapseList(self.data, "Peer_names", "Peer*_name")
         
        # Collapse the rating data regarding the teamwork items into a list
        for tw in self.tw_list: collapseList(self.data, tw, tw+"_*")
        collapseList(self.data, "TeamWork", self.tw_list)
        
        # Collapse the feedback data
        collapseList(self.data, "Feedback1", "Feedback#1_*_1", delete_blank = False)
        collapseList(self.data, "Feedback2", "Feedback#2_*_1", delete_blank = False)
            
        # Eliminate response if no peer ratings were given at all
        delete_list = []
        for student in self.data:
            if len(self.data[student]["TeamWork"][0]) == 0:
                delete_list.append(student)
        for student in delete_list:
            del self.data[student]
               
        # Replace strings with numbers for the TeamWork items; set zero as the bottom score rather than 1
        for student in self.data:
            for y, item in enumerate(self.data[student]["TeamWork"]):
                for x, rating in enumerate(self.data[student]["TeamWork"][y]):
                    if type(rating) == str: self.data[student]["TeamWork"][y][x] = number_beginning(rating) 
                    self.data[student]["TeamWork"][y][x] -= 1 # Set zero as the bottom score
    
    def cleanupLimeSurvey(self):
        print("\nCleaing up LimeSurvey file...")
        
        # Check to see whether the dataset has been downloaded in the correct format
        if "id. Response ID" not in list(self.data.keys())[0]:
            messagebox.showinfo("Error", "The LimeSurvey data file does not appear to be in the correct format. Please export the questions in the \"Question code & question text\" format")
            self.error = True
            return

        # Change the name of the "EndDate" column
        self.data["EndDate"] = self.data.pop("submitdate. Date submitted")

        # Check that attributes have been included
        found = False
        for column in self.data:
            if "Student_id" in column and "attribute_" in column:
                found = True
                self.data["ExternalReference"] = self.data.pop(column) # Rename as "ExternalReference"
                break
        if found == False:
            messagebox.showinfo("Error", "The LimeSurvey data file does not contain the team membership data. Please select all of the participant fields from the \"Participant control\" box when exporting the survey data.")
            self.error = True
            return
        
        # Identify all of the TeamWork items
        for column in self.data:
            self.colnames.append(column)
            if (column[0:8].upper() == "TEAMWORK"):
                column = column.replace("&shy", "\u00AD")
                if (column.find("[") != -1):
                    self.tw_list.append(column[:column.find("[")])
                    self.tw_long.append(column)
                    if column.count("\u00AD") == 2:
                        str_loc = find_all(column, "\u00AD")
                        self.tw_names.append(column[str_loc[0]+1:str_loc[1]])
                    elif column.count("*") == 2:
                        str_loc = find_all(column, "*")
                        self.tw_names.append(column[str_loc[0]+1:str_loc[1]])
                    else: # Eliminate all of the other information to get the variable name...
                        tw_name = re.sub("TeamWork..?\[..?\]\. Please rate each team member in relation to: ", "", column)
                        tw_name = re.sub("\[Myself\]", "", tw_name)
                        tw_name = re.sub("\[{TOKEN:ATTRIBUTE_..?}\]", "", tw_name)
                        tw_name = tw_name.replace("\xad", "")
                        tw_name = clean_text(tw_name)
                        self.tw_names.append(tw_name)     
            elif column[0:8].upper() != "FEEDBACK" and "attribute_" not in column and column != "ExternalReference":
                self.other_list.append(column)
        self.tw_list = unique(self.tw_list)
        self.tw_names = unique(self.tw_names)
        self.tw_long = unique(self.tw_long)

        # Check if any teamwork items were found
        if len(self.tw_list)==0:
            messagebox.showinfo("Error", "Cannot find the peer ratings in the survey data file. These columns should be labelled TeamWork1_1, TeamWork2_1, etc.")
            self.error = True
            return
        
        # Create a "Finished" column, based on whether there is a submission date
        self.data["Finished"] = []
        for b, case in enumerate(self.data["EndDate"]):
            if case is not None and case != "":
                self.data["Finished"].append(True)
            else:
                self.data["Finished"].append(False)
        
        # Remove respondents with a blank student ID
        max_len = len(self.data["ExternalReference"])
        b = 0
        while b < max_len:            
            if self.data["ExternalReference"][b] is None or self.data["ExternalReference"][b] == "":
                for column in self.data:
                    del self.data[column][b]
                max_len = max_len - 1
            else:
                b = b + 1
        
        # Detect duplicate rows
        for b, case in enumerate(self.data["ExternalReference"]):
            if self.data["ExternalReference"].count(case) > 1:
                messagebox.showinfo("Error", "Duplicate rows detected. Problem with student ID" + str(case) + ". Please remove duplicate rows and try again.")
                self.error = True
                return
            else:
                self.data["ExternalReference"][b] = cleanupNumber(case) # Convert to an integer
        
        # Convert to a "records" data frame
        self.data = df_byRow(self.data, "ExternalReference")
        
        # Collapse the data about peers into a list   
        collapseList(self.data, "Group_id", [b for b in self.colnames if re.search("Group_id", b)], unlist = True)
        collapseList(self.data, "Group_name", [b for b in self.colnames if re.search("Group_name", b)], unlist = True)
        
        # Collapse the data about peers into a list   
        collapseList(self.data, "Peers", [b for b in self.colnames if re.search("Peer._id", b)])
        collapseList(self.data, "Peer_names", [b for b in self.colnames if re.search("Peer._name", b)])
         
        # Collapse the rating data regarding the teamwork items into a list
        for tw in self.tw_list:
            collapseList(self.data, tw, [b for b in self.tw_long if tw in b])
        collapseList(self.data, "TeamWork", self.tw_list)
        
        # Collapse the feedback data
        collapseList(self.data, "Feedback1", [b for b in self.colnames if re.search("Feedback\[..?_1", b)], delete_blank = False)
        collapseList(self.data, "Feedback2", [b for b in self.colnames if re.search("Feedback\[..?_2", b)], delete_blank = False)    

        # Eliminate response if no peer ratings were given at all
        delete_list = []
        for student in self.data:
            if len(self.data[student]["TeamWork"][0]) == 0:
                delete_list.append(student)
        for student in delete_list:
            del self.data[student]
               
        # Replace strings with numbers for the TeamWork items; set zero as the bottom score rather than 1
        for student in self.data:
            for y, item in enumerate(self.data[student]["TeamWork"]):
                for x, rating in enumerate(self.data[student]["TeamWork"][y]):
                    if type(rating) == str: self.data[student]["TeamWork"][y][x] = number_beginning(rating) 
                    self.data[student]["TeamWork"][y][x] -= 1 # Set zero as the bottom score
    
    def submission_date(self, ID):
        '''This function returns the date the Qualtrics survey was submitted (the EndDate field).'''
        if ID not in self.data: return(None)
        if "EndDate" in self.data[ID]:
            try:
                return(datetime.strptime(self.data[ID]["EndDate"], '%Y-%m-%d %H:%M:%S'))
            except:
                globals.session["date_error"] = True
                return(None)
        else:
            return(None)
    
    def days_late(self, ID, weekdays_only = True, integer = True):
        '''This function returns the number of days late an assignment was (or 0 if it was submitted before time). By default, it only counts weekdays.'''
        sub_date = self.submission_date(ID)
        if "DueDate" in globals.session and type(sub_date) is datetime:
            due_date = globals.session["DueDate"]
            if integer:
                sub_date = datetime(sub_date.year, sub_date.month, sub_date.day, 0,0,0)
                due_date = datetime(due_date.year, due_date.month, due_date.day, 0,0,0)
            else:
                sub_date = datetime(sub_date.year, sub_date.month, sub_date.day, sub_date.hour, sub_date.minute, sub_date.second)
                due_date = datetime(due_date.year, due_date.month, due_date.day, 23,59,59)
            d_late = max(0, (sub_date - due_date).days + (sub_date - due_date).seconds / 86400) # Divide by the number of seconds per day to get a decimal
            for a in range(0,int(d_late)*weekdays_only):
                d_late -= ((sub_date - timedelta(days=a)).weekday() >= 5)
            return(cleanupNumber(d_late))
        else:
            return(None)
    
    def ratings_given(self, ID):
        '''This function returns all of the ratings given by a rater (ID) in a list, including self-ratings'''
        if ID not in self.data: return([None])
        return(self.data[ID]["TeamWork"])
    
    def rater_total(self, ID):
        '''A function that calculates the total number of points allocated by a rater to each peer'''
        if ID not in self.data: return([None])
        tot = [0] * (len(self.data[ID]["Peers"]) + 1) # Including self-rating
        for item in self.ratings_given(ID):
            for x, rating in enumerate(item):
                try:
                    tot[x] = tot[x] + rating
                except:
                    pass
        return(tot)
    
    def rater_peerID(self, ID):
        ''' Returns the IDs of the peer raters of a person in a list using 2 methods'''
        tot = []
        if ID in self.data: # Method 1 - look up the person's record
            for peer in self.data[ID]["Peers"]: tot = tot + [int(peer)]
        else: # Method 2 - Look through all of their peers and grab IDs  
            for user in self.data:
                if "Peers" in self.data[user]:
                    if ID in self.data[user]["Peers"]:
                        tot.append(int(user))
        tot.insert(0, int(ID))
        return(tot)
    
    def NonComplete(self, ID):
        '''A function to detect whether a person has not completed the peer survey'''
        # Remember, we remove all students who have not filled out any of the TeamWork items
        return(not(ID in self.data))
    
    def PartialComplete(self, ID):
        '''A function to detect whether a person has only partially completed the peer survey'''
        n_peers = len(self.rater_peerID(ID))
        for item in self.data[ID]["TeamWork"]:
            if len(item) < n_peers: return(True)
        return(False)
    
    def SelfPerfect(self, ID):
        '''A function to detect whether a person has given themselves all top ratings'''
        if self.NonComplete(ID) == True: return(False)
        for item in self.ratings_given(ID):
            if len(item) > 0:
                if item[0] < globals.config["PointsPossible"] - 1: return(False)    
        return(True)
    
    def PeersAllZero(self, ID):
        '''A function to detect whether a person has given their peers all zeroes. It assumes that the minimum number of peers has been met.'''
        if len(self.data[ID]["Peers"]) < globals.config["MinimumPeers"]: return(False)
        if self.NonComplete(ID) == True: return(False)
        for item in self.ratings_given(ID):
            for rating in item[1:]:
                if rating > 0: return(False)
        return(True)
    
    def ratings_received(self, ID):
        '''A function that returns all of the ratings received by a person (ratings nested within items)'''
        rr = []
        for y in self.tw_list: rr.append([])
        for x, rater_ID in enumerate(self.rater_peerID(ID)):
            # Only include the ratings if the person: (1) is in the Qualtrics file, (2) is not excluded because of partial completes, perfect self scores, or rating peers as all zeroes
            peer_ratings = self.ratings_given(rater_ID)
            for y in range(len(rr)):
                if rater_ID in self.data and not(
                        (globals.config["Exclude_PartialComplete"] == 1 and self.PartialComplete(rater_ID) == True) or
                        (globals.config["Exclude_SelfPerfect"] == 1 and self.SelfPerfect(rater_ID) == True) or
                        (globals.config["Exclude_PeersAllZero"] == 1 and self.PeersAllZero(rater_ID) == True) or
                         len(peer_ratings[y])==0):
                    try: 
                        rr[y].append(peer_ratings[y][self.rater_peerID(rater_ID).index(ID)])
                    except: # If there is an irregularity in the dataset, simply add "None"
                        rr[y].append(None)
                else:
                    rr[y].append(None)
        return(rr)
        
    def ratee_total(self, ID):
        '''A function that calculates the total number of points received from each peer, including oneself.'''
        tot = []
        rr = self.ratings_received(ID)
        for x, rater in enumerate(self.rater_peerID(ID)):
            tot.append(None)
            for y, item in enumerate(rr):
                if item[x] is not None:
                    if tot[x] is None: tot[x] = 0
                    tot[x] += item[x]
        return(tot)
    
    def n_ratings(self, ID):
        '''A function that returns the number of valid ratings received from peers. It does not count the person's self-rating'''
        if len(self.rater_peerID(ID)) > 1:
            return(nansum([True if x is not None else False for x in self.ratee_total(ID)[1:]]))
        else:
            return(0)
    
    def ratee_ratio(self, ID):
        '''A function that calculates the ratio of points received from a peer'''
        ratio = []
        for x, rater in enumerate(self.rater_peerID(ID)):
            if rater in self.data and ID in self.rater_peerID(rater):
                    ratings = self.rater_total(rater)
                    ID_index = self.rater_peerID(rater).index(ID)
                    if globals.config["SelfVote"] not in [1,3]: ratings[0] = None # Only count the self-vote towards the ratio if self-voting allowed
                    nratings = nansum([True if x is not None else False for x in ratings])
                    if nansum(ratings) == 0: # Return NAN if the rater has given everyone a zero
                        ratio.append(None)
                    elif ID == rater: # Only calculate the ratio for one's own score if self-scoring allowed
                        if globals.config["SelfVote"] in [1,3]:
                            ratio.append(None)
                        else:
                            if ratings[ID_index] is None:
                                ratio.append(None)
                            else:
                                ratio.append(nratings * ratings[ID_index] / nansum(ratings))
                    else: # Otherwise, return the ratio
                        ratio.append(nratings * ratings[ID_index] / nansum(ratings))
            else: # Rater not found
                ratio.append(None)
        return(ratio)
    
    
    def feedback(self, ID):
        '''A function to locate the feedback provided to each person'''
        fb = [[], []]
        for rater_ID in self.rater_peerID(ID)[1:]:
            if rater_ID in self.data:
                id_index = self.data[rater_ID]["Peers"].index(ID)
                if len(self.data[rater_ID]["Feedback1"]) > 0:
                    fb[0].append(self.data[rater_ID]["Feedback1"][id_index])
                else:
                    fb[0].append(None)
                if len(self.data[rater_ID]["Feedback2"]):
                    fb[1].append(self.data[rater_ID]["Feedback2"][id_index])
                else:
                    fb[1].append(None)
            else:
                fb[0].append(None)
                fb[1].append(None)
        return(fb)
    
    def general_feedback(self, ID):
        '''A function that returns general feedback provided to the team'''
        gen_feed = [[], []]
        for rater_ID in self.rater_peerID(ID):
            if rater_ID in self.data:
                feed1 = ""
                feed2 = ""
                if len(self.data[rater_ID]["Feedback1"]) > 0: feed1 = self.data[rater_ID]["Feedback1"][-1]
                if len(self.data[rater_ID]["Feedback2"]) > 0: feed2 = self.data[rater_ID]["Feedback2"][-1]
                if len(feed1) > 0: gen_feed[0].append(feed1)
                if len(feed2) > 0: gen_feed[1].append(feed2)
        return(gen_feed)
    
    def penalty(self, ID):
        '''A function to calculate the % deducted as a penalty'''
        pen = 1
        if self.NonComplete(ID): # Apply penalty if they haven't filled out the survey
            pen = pen - (globals.config["Penalty_NonComplete"]/100)
        else:
            if self.PartialComplete(ID): pen = pen - (globals.config["Penalty_PartialComplete"]/100)
            if self.SelfPerfect(ID): pen = pen - (globals.config["Penalty_SelfPerfect"]/100)
            if self.PeersAllZero(ID): pen = pen - (globals.config["Penalty_PeersAllZero"]/100)
            if self.days_late(ID, globals.config["weekdays_only"]) > 0 and globals.config['Penalty_Late_Custom']: pen = pen - (globals.config["Penalty_PerDayLate"]/100)*self.days_late(ID, globals.config["weekdays_only"]) # Otherwise, simply use the Canvas default late system
        return(max(pen,0))
    
    def adj_score(self, ID, apply_penalty = True):
        '''A function to calculate the adjusted score. Also provides a status regarding whether the assignment should be excused, given the shared team mark, or given a zero.'''
        # status can be: normal, noncomplete, excused, original_score, zero, error
    
        adj = {"score": None, "status": "normal", "penalty": 1, "seconds_late": 0, "adjustment": False}
        if ID in globals.adj_dict: 
            adj = globals.adj_dict[ID]
        else:
            if (globals.config["ScoreType"] == 1): # Simply return the mean score the ratee received
                if self.n_ratings(ID) >= globals.config["MinimumPeers"]:
                    rat_total = self.ratee_total(ID)
                    
                    # Apply policy regarding self-rating
                    if globals.config["SelfVote"]==0: rat_total[0] = None # Get rid of the self-rating
                    if len(self.rater_total(ID)) > 1:
                        if globals.config["SelfVote"]==2: rat_total[0] = nanmean(self.rater_total(ID)[1:]) # Automatic mean vote
                        if globals.config["SelfVote"]==3: rat_total[0] = npmin([rat_total[0], nanmean(self.rater_total(ID)[1:])]) # Person cannot give themselves more than the average mark they have assigned others
                    else: # Also get rid of the self-rating if the person has not rated anyone else
                        rat_total[0] = None
                    
                    # Calculate the Peer Mark
                    rat_mean = nanmean(rat_total) # Calculate the mean total score
                    adj["score"] = rat_mean / len(self.tw_list) # Divide by the number of items
                    if globals.config["RescaleTo"] is not None and globals.config["RescaleTo"] != 0 and type(adj["score"]) is not str:
                        adj["score"] = (adj["score"]*globals.config["RescaleTo"])/(globals.config["PointsPossible"]-1) # Rescale the score to the desired range
                else: # Apply policy is minimum peers are not met
                    if globals.config["MinimumPeersPolicy"] == 1: adj["status"] = "excused"
                    if globals.config["MinimumPeersPolicy"] == 2:
                        try:
                            adj["score"] = float(canvas.data[str(ID)]["orig_grade"])
                            if globals.config["RescaleTo"] is not None and globals.config["RescaleTo"] != 0: adj["score"] = adj["score"]*(globals.config["RescaleTo"]/globals.session["assn_points_possible"]) # Rescale the score to the desired range
                            adj["status"] = "original_score"
                        except:
                            print("Original score cannot be found")
                            adj["status"] = "error" # The user's original score cannot be found
                    if globals.config["MinimumPeersPolicy"] == 3: 
                        adj["status"] = "zero" # Substitute with a zero
                        adj["score"] = 0
            elif (globals.config["ScoreType"] == 2): # Return an adjusted score based on group mark
                try:
                    adj["score"] = float(canvas.data[str(ID)]["orig_grade"])
                except:
                    adj["status"] = "error"
                    return adj
                ratio = self.ratee_ratio(ID)
                if globals.config["SelfVote"] == 0: ratio[0] = None
                if globals.config["SelfVote"] == 2: ratio[0] = 1 # For Option 2, automatically assign a vote of 1
                if globals.config["SelfVote"] == 3: ratio[0] = min(1, ratio[0]) # For Option #3, do not allow self-votes to raise their score
                if self.n_ratings(ID) >= globals.config["MinimumPeers"]: # Number of legit scores must exceed minimum 
                    if abs(nanmean(ratio)-1)*100 >= globals.config["MinimumAdjustment"]:
                        adj["score"] = adj["score"]*nanmean(ratio)
                        adj["adjustment"] = True
                else:
                    if globals.config["MinimumPeersPolicy"] == 1: adj["status"] = "excused"
                    if globals.config["MinimumPeersPolicy"] == 2: adj["status"] = "original_score"
                    if globals.config["MinimumPeersPolicy"] == 3: adj["status"] = "zero"
                adj["score"] = min(adj["score"], globals.session["assn_points_possible"]) # Score is not allowed to exceed number of points possible.
                if globals.config["RescaleTo"] is not None and globals.config["RescaleTo"] != 0 and adj["score"] is not None: adj["score"] = adj["score"]*(globals.config["RescaleTo"]/globals.session["assn_points_possible"]) # Rescale the score to the desired range
            try:
                adj["penalty"] = self.penalty(ID)
                adj["seconds_late"] = self.days_late(ID, globals.config["weekdays_only"], integer = False)*86400
            except:
                pass
            globals.adj_dict[ID] = adj
        return(adj)
    
    # Comments - a function to provide the students with their scores on their peer ratings, as well as peer feedback
    def comments(self, ID):
        comment = "Feedback is unavailable because not enough of your team members completed the peer evaluation.\n\n"
        
        # Quit if one of the following conditions is met.
        if self.adj_score(ID)["status"] == "zero":
            comment = comment + "As a result, you were automatically assigned a zero for the peer mark."
            return comment
        if self.adj_score(ID)["status"] == "excused" and self.adj_score(ID)["penalty"] > 0:
            comment = comment + "As a result, you were excused from this assessment. Please note, you have not been penalised. Your final grade will be based on the weighted average of your other assignments."
            return comment
        if self.adj_score(ID)["status"] == "error":
            comment = "No peer mark is available. This error may occur if you are not enrolled in any groups."
            return comment
            
        # Otherwise, keep going...
        
        if self.adj_score(ID)["status"] == "original_score":
            orig_score = cleanupNumber(canvas.data[str(ID)]["orig_grade"])
            comment = comment + "As there were insufficient team members for the peer evaluation, your shared team mark, " + str(round(orig_score,2)) + " out of " + str(globals.session["assn_points_possible"]) + " (" + str(round(100 * orig_score / globals.session["assn_points_possible"], 1)) + "%), was substituted as your peer mark.\n\n"
        penalty_text = ""
        if self.NonComplete(ID)==True:
            if globals.config["Penalty_NonComplete"] > 0: penalty_text = penalty_text + "-" + str(globals.config["Penalty_NonComplete"]) + "% because you did not complete the Peer Evaluation survey."
        else:
            if self.PartialComplete(ID) == True and globals.config["Penalty_PartialComplete"] > 0: penalty_text = penalty_text + "-" + str(globals.config["Penalty_PartialComplete"]) + "% because you only partially completed the Peer Evaluation survey.\n"
            if self.SelfPerfect(ID) == True and globals.config["Penalty_SelfPerfect"] > 0: penalty_text = penalty_text + "-" + str(globals.config["Penalty_SelfPerfect"]) + "% because you assigned yourself a perfect score on the Peer Evaluation survey across all questions.\n"
            if self.PeersAllZero(ID) == True and globals.config["Penalty_PeersAllZero"] > 0: penalty_text = penalty_text + "-" + str(globals.config["Penalty_PeersAllZero"]) + "% because you gave all of your peers the bottom score on the Peer Evaluation survey across all questions.\n"
            if self.days_late(ID, globals.config["weekdays_only"]) is not None:
                if self.days_late(ID, globals.config["weekdays_only"]) > 0 and globals.config["Penalty_PerDayLate"] > 0 and globals.config["Penalty_Late_Custom"] == True: penalty_text = penalty_text + "-" + str(globals.config["Penalty_PerDayLate"] * self.days_late(ID, globals.config["weekdays_only"])) + "% because you submitted your evaluation " + str(self.days_late(ID, globals.config["weekdays_only"])) + " day(s) late.\n"
        if len(penalty_text) > 0: comment = comment + "The following penalties have been subtracted from your Peer Mark:\n" + penalty_text
        if self.n_ratings(ID) < globals.config["MinimumPeers"]: return(comment)
        comment = ""
        general_comments = self.general_feedback(ID)[0]
    
        if len(general_comments) > 0:
            comment = comment + "The following general comments about your team (not about yourself personally) were provided by your peers:\n\n"
            shuffle(general_comments) # Shuffle the general comments so the rater is not identifiable
            for temp_comment in general_comments:
                comment = comment + '"' + temp_comment + '"\n\n'
        peer_feedback = remove_blanks(self.feedback(ID)[0])
        if len(peer_feedback) > 0:
            comment = comment + "You have received the following specific feedback from your peers:\n\n"
            shuffle(peer_feedback) # Shuffle the order of the feedback so the rater is not identifiable
            for feed in peer_feedback: comment = comment + '"' + str(feed) + '"\n\n'
        comment = comment + "You have received the following ratings from your peers:\n\n"
        rr = self.ratings_received(ID)
        for y, TeamWork in enumerate(self.tw_names):
            comment = comment + TeamWork + "\n"
            self_rating = rr[y][0]
            peer_rating = nanmean(rr[y][1:])
            if self_rating is not None: comment = comment + "Your Self-Rating: {}/{}\n".format(round(globals.config["RescaleTo"]*self_rating/(globals.config["PointsPossible"]-1),1),globals.config["RescaleTo"])
            if peer_rating is not None: comment = comment + "Average Rating Received from Peers: {}/{}\n\n".format(round(globals.config["RescaleTo"]*peer_rating/(globals.config["PointsPossible"]-1),1),globals.config["RescaleTo"])
        if globals.config["ScoreType"] == 1:
            comment = comment + "Your Peer Mark was calculated as the average peer rating you received.\n\n"
        elif globals.config["ScoreType"] == 2:
            orig_score = cleanupNumber(canvas.data[str(ID)]["orig_grade"])
            comment = comment + "Your mark was calculated using the following formula. "
            comment = comment + "You received " + str(round(orig_score,2)) + " out of " + str(globals.session["assn_points_possible"]) + " (" + str(round(100 * orig_score / globals.session["assn_points_possible"], 1)) + "%) for your team assignment. "
            comment = comment + "Based on this, you were assigned an initial mark of " + str(round(globals.config["RescaleTo"] * orig_score / globals.session["assn_points_possible"],1)) + " out of " + str(globals.config["RescaleTo"]) + " ("+ str(round(100 * orig_score / globals.session["assn_points_possible"], 1)) + "%). "
            comment = comment + "You received peer ratings that were "
            ratio = self.ratee_ratio(ID)
            if globals.config["SelfVote"] == 0: ratio[0] = None
            if globals.config["SelfVote"] == 2: ratio[0] = 1 # For Option 2, automatically assign a vote of 1
            if globals.config["SelfVote"] == 3: ratio[0] = min(1, ratio[0]) # For Option #3, do not allow self-votes to raise their score
            ratio = nanmean(ratio)
            a_score = ratio * globals.config["RescaleTo"] * orig_score / globals.session["assn_points_possible"]
            ratio = ratio - 1
            ratio = round(ratio * 100,1)
            if self.adj_score(ID)["adjustment"] == False:
                comment = comment + "within " + str(globals.config["MinimumAdjustment"]) + "% of the average rating received by your team members. "
            elif ratio < 0: 
                comment = comment + str(abs(ratio)) + "% lower than the average rating received by your other team members. "
            elif ratio == 0: 
                comment = comment + "equal to  the average rating received by your other team members. "
            elif ratio > 0: 
                comment = comment + str(ratio) + "% higher than  the average rating received by your other team members. "
            if ratio != 0 and self.adj_score(ID)["adjustment"] == True:
                comment = comment + "Your initial mark was therefore adjusted from " + str(round(globals.config["RescaleTo"] * orig_score / globals.session["assn_points_possible"],2)) + " to your final Peer Mark of " + str(round(a_score, 2)) + ".\n\n"
            else: 
                comment = comment + "Your initial mark was therefore not adjusted.\n\n"
        if globals.config["SelfVote"] == 0:
            comment = comment + "Your self-rating is not counted towards your Peer Mark, and is provided as feedback only.\n\n"
        elif globals.config["SelfVote"] == 1:
            comment = comment + "Your self-rating is included as part of your Peer Mark.\n\n"
        elif globals.config["SelfVote"] == 2:
            comment = comment + "Your self-rating is not counted towards your Peer Mark. However, you received an automatic \"self-vote\" that is equivalent to an equal contribution to the team project based on your ratings of other team members.\n\n"
        elif globals.config["SelfVote"] == 3:
            comment = comment + "Your self-rating is included as part of your Peer Mark. It is capped at the average rating you gave your other team members.\n\n"
        if len(penalty_text) > 0: comment = comment + "Your Peer Mark has received the following penalties:\n" + penalty_text
        return(comment)