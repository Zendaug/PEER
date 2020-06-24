# -*- coding: utf-8 -*-

try:
    f = open("Version History.txt", "r")
    updates = f.read()
    f.close()
    print(updates)
except:
    print("Unable to locate Version History.")

print("\n(C) Tim Bednall 2019-2020.")

import tkinter
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from tkcalendar import DateEntry
from canvasapi import Canvas
from random import shuffle
import pandas as pd
import numpy as np
import json
from datetime import datetime, timedelta
import time
import os
import requests
import sys
import subprocess

#%% Global functions and variables

canvas = None
qualtrics = ""
orig_data = ""
tw_list = [] # Questions that make up the teamwork score
tw_names = [] # The names of the teamwork variables 
other_list = [] # The names of other variables

# Set up dictionary for ratios and adjusted scores for fast lookup
ratio_dict = {} 
adj_dict = {}

def clean_text(text):
    text = text.strip()
    text = text.replace("'", "''")
    return(text)
    
def find_all(text, search_for):
    find_list = []
    start = 0
    for a in range(0, text.count(search_for)):
        find_list.append(text.find(search_for, start))
        start = find_list[-1]+1
    return(find_list)

def only_numbers(char, event_type = ""):
    return char.isdigit()

def unique(list1):
    unique_list = []
    for item in list1:
        if item not in unique_list: unique_list.append(item)
    return(unique_list)

def replace_multiple(str_text, rep_list, str_rep = ""):
    for item in rep_list: str_text = str_text.replace(str(item), str_rep)
    return(str_text)

def save_config():
    try:
        f = open("config.txt", "w")
        f.write(json.dumps(config, sort_keys=True, separators=(',\n', ':')))
        f.close()   
    except:
        print("Unabled to save configuration file (config.txt).")

# Load the configuration file
# First of all, load the hard defaults
hard_defaults = {
    'API_URL': "",
    'API_TOKEN': "",
    'EmailSuffix': "",
    'ScoreType': 2,
    'SelfVote': 2,
    'PeerMark_Name': "Peer Mark",
    'MinimumPeers': 2,
    'Penalty_NonComplete': 100,
    'Penalty_PartialComplete': 0,
    'Penalty_SelfPerfect': 0,
    'Penalty_PeersAllZero': 0,
    'Penalty_PerDayLate': 0,
    'Exclude_PartialComplete': 0,
    'Exclude_SelfPerfect': 0,
    'Exclude_PeersAllZero': 0,
    'SaveLong': 0,
    'SaveWide': 0,
    'PointsPossible': 5,
    'RescaleTo': 5,
    'publish_assignment': True,
    'weekdays_only': True,
    'subject_confirm': "Confirming Group Membership",
    'message_confirm': "Dear [firstnames_list],\n\nI am writing to confirm your membership in: [name].\n\nAccording to our records, the members of this group include:\n\n[members_bulletlist]\n\nIf there are any errors in our records -- such as missing group members or people incorrectly listed as group members -- please contact me as soon as possible. You may ignore this message if these records are correct.\n\nWe encourage you to visit your group's homepage, a platform where you can collaborate with your fellow members. Your group's homepage is located at: [group_homepage]",
    'subject_nogroup' : "Lack of Group Membership",
    'message_nogroup' : "Dear [first_name],\n\nI am writing because according to our records, you are not currently listed as a member of any student team for this unit.\n\nAs you would be aware, there is a group assignment due later in the semester. If you do not belong to a team, you will be unable to complete the assignment.\n\nIf there is another circumstance that would prevent you from joining a group project, please let me know as soon as possible.\n\nPlease let me know if there is anything I can do to support you in the meantime.",
    'subject_invitation': "Peer Evaluation now available",
    'message_invitation': "Dear [First Name],\n\nThe Peer Evaluation survey is now available.\n\nYou may access it via this link: [Link]\n\nThis link is unique to yourself, so please do not share it with other people.",
    'subject_reminder': "Reminder: Peer Evaluation due",
    'message_reminder': "Dear [First Name],\n\nThis is a courtesy reminder to complete the Peer Evaluation survey. According to our records, you have not yet completed this survey.\n\nYou may access the survey via this link: [Link]\n\nThis link is unique to yourself, so please do not share it with other people.",
    'subject_thankyou' : "Thank you for completing the Peer Evaluation",
    'message_thankyou' : "Dear [First Name],\n\nThank you for completing the Peer Evaluation. There is no need to take any further action."
    }
config = hard_defaults

# Overwrite hard defaults with soft defaults
try:
    f = open("defaults.txt", "r")
    defaults = json.loads(f.read())
    f.close()
    for item in defaults: config[item] = defaults[item]
except:
    defaults = hard_defaults

# Overwrite defaults with user's own configuration file
try:
    f = open("config.txt", "r")
    user_config = json.loads(f.read())
    f.close()
    for item in user_config: config[item] = user_config[item]
except:
    print("Creating user configuration file from defaults...")

# Overwrite the user's configuration with any updates
try:
    f = open("update.txt", "r")
    user_config = json.loads(f.read())
    f.close()
    for item in user_config: config[item] = user_config[item]
    os.remove("update.txt")
    save_config()
except:
    pass

session = {}

def submission_date(ID):
    '''This function returns the date the Qualtrics survey was submitted (the EndDate field).'''
    if (qualtrics["Student_id"]==ID).any() == False: return(np.nan)
    y = qualtrics[qualtrics["Student_id"] == ID].index[0]
    try:
        return(datetime.strptime(qualtrics.loc[y, "EndDate"], '%Y-%m-%d %H:%M:%S'))
    except:
        return(np.nan)

def days_late(ID, weekdays_only = True):
    '''This function returns the number of days late an assignment was (or 0 if it was submitted before time). By default, it only counts weekdays.'''
    sub_date = submission_date(ID)
    if "DueDate" in session and type(sub_date) is datetime:
        due_date = session["DueDate"]
        sub_date = datetime(sub_date.year, sub_date.month, sub_date.day, 0,0,0)
        due_date = datetime(due_date.year, due_date.month, due_date.day, 0,0,0)
        d_late = max(0, (sub_date - due_date).days)
        for a in range(0,d_late*weekdays_only):
            d_late -= ((sub_date - timedelta(days=a)).weekday() >= 5)
        return(d_late)
    else:
        return(np.nan)

def ratings_given(ID):
    '''This function returns all of the ratings given by a rater (ID) in a list, including self-ratings'''
    if (qualtrics["Student_id"]==ID).any() == False: return([np.nan])
    y = qualtrics[qualtrics["Student_id"] == ID].index[0]
    n_peers = int(float(qualtrics[(qualtrics["Student_id"]==ID)]["NumPeers"]))
    ratings_temp = []
    for peer in range(0, n_peers+1):
        ratings_temp.append([])
        for item in tw_list:
            #print(item + '_' + str(peer))
            ratings_temp[peer].append(qualtrics.loc[y, item + '_' + str(peer+1)])
    return(ratings_temp)

def rater_total(ID, NoSelf = True):
    '''A function that calculates the total number of points allocated by a rater to each peer'''
    if (qualtrics["Student_id"]==ID).any() == False: return([np.nan])
    tot = []
    peer_lower = 1
    temp_ratings = ratings_given(ID)
    if (config["SelfVote"] == 1 or NoSelf == False): peer_lower = 0 # Count the person's own vote if the SelfVote setting is set to 1
    for peer in range(peer_lower, len(temp_ratings)):
        #print(temp_ratings[peer])
        tot.append(np.nansum(temp_ratings[peer]))
    return(tot)


def rater_peerID(ID):
    ''' Returns the IDs of the peer raters of a person in a list using 2 methods'''
    tot = []
    # Method 1 - If they are in the Qualtrics database, simply grab their peers
    if (qualtrics["Student_id"]==ID).any(): # If the rater exists in the qualtrics database, then retrieve their peers from their line
        y = qualtrics[qualtrics["Student_id"] == ID].index[0]
        for peer in range(1, qualtrics.loc[y, "NumPeers"]+1):
            tot.append(qualtrics.loc[y, "Peer" + str(peer) + '_id'])
    else: # Method 2 - Look through all of their peers and grab IDs          
        for y in range(0, qualtrics.shape[0]):
            for peer in range(1, qualtrics.loc[y, "NumPeers"]+1):
                if (qualtrics.loc[y, "Peer" + str(peer) + '_id'] == ID):
                    tot.append(qualtrics.loc[y, "Student_id"])
    #if (len(tot) > 0): tot.insert(0, ID) # Add the person's own ID to the beginning of the list
    tot.insert(0, ID)
    return(tot)

def max_peers():
    '''A function that shows the largest number of peers'''
    maxp = 0
    for ID in orig_data["ID"]: maxp = np.max([len(rater_peerID(ID)), maxp])
    return(maxp) 

def NonComplete(ID):
    '''A function to detect whether a person has not completed the peer survey'''
    return(not((qualtrics["Student_id"]==ID).any()))

def PartialComplete(ID):
    '''A function to detect whether a person has only partially completed the peer survey'''
    return(np.isnan(np.sum(ratings_given(ID))))

def SelfPerfect(ID):
    '''A function to detect whether a person has given themselves all top ratings'''
    return(np.nanmean(ratings_given(ID)[0])==4)

def PeersAllZero(ID):
    '''A function to detect whether a person has given their peers all zeroes. It assumes that the minimum number of peers has been met.'''
    return(np.nansum(rater_total(ID, True))==0 and len(rater_peerID(ID)) > config["MinimumPeers"])
    
# A function that returns all of the ratings received by a person
def ratings_received(ID, rater = 0):
    rr = []
    for rater_ID in rater_peerID(ID):
        if (rater in [0, rater_ID]):
            rr.append([])
            if (qualtrics["Student_id"]==rater_ID).any() and not(
                    (config["Exclude_PartialComplete"] == 1 and PartialComplete(ID) == True) or
                    (config["Exclude_SelfPerfect"] == 1 and SelfPerfect(ID) == True) or
                    (config["Exclude_PeersAllZero"] == 1 and PeersAllZero(ID) == True)):
                y = qualtrics[qualtrics["Student_id"] == rater_ID].index[0]
                if (ID == rater_ID):
                    x = 1
                else: # Identify peer raters
                    for peer in range(1, qualtrics.loc[y, "NumPeers"]+1):
                        if (qualtrics.loc[y, "Peer" + str(peer) + '_id'] == ID): x = peer + 1
                for item in tw_list:
                    try:
                        rr[-1].append(qualtrics.loc[y, item + '_' + str(x)])
                    except:
                        rr[-1].append(np.nan)
            else:
                for item in tw_list:
                    rr[-1].append(np.nan)
    return(rr)

# Ratee Total: A function that calculates the total number of points received from all peers, including oneself
def ratee_total(ID):
    tot = []
    for a,rater in enumerate(ratings_received(ID)):
        if np.sum(np.isnan(ratings_received(ID)[a]))==len(ratings_received(ID)[a]):
            tot.append(np.nan)
        else:
            tot.append(np.nansum(ratings_received(ID)[a]))
    return(tot)

# A function that returns the number of valid ratings received from peers. It does not count the person's self-rating
def n_ratings(ID):
    if len(rater_peerID(ID)) > 1:
        return(np.sum(np.isnan(ratee_total(ID)[1:])==False))
    else:
        return(0)

# Ratee Ratio: A function that calculates the ratio of points received from a peer
def ratee_ratio(ID):
    if ID in ratio_dict: return(ratio_dict[ID])
    ratio = []
    raters = rater_peerID(ID)
    ratee_total_temp = ratee_total(ID)
    for a, rater_ID in enumerate(raters):
        rater_temp = (qualtrics["Student_id"]==rater_ID)
        if (rater_temp).any(): # If the rater exists in the qualtrics database, then retrieve their score
            y = qualtrics[rater_temp].index[0]
            npeers = qualtrics.loc[y, "NumPeers"]
            if config["SelfVote"] == 1: npeers = npeers + 1 # Count an additional peer if SelfVote is allowed
            rat_total = ratee_total_temp[a]
            exp_total = 0 + (np.nansum(rater_total(rater_ID)) / npeers)
            if exp_total==0: rat_exp = np.nan
            else: rat_exp = rat_total / exp_total
        else:
            rat_exp = np.nan
        ratio.append(rat_exp)
    ratio_dict[ID] = ratio
    return(ratio)

# Feedback: A function to locate the feedback provided to each person
def feedback(ID):
    if "Feedback#1_1_1" in qualtrics.columns: # To maintain compatability with the old format
        feed_peer = "Feedback#1_x_1"
        feed_inst = "Feedback#2_x_1"
    else:
        feed_peer = "Feedback_x_1"
        feed_inst = "Feedback_x_2"
    fb = [[],[]]
    if (qualtrics["Student_id"]==ID).any():
        y = qualtrics[qualtrics["Student_id"] == ID].index[0]
        fb[0].append(qualtrics.loc[y, feed_peer.replace("x","10")]) # Any other feedback "Feedback#1_10_1"
        fb[1].append(qualtrics.loc[y, feed_inst.replace("x","10")]) # Any other feedback "Feedback#2_10_1"
    else:
        fb[0].append(np.nan) # Any other feedback
        fb[1].append(np.nan) # Any other feedback
    for rater_ID in rater_peerID(ID)[1:]:
        if (qualtrics["Student_id"] == rater_ID).any():
            y = qualtrics[qualtrics["Student_id"] == rater_ID].index[0]
            for peer in range(1, qualtrics.loc[y, "NumPeers"]+1): # Identify peer raters
                if (qualtrics.loc[y, "Peer" + str(peer) + '_id'] == ID): x = peer
            fb[0].append(qualtrics.loc[y, feed_peer.replace("x",str(x))])
            fb[1].append(qualtrics.loc[y, feed_inst.replace("x",str(x))])
        else:
            fb[0].append(np.nan)
            fb[1].append(np.nan)           
    return(fb)

def penalty(ID):
    '''A function to calculate the % deducted as a penalty'''
    pen = 1
    if NonComplete(ID): # Apply penalty if they haven't filled out the survey
        pen = pen - (config["Penalty_NonComplete"]/100)
    else:
        if PartialComplete(ID): pen = pen - (config["Penalty_PartialComplete"]/100)
        if SelfPerfect(ID): pen = pen - (config["Penalty_SelfPerfect"]/100)
        if PeersAllZero(ID): pen = pen - (config["Penalty_PeersAllZero"]/100)
        if days_late(ID, config["weekdays_only"]) > 0: pen = pen - (config["Penalty_PerDayLate"]/100)*days_late(ID, config["weekdays_only"])
    return(max(pen,0))

def adj_score(ID, apply_penalty = True):
    '''A function to calculate the adjusted score'''
    if ID in adj_dict: 
        orig_score = adj_dict[ID]
    else:
        if (config["ScoreType"] == 1): # Simply return the mean score the ratee received
            if n_ratings(ID) >= config["MinimumPeers"]:
                rat_total = ratee_total(ID)
                if config["SelfVote"]==0: rat_total[0] = np.nan # Get rid of the self-rating
                if len(rater_total(ID)) > 1:
                    if config["SelfVote"]==2: rat_total[0] = np.nanmean(rater_total(ID, False)[1:]) # Automatic mean vote
                    if config["SelfVote"]==3: rat_total[0] = np.min([rat_total[0], np.nanmean(rater_total(ID, False)[1:])]) # Person cannot give themselves more than the average mark they have assigned others
                else:
                    rat_total[0] = np.nan
                rat_mean = np.nanmean(rat_total) # Take the mean total score
                orig_score = rat_mean / len(tw_list) # Divide by the number of items
                if np.isnan(config["RescaleTo"]) == False and config["RescaleTo"] != 0 and type(orig_score) is not str: orig_score = orig_score*config["RescaleTo"]/(session["points_possible"]-1) # Rescale the score to the desired range
            else:
                orig_score = "Insufficient ratings received (" + str(n_ratings(ID)) + ")"
        elif (config["ScoreType"] == 2): # Return an adjusted scorem based on group mark
            orig_score = float(orig_data[orig_data["ID"]==ID]["orig_grade"])
            ratio = ratee_ratio(ID)
            if config["SelfVote"] == 0: ratio[0] = np.nan
            if config["SelfVote"] == 2: ratio[0] = 1 # For Option 2, automatically assign a vote of 1
            if config["SelfVote"] == 3: ratio[0] = min(1, ratio[0]) # For Option #3, do not allow self-votes to raise their score
            if n_ratings(ID) >= config["MinimumPeers"]: # Number of legit scores must exceed minimum peers
                orig_score = orig_score*np.nanmean(ratio)
            orig_score = np.min([orig_score, session["points_possible"]]) # Score is not allowed to exceed number of points possible.
            if np.isnan(config["RescaleTo"]) == False and config["RescaleTo"] != 0 and type(orig_score) is not str: orig_score = orig_score*(config["RescaleTo"]/session["points_possible"]) # Rescale the score to the desired range
        adj_dict[ID] = orig_score
    if apply_penalty == True: 
        try:
            orig_score = orig_score*penalty(ID) # Apply the penalty if the person does not complete the peer evaluation
        except:
            pass
    return(orig_score)

# Comments - a function to provide the students with their scores on their peer ratings, as well as peer feedback
def comments(ID):
    comment = "Feedback is unavailable because not enough of your team members completed the peer evaluation.\n\n"
    penalty_text = ""
    if NonComplete(ID)==True:
        if config["Penalty_NonComplete"] > 0: penalty_text = penalty_text + "-" + str(config["Penalty_NonComplete"]) + "% because you did not complete the Peer Evaluation survey."
    else:
        if PartialComplete(ID) == True and config["Penalty_PartialComplete"] > 0: penalty_text = penalty_text + "-" + str(config["Penalty_PartialComplete"]) + "% because you only partially completed the Peer Evaluation survey.\n"
        if SelfPerfect(ID) == True and config["Penalty_SelfPerfect"] > 0: penalty_text = penalty_text + "-" + str(config["Penalty_SelfPerfect"]) + "% because you assigned yourself a perfect score on the Peer Evaluation survey across all questions.\n"
        if PeersAllZero(ID) == True and config["Penalty_PeersAllZero"] > 0: penalty_text = penalty_text + "-" + str(config["Penalty_PeersAllZero"]) + "% because you gave all of your peers the bottom score on the Peer Evaluation survey across all questions.\n"
        if days_late(ID, config["weekdays_only"]) > 0 and config["Penalty_PerDayLate"] > 0: penalty_text = penalty_text + "-" + str(config["Penalty_PerDayLate"] * days_late(ID, config["weekdays_only"])) + "% because you submitted your evaluation " + str(days_late(ID, config["weekdays_only"])) + " day(s) late.\n"
    if len(penalty_text) > 0: comment = comment + "The following penalties have been subtracted from your Peer Mark:\n" + penalty_text
    if n_ratings(ID) < config["MinimumPeers"]: return(comment)
    comment = ""
    general_comments = []
    for peer in rater_peerID(ID)[1:]:
        if type(feedback(peer)[0][0])==str:
            if len(feedback(peer)[0][0])>0:
                general_comments.append(feedback(peer)[0][0])
    if len(general_comments) > 0:
        comment = comment + "The following general comments about your team (not about yourself personally) were provided by your peers:\n\n"
        shuffle(general_comments) # Shuffle the general comments so the rater is not identifiable
        for temp_comment in general_comments:
            comment = comment + '"' + temp_comment + '"\n\n'
    if sum(pd.notnull(feedback(ID)[0][1:])) > 0:
        comment = comment + "You have received the following specific feedback from your peers:\n\n"
        peer_feedback = feedback(ID)[0][1:]
        shuffle(peer_feedback) # Shuffle the order of the feedback so the rater is not identifiable
        for feed in peer_feedback:
            if pd.isnull(feed) == False: comment = comment + '"' + feed + '"\n\n'
    comment = comment + "You have received the following ratings from your peers:\n\n"
    for x, TeamWork in enumerate(tw_names):
        comment = comment + TeamWork + "\n"
        rr = ratings_received(ID)
        peer_rating = []
        for y in range(1,len(rr)): peer_rating.append(rr[y][x])
            #peer_rating = peer_rating + rr[y][x]
        #peer_rating = peer_rating / y
        peer_rating = np.nanmean(peer_rating)
        if np.isnan(rr[0][x]) == False: comment = comment + "Your Self-Rating: " + str(rr[0][x]+1) + "/5\n"
        comment = comment + "Average Rating Received from Peers: " + str(round(peer_rating,2)+1) + "/5\n\n"
    if config["ScoreType"] == 2:
        orig_score = float(orig_data[orig_data["ID"]==ID]["orig_grade"])
        comment = comment + "Your mark was calculated using the following formula. "
        comment = comment + "You received " + str(orig_score) + " out of " + str(session["points_possible"]) + " (" + str(round(100 * orig_score / session["points_possible"], 1)) + "%) for your team assignment. "
        comment = comment + "Based on this, you were assigned an initial mark of " + str(round(config["RescaleTo"] * orig_score / session["points_possible"],1)) + " out of " + str(config["RescaleTo"]) + " ("+ str(round(100 * orig_score / session["points_possible"], 1)) + "%). "
        comment = comment + "You received peer ratings that were "
        ratio = ratee_ratio(ID)
        if config["SelfVote"] == 0: ratio[0] = np.nan
        if config["SelfVote"] == 2: ratio[0] = 1 # For Option 2, automatically assign a vote of 1
        if config["SelfVote"] == 3: ratio[0] = min(1, ratio[0]) # For Option #3, do not allow self-votes to raise their score
        ratio = np.nanmean(ratio)
        a_score = ratio * config["RescaleTo"] * orig_score / session["points_possible"]
        ratio = ratio - 1
        ratio = round(ratio * 100,1)
        if ratio < 0: comment = comment + str(abs(ratio)) + "% lower than"
        if ratio == 0: comment = comment + "equal to"
        if ratio > 0: comment = comment + str(ratio) + "% higher than"
        comment = comment + " the average rating received by your other team members. "
        if ratio != 0:
            comment = comment + "Your initial mark was therefore adjusted from " + str(round(config["RescaleTo"] * orig_score / session["points_possible"],2)) + " to your final Peer Mark of " + str(round(a_score, 2)) + ".\n\n"
        else: 
            comment = comment + "Your initial mark was therefore not adjusted.\n\n"
    if len(penalty_text) > 0: comment = comment + "Your Peer Mark has received the following penalties:\n" + penalty_text
    return(comment)

def cleanupURL(URL):
    URL = URL.strip().lower()
    if URL[-1] != "/": URL += "/"
    URL = URL.replace('http:', 'https:')
    return(URL)

#%% GraphQL
class GraphQL():
    def __init__(self, course_id = 0, groupset_id = 0):
        self.course_id = course_id
        self.groupset_id = groupset_id
        self.cache = {}
    
    def query(self, query_text):
        global mm
        query = {'access_token': config["API_TOKEN"], 'query': query_text}
        for attempt in range(5):
            try:
                result = requests.post(url = "{}api/graphql".format(config["API_URL"]), data = query).json()
            except:
                print("Trying to connect to Canvas via GraphQL. Attempt {} of 5.".format(attempt+1))
            else:
                return(result)
        messagebox.showinfo("Error", "Cannot connect to Canvas via GraphQL. There may be an Internet connectivity problem.")
        mm.master.destroy()
        sys.exit()
    
    def courses(self):
        if "courses" in self.cache: return(self.cache["courses"])
        temp = self.query('query MyQuery {allCourses {_id\nname}}')["data"]["allCourses"]
        temp2 = {}
        for course in temp:
            temp2[int(course["_id"])] = {}
            temp2[int(course["_id"])]["name"] = course["name"]
        df = {}
        for a in sorted(temp2.keys(), reverse = True):
            df[str(a)] = temp2[a]
        self.cache["courses"] = df
        return(df)

    def group_sets(self):
        if "groupset_" + str(self.course_id) not in self.cache:
            self.cache["groupset_" + str(self.course_id)] = self.query('query MyQuery {course(id:"' + str(self.course_id) + '") {groupSetsConnection {nodes {_id\nname}}}}')["data"]["course"]["groupSetsConnection"]["nodes"]
        return(self.cache["groupset_" + str(self.course_id)])

    def sections(self):
        if "sections_" + str(self.course_id) not in self.cache:
            query_result = self.query('query MyQuery {course(id: "' + str(self.course_id) + '") {sectionsConnection {nodes {_id\nname}}}}')["data"]["course"]["sectionsConnection"]["nodes"]
            self.cache["sections_" + str(self.course_id)] = {}
            for section in query_result:
                self.cache["sections_" + str(self.course_id)][section["_id"]] = {}
                self.cache["sections_" + str(self.course_id)][section["_id"]]["name"] = section["name"]
        return(self.cache["sections_" + str(self.course_id)])
    
    def students(self):
        if "students_" + str(self.course_id) not in self.cache:
            query_result = self.query('query MyQuery {course(id: "' + str(self.course_id) + '") {enrollmentsConnection {nodes {type\nuser {_id\nname\nemail}}}}}')
            self.cache["students_" + str(self.course_id)] = {"id": [], "name": [], "email": []}
            for student in query_result["data"]["course"]["enrollmentsConnection"]["nodes"]:
                if (student["type"] == "StudentEnrollment"):
                    if (student["user"]["_id"] not in self.cache["students_" + str(self.course_id)]["id"]):
                        self.cache["students_" + str(self.course_id)]["id"].append(student["user"]["_id"])
                        self.cache["students_" + str(self.course_id)]["name"].append(student["user"]["name"])
                        self.cache["students_" + str(self.course_id)]["email"].append(student["user"]["email"])
        return(self.cache["students_" + str(self.course_id)])
    
    def groups(self):
        if "groups_" + str(self.course_id) in self.cache: return(self.cache["groups_" + str(self.course_id)])
        query_result = self.query('query MyQuery {course(id: "' + str(self.course_id) + '") {groupSetsConnection {nodes {_id\ngroupsConnection {nodes {_id\nname\nmembersConnection {nodes {user {_id\nname}}}}}}}}}')
        try:
            for groupset in query_result["data"]["course"]["groupSetsConnection"]["nodes"]:
                if (groupset["_id"]==str(self.groupset_id)):
                    temp = groupset["groupsConnection"]["nodes"]
                    df = {}
                    for group in temp:
                        df[group["_id"]] = {"name": group["name"], "users": []}
                        for user in group["membersConnection"]["nodes"]:
                            df[group["_id"]]["users"].append(user["user"]["_id"])
                    self.cache["groups_" + str(self.course_id)] = df
                    return(df)
                    break
        except:
            return({})
        return({})
    
    def students_comprehensive(self):
        if "students_" + str(self.course_id) + "_" + str(self.groupset_id) in self.cache:
            return(self.cache["students_" + str(self.course_id) + "_" + str(self.groupset_id)])
        query_result = self.query('query MyQuery {course(id: "' + str(self.course_id) + '") {enrollmentsConnection {nodes {type\nuser {_id\nname\nemail}\nsection {_id}}}}}')["data"]["course"]["enrollmentsConnection"]["nodes"]
        student_series = {}
        # Set up dictionary for students
        for student in query_result:
            if (student["type"] == "StudentEnrollment"):
                if (student["user"]["_id"] not in student_series):
                    student_series[student["user"]["_id"]] = {"name": student["user"]["name"], "email": student["user"]["email"]}
                    student_series[student["user"]["_id"]]["first_name"] = student["user"]["name"][0:student["user"]["name"].find(" ")]
                    student_series[student["user"]["_id"]]["last_name"] = student["user"]["name"][student["user"]["name"].rfind(" ")+1:]
                    student_series[student["user"]["_id"]]["sections"] = []
                student_series[student["user"]["_id"]]["sections"].append(student["section"]["_id"])
        
        # Identify which group the student is in
        groups = self.groups()
        def find_group(student, groups):
            for group in groups:
                for member in groups[group]["users"]:
                    if (student == member):
                        return(group)
            return("0")
        for student in student_series:
            group = find_group(student, groups)
            if (group != "0"):
                student_series[student]["group_id"] = group
                student_series[student]["group_name"] = groups[group]["name"]
            
        # Find the students' peers (if they are in a group)
        for student in student_series:
            if ("group_id" in student_series[student]):
                student_series[student]["peers"] = []
                for member in groups[student_series[student]["group_id"]]["users"]:
                    if (member != student):
                        student_series[student]["peers"].append(member)
                
        self.cache["students_" + str(self.course_id) + "_" + str(self.groupset_id)] = student_series                
        return(student_series)

    def teachers_comprehensive(self):
        if "teachers_" + str(self.course_id) in self.cache:
            return(self.cache["teachers_" + str(self.course_id)])
        query_result = self.query('query MyQuery {course(id: "' + str(self.course_id) + '") {enrollmentsConnection {nodes {type\nuser {_id\nname\nemail}\nsection {_id}}}}}')["data"]["course"]["enrollmentsConnection"]["nodes"]
        teacher_series = {}
        # Set up dictionary for teachers
        for teacher in query_result:
            if (teacher["type"] == "TeacherEnrollment"):
                if (teacher["user"]["_id"] not in teacher_series):
                    teacher_series[teacher["user"]["_id"]] = {"name": teacher["user"]["name"], "email": teacher["user"]["email"]}
                    teacher_series[teacher["user"]["_id"]]["first_name"] = teacher["user"]["name"][0:teacher["user"]["name"].find(" ")]
                    teacher_series[teacher["user"]["_id"]]["last_name"] = teacher["user"]["name"][teacher["user"]["name"].rfind(" ")+1:]
                    teacher_series[teacher["user"]["_id"]]["sections"] = []
                teacher_series[teacher["user"]["_id"]]["sections"].append(teacher["section"]["_id"])
                
        self.cache["teachers_" + str(self.course_id) + "_" + str(self.groupset_id)] = teacher_series                
        return(teacher_series)
        
        
#%% Tool Tips
class ToolTip(object):
# From: https://stackoverflow.com/questions/20399243/display-message-when-hovering-over-something-with-mouse-cursor-in-python
    
    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 57
        y = y + cy + self.widget.winfo_rooty() +27
        self.tipwindow = tw = tkinter.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tkinter.Label(tw, text=self.text, justify=tkinter.LEFT,
                      background="#ffffe0", relief=tkinter.SOLID, borderwidth=1,
                      font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def CreateToolTip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)

#%% Settings
    
class Settings:
    def __init__(self, window):
        # Again, try to load the configuration file and overwrite defaults
        try:
            f = open("config.txt", "r")
            user_config = json.loads(f.read())
            f.close()
            for item in user_config: config[item] = user_config[item]
        except:
            pass
        
        self.window = window
        window.title("Settings")
        window.resizable(0, 0)
        window.grab_set()
        
        tkinter.Label(window, text = "\nPlease enter the URL for the Canvas API:", fg = "black").grid(row = 0, column = 1, sticky = "W")
        tkinter.Label(window, text = "API URL:", fg = "black").grid(row = 1, column = 0, sticky = "E")
        self.settings_URL = tkinter.Entry(window, width = 50)
        self.settings_URL.grid(row = 1, column = 1, sticky="W", padx = 5)
        tkinter.Label(window, text = "\nPlease enter your API Key from Canvas:", fg = "black").grid(row = 2, column = 1, sticky = "W")
        tkinter.Label(window, text = "Token:", fg = "black").grid(row = 3, column = 0, sticky = "E")
        self.settings_Token = tkinter.Entry(window, width = 80)
        self.settings_Token.grid(row = 3, column = 1, sticky="W", padx = 5)
        
        #tkinter.Label(window, text = "\nPlease enter the student email suffix:", fg = "black").grid(row = 4, column = 1, sticky = "W")
        #tkinter.Label(window, text = "Suffix:", fg = "black").grid(row = 5, column = 0, sticky = "E")
        #self.settings_suffix = tkinter.Entry(window, width = 50)
        #self.settings_suffix.grid(row = 5, column = 1, sticky="W", padx = 5)
        
        tkinter.Label(window, text = "You will only need to enter these settings once.", fg = "black").grid(row = 6, columnspan = 2)
        self.settings_openfolder = tkinter.Button(window, text = "Open App Folder", fg = "black", command = self.open_folder).grid(row = 7, columnspan = 2, pady = 5)
        self.settings_default = tkinter.Button(window, text = "Restore Default Settings", fg = "black", command = self.restore_defaults).grid(row = 8, columnspan = 2, pady = 5)
        self.settings_save = tkinter.Button(window, text = "Save Settings", fg = "black", command = self.save_settings).grid(row = 9, columnspan = 2, pady = 5)
        self.update_fields()

    def open_folder(self):
        subprocess.Popen(['explorer', os.path.abspath(os.path.dirname(sys.argv[0]))])
        self.window.destroy()

    def update_fields(self):
        self.settings_URL.delete(0, tkinter.END)
        self.settings_URL.insert(0, config["API_URL"])
        self.settings_Token.delete(0, tkinter.END)
        self.settings_Token.insert(0, config["API_TOKEN"])
        #self.settings_suffix.delete(0, tkinter.END)
        #self.settings_suffix.insert(0, config["EmailSuffix"])
        
    def restore_defaults(self):
        global config
        config = {}
        for item in defaults: config[item] = defaults[item]
        self.update_fields()
        
    def save_settings(self):
        config["API_URL"] = cleanupURL(self.settings_URL.get())
        config["API_TOKEN"] = self.settings_Token.get().strip()
        save_config()
        self.window.destroy()

def Settings_call():
    settings_win = tkinter.Toplevel()
    set_win = Settings(settings_win)

#%% Export Groups
class ExportGroups():
    def __init__(self, master, preset = ""):
        self.preset = preset
        self.master = master
        self.exportgroups = tkinter.Frame(self.master)
        self.exportgroups.pack()
        
        if preset=="classlist":
            self.master.title("Download Class List")
            tkinter.Label(self.exportgroups, text = "\nDownload list of students and teachers", fg = "black").grid(row = 0, column = 0, columnspan = 2)            
        else:
            self.master.title("Create Contact List")
            tkinter.Label(self.exportgroups, text = "\nDownload student groups and group membership", fg = "black").grid(row = 0, column = 0, columnspan = 2)            

        self.master.resizable(0, 0)
        self.master.grab_set()

        tkinter.Label(self.exportgroups, text = "Unit:", fg = "black").grid(row = 1, column = 0, sticky = "E", pady = 5)           
        self.exportgroups_unit = ttk.Combobox(self.exportgroups, width = 60, values = [], state="readonly")
        self.exportgroups_unit.bind("<<ComboboxSelected>>", self.group_sets)
        self.exportgroups_unit.grid(row = 1, column = 1, sticky = "W", padx = 5)

        tkinter.Label(self.exportgroups, text = "Group Set:", fg = "black").grid(row = 2, column = 0, sticky = "E", pady = 5)
        self.exportgroups_groupset = ttk.Combobox(self.exportgroups, width = 60, state="readonly")
        self.exportgroups_groupset.grid(row = 2, column = 1, sticky = "W", padx = 5)

        #tkinter.Label(self.exportgroups, text = "Confirmation:", fg = "black").grid(row = 3, column = 0, sticky = "E", pady = 5)
        #self.contactgroups = tkinter.BooleanVar()
        #self.contactgroups.set(False)
        #self.exportgroups_contact = tkinter.Checkbutton(self.exportgroups, variable = self.contactgroups, text = "Automatically contact students via Canvas to confirm group membership")
        #self.exportgroups_contact.grid(row = 3, column = 1, sticky = "W")
        
        self.exportgroups_export = tkinter.Button(self.exportgroups, text = "Start Download", fg = "black", command = self.begin_export)
        self.exportgroups_export.grid(row = 4, columnspan = 2, pady = 5)

        self.group_sets_id = []
        self.group_sets_name = []          
        
        self.exportgroups_unit["values"] = tuple(session["course.names"])
        self.exportgroups_unit.current(0)
        self.group_sets()
        
        self.statusbar = tkinter.Label(self.exportgroups, text="", bd=1, relief=tkinter.SUNKEN, anchor=tkinter.W)
        self.statusbar.grid(row = 5, columnspan = 2, sticky = "ew")
        
        self.status("")
    
    def group_sets(self, event_type = ""):
        self.group_sets_id = []
        self.group_sets_name = []
        GQL.course_id = session["course.ids"][self.exportgroups_unit.current()]
        for group_set in GQL.group_sets():
            self.group_sets_id.append(group_set["_id"])
            self.group_sets_name.append(group_set["name"])
        if len(self.group_sets_id) > 0:
            self.exportgroups_groupset["values"] = tuple(self.group_sets_name)
            self.exportgroups_groupset.current(0)
        else:
            self.exportgroups_groupset["values"] = ()
            self.exportgroups_groupset.set('')
    
    def status(self, text): 
        self.statusbar["text"] = text
        self.statusbar.update_idletasks()
    
    def begin_export(self, event_type = ""):
        foldername = ""
        foldername = tkinter.filedialog.askdirectory(title = "Choose Folder to Save Files...")                
        GQL.course_id = session["course.ids"][self.exportgroups_unit.current()]
        if self.exportgroups_groupset.current() >= 0:
            GQL.groupset_id = self.group_sets_id[self.exportgroups_groupset.current()]
            session["group_set"] = self.group_sets_id[self.exportgroups_groupset.current()]    
        session["course.id"] = session["course.ids"][self.exportgroups_unit.current()]
        session["course.name"] = session["course.names"][self.exportgroups_unit.current()]
        print("Accessing " + session["course.name"] + "...")        
        students = GQL.students_comprehensive()
        print("Found {} students...".format(len(students)))

        if self.preset=="classlist":
            sections = GQL.sections()
            teachers = GQL.teachers_comprehensive()
            print("Found {} teachers...".format(len(teachers)))
            
            student_list = {"id": [],
                         "name":[],
                         "first_name":[],
                         "last_name":[],
                         "email":[],
                         "group_name":[]}
            
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
                else:
                    student_list["group_name"].append("")
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
            
            student_list = pd.DataFrame(student_list)
            teacher_list = pd.DataFrame(teacher_list)
            
            with pd.ExcelWriter(os.path.join(foldername, session["course.names"][self.exportgroups_unit.current()] + ".xlsx")) as writer:
                student_list.to_excel(writer, sheet_name='Student List', index = False)
                teacher_list.to_excel(writer, sheet_name='Teacher List', index = False)
            
            messagebox.showinfo("Download complete.", "Student list saved as \"{}.xlsx\".".format(session["course.names"][self.exportgroups_unit.current()]))
            
        else:
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
            
            print("Processing student details...")
            for student in students:
                student_list["ExternalDataReference"].append(student)
                student_list["Student_name"].append(students[student]["name"])
                student_list["FirstName"].append(students[student]["first_name"])
                student_list["LastName"].append(students[student]["last_name"])
                student_list["Email"].append(students[student]["email"])
                if ("group_id" in students[student]):
                    print("> {}: {}".format(students[student]["name"], students[student]["group_name"]))
                    student_list["Group_id"].append(students[student]["group_id"])            
                    student_list["Group_name"].append(students[student]["group_name"])
                    for j, peer in enumerate(students[student]["peers"],1):
                        student_list["Peer" + str(j) + "_id"].append(peer)
                        student_list["Peer" + str(j) + "_name"].append(students[peer]["name"])
                else:
                    print("> {}: No group found".format(students[student]["name"]))
                    student_list["Group_id"].append(np.nan)            
                    student_list["Group_name"].append(np.nan)
                    j = 0
                for k in range(j+1, 10):
                    student_list["Peer" + str(k) + "_id"].append(np.nan)
                    student_list["Peer" + str(k) + "_name"].append(np.nan)
                            
            # Write CSV of group members
            print("Creating Contacts file to be uploaded to Qualtrics")
            student_list = pd.DataFrame(student_list)
            student_list.loc[pd.isnull(student_list["Group_id"]) == False].to_csv(os.path.join(foldername,"Contacts.csv"), index = False)
            
            # Write CSV of students lacking groups
            print("Creating list of students without groups...")
            student_list.loc[pd.isnull(student_list["Group_id"]), ["Email", "ExternalDataReference", "FirstName", "Student_name"]].to_csv(os.path.join(foldername,"NoGroups.csv"), index = False)
            messagebox.showinfo("Download complete.", "Qualtrics Contact file saved as \"Contacts.csv\" and a list of students without groups has been saved in \"NoGroups.csv\".")
        
        self.master.destroy()

def ExportGroups_call():
    exportgroups = tkinter.Toplevel()
    exp_grp = ExportGroups(exportgroups)

def ExportGroups_classlist():
    exportgroups = tkinter.Toplevel()
    exp_grp = ExportGroups(exportgroups, preset="classlist")

#%% Peer Mark section
class PeerMark():
   def __init__(self, master):
       self.master = master
       self.master.title("Calculate Peer Marks")
       self.master.resizable(0, 0)
       
       self.peermark = tkinter.Frame(self.master)
       self.peermark.pack()
       
       #self.sentinel = None
    
       self.UploadPeerMark = tkinter.BooleanVar()
       self.UploadPeerMark.set(True)
    
       tkinter.Label(self.peermark, text = "Unit:", fg = "black").grid(row = 0, column = 0, sticky = "E", pady = 5)
       self.peermark_unit = ttk.Combobox(self.peermark, width = 60, state="readonly")
       self.peermark_unit.bind("<<ComboboxSelected>>", self.group_assns)
       self.peermark_unit.grid(row = 0, column = 1, sticky = "W", padx = 5)
    
       tkinter.Label(self.peermark, text = "Team assessment task:", fg = "black").grid(row = 1, column = 0, sticky = "E", pady = 5)
       self.peermark_assn = ttk.Combobox(self.peermark, width = 60, state="readonly")
       self.peermark_assn.grid(row = 1, column = 1, sticky = "W", padx = 5)
    
       tkinter.Label(self.peermark, text = "Qualtrics file:", fg = "black").grid(row = 2, column = 0, sticky = "E", pady = 5)
       self.peermark_peerdata = tkinter.Button(self.peermark, text = "Select...", fg = "black", command = self.select_file)
       self.peermark_peerdata.grid(row = 2, column = 1, pady = 5, sticky = "W")
       self.datafile = tkinter.Label(self.peermark, text = "", fg = "black")
       self.datafile.grid(row = 3, column = 1, pady = 5, sticky = "W")
    
       tkinter.Label(self.peermark, text = "Peer evaluation due date:", fg = "black").grid(row = 4, column = 0, sticky = "W")
       self.peermark_duedate = DateEntry(self.peermark, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern = "yyyy-mm-dd")
       self.peermark_duedate.grid(row = 4, column = 1, sticky = "W")
    
       self.peermark_customise = tkinter.Button(self.peermark, text = "Customise policies", fg = "black", command = Policies_call).grid(row = 5, column = 0, padx = 5, pady = 5, sticky = "W")
    
       self.peermark_calculate = tkinter.Button(self.peermark, text = "Calculate marks", fg = "black", command = self.start_calculate, state = 'disabled')
       self.peermark_calculate.grid(row = 5, column = 1, pady = 5)
    
       self.peermark_exportdata = ttk.Checkbutton(self.peermark, text = "Export peer data", variable = self.UploadPeerMark)
       self.peermark_exportdata.grid(row = 5, column = 2, sticky = "W")
       
       self.peermark_unit["values"] = tuple(session["course.names"])
       self.peermark_unit.current(0)
       self.group_assns()
        
       self.status_text = tkinter.StringVar()
       self.statusbar = tkinter.Label(self.peermark, bd=1, relief=tkinter.SUNKEN, anchor=tkinter.W, textvariable = self.status_text)
       self.statusbar.grid(row = 6, columnspan = 3, sticky = "ew")       
       
       self.master.bind("<FocusIn>", self.check_status)

   def select_file(self):
       session["DataFile"] = filedialog.askopenfilename(initialdir = ".",title = "Select file",filetypes = (("Compatible File Types",("*.csv","*.tsv","*.xlsx")),("All Files","*.*")))
       if len(session["DataFile"]) > 0:
           self.peermark_calculate["state"] = 'normal'
       else:
           self.peermark_calculate["state"] = 'disabled'
       self.datafile["text"] = session["DataFile"]

   def group_assns(self, event_type = ""):
        self.assn_id = []
        self.assn_name = []
        course = canvas.get_course(session["course.ids"][self.peermark_unit.current()])
        assns = course.get_assignments()
        for assn in assns:
            self.assn_id.append(assn.id)
            self.assn_name.append(assn.name)
        self.peermark_assn["values"] = tuple(self.assn_name)
        self.peermark_assn.current(0)

   def check_status(self, event_type):
       self.master.grab_set()    
       if config["ScoreType"] == 1:
           self.peermark_assn["state"] = "disabled"
       else:
           self.peermark_assn["state"] = "readonly"
    
   def status(self, stat_text):
       self.status_text.set(stat_text)
       time.sleep(0.01)
       
   def start_calculate(self):
       self.status("Loading, please be patient...")
       session["DueDate"] = self.peermark_duedate.get_date()
       session["GroupAssignment_ID"] = self.assn_id[self.peermark_assn.current()]
       session["SaveData"] = self.UploadPeerMark.get()
       course = canvas.get_course(session["course.ids"][self.peermark_unit.current()])
       calculate_marks(course, self)

def PeerMark_call():
    global pm
    pm = tkinter.Toplevel()
    window = PeerMark(pm)

#%% Calculate marks   
def calculate_marks(course, pm): 
   if session["SaveData"]:
       filename = ""
       filename = tkinter.filedialog.asksaveasfilename(initialfile="SaveData.xlsx", defaultextension=".xlsx", title = "Save Exported Data As...", filetypes = (("Excel files","*.xlsx"),("all files","*.*")))
    
   print("Uploading peer marks... please be patient.")
   #messagebox.showinfo("Uploading Peer Marks...", "This process may take a long time. Please do not close down the app, even if it appears to hang.")
   global qualtrics, orig_data, tw_list, tw_names, other_list, ratio_dict, adj_dict

   # Set up data frame
   orig_data = {"Student":[], "ID":[], "SIS User ID":[], "SIS Login ID":[], "orig_grade":[]}
   other_list = [] # Flush the "other" list.
   
   # Load the Qualtrics file
   print("Loading the Qualtrics file...")
   try:
       if session["DataFile"][-3:].upper()=="CSV":
           qualtrics = pd.read_csv(session["DataFile"])
       elif session["DataFile"][-3:].upper() in ["TSV","TXT"]:
           qualtrics = pd.read_table(session["DataFile"], encoding='utf-16')
       elif session["DataFile"][-3:].upper() == "XLS" or session["DataFile"][-4:].upper() in ["XLSX", "XLSM"]:
           qualtrics = pd.read_excel(session["DataFile"])
       else:
           messagebox.showinfo("Error", "Unable to import data from this file type.")
           pm.master.destroy()
   except:
       messagebox.showinfo("Error", "Unable to parse data from this file.")
       pm.master.destroy()
   
   # Create new assignment
   print("Creating new assignment for Peer Mark")
   new_assignment = course.create_assignment({'name': config["PeerMark_Name"],
                                              'notify_of_update': False,
                                              'points_possible': config["RescaleTo"],
                                              'description': 'This assignment is your Peer Mark for the team project.',
                                              'published': True,
                                              'due_at': datetime.strftime(session["DueDate"], '%Y-%m-%dT23:59:59')})
     
   print("New assignment created: " + new_assignment.name + " (" + str(new_assignment.id) + ")")
   
   # Set the assignment to manually reveal grades to students (i.e., hide the grades until released), using GraphQL
   query = {'access_token': config["API_TOKEN"],
            'query': 'mutation MyMutation {\nsetAssignmentPostPolicy(input: {assignmentId: ' + str(new_assignment.id) + ', postManually: true}) {\npostPolicy {\npostManually\n}\n}}'
            }
   url = "{}api/graphql".format(config["API_URL"])
   requests.post(url, data = query)

   # Get sections
   session["sections"] = []
   for section in course.get_sections():
       session["sections"].append(section.name)
       orig_data[section.name] = []

   # Get assignment details    
   if (config["ScoreType"]==1):
       # Use this method if we are only calculating totals. Download all users, but not submission score.
       print("Downloading list of students...")
       for user in course.get_users(enrollment_type=['student']):
           orig_data["Student"].append(user.name)
           orig_data["ID"].append(user.id)
           orig_data["SIS User ID"].append(user.sis_user_id)
           orig_data["SIS Login ID"].append(user.sis_user_id)
           orig_data["orig_grade"].append(np.nan)
           session["points_possible"] = config["PointsPossible"]
           for section in session["sections"]: orig_data[section].append(False)
           print("> " + user.name)
           #self.peermark.after(200, self.status, "--" + user.name)
   elif (config["ScoreType"]==2):
       # Download student details and grades and submission score
       assignment = course.get_assignment(session["GroupAssignment_ID"])    
       print("Downloading group marks for " + assignment.name)
       session["points_possible"] = assignment.points_possible
       submissions = assignment.get_submissions()
       for submission in submissions:
           user = course.get_user(submission.user_id)
           orig_data["Student"].append(user.name)
           orig_data["ID"].append(submission.user_id)
           orig_data["SIS User ID"].append(user.sis_user_id)
           orig_data["SIS Login ID"].append(user.login_id)
           orig_data["orig_grade"].append(submission.score)
           for section in session["sections"]: orig_data[section].append(False)
           print("> " + user.name)
           #self.peermark.after(200, self.status, "--" + user.name)

   # Convert to pandas data frame
   orig_data = pd.DataFrame(data = orig_data)
   
   # Populate information about sections
   print("Downloading class membership information...")
   for section in course.get_sections():
       print("> " + section.name)
       for student in section.get_enrollments():
           if (orig_data["ID"]==student.user_id).any() == True:
               y = orig_data[orig_data["ID"] == student.user_id].index[0]
               orig_data.loc[y, section.name] = True
    
   # Clean up Qualtrics
   print("Cleaning up Qualtrics data file...")
   rep_list = ["Please rate each team member in relation to: ", " - Myself"]
   for a in range(1,10): rep_list.append(" - [Field-Peer" + str(a) + "_name]")
   
   # Determine the teamwork items in the Qualtrics file
   for b,a in enumerate(qualtrics.columns):
       if (a[0:8].upper() == "TEAMWORK"):
           a.replace("&shy", "\u00AD") # Replace soft hyphen HTML with actual soft hyphen
           if (a.find("_") != -1):
               tw_list.append(a[0:a.rfind("_")])
               if qualtrics.iloc[0, b].count("\u00AD") == 2:
                   str_loc = find_all(qualtrics.iloc[0, b], "\u00AD")
                   tw_names.append(qualtrics.iloc[0, b][str_loc[0]+1:str_loc[1]])
               elif qualtrics.iloc[0, b].count("*") == 2:
                   str_loc = find_all(qualtrics.iloc[0, b], "*")
                   tw_names.append(qualtrics.iloc[0, b][str_loc[0]+1:str_loc[1]])
               else:
                   tw_names.append(replace_multiple(qualtrics.iloc[0, b], rep_list, ""))
       elif a[0:8].upper() != "FEEDBACK" and replace_multiple(a.upper(), range(0,10)) not in ["PEER_NAME", "PEER_ID"] and a not in ["Student_name", "Student_id"]:
           other_list.append(a)
   tw_list = unique(tw_list)
   tw_names = unique(tw_names)
   
   # Drop redundant rows
   qualtrics = qualtrics.drop(qualtrics.index[[0,1]]) # Drop rows 2 and 3
   qualtrics.dropna(subset = ["Student_id"], axis = 0, inplace = True) # Drop rows where there is no Student ID
   qualtrics = qualtrics.reset_index(drop = True) # Reset the index
   
   # Convert the "Finished" column to Boolean values
   qualtrics["Finished"] = qualtrics["Finished"]=="True"
      
   # Delete duplicate rows
   dups_list = qualtrics["Student_id"].value_counts()
   dups_list = dups_list.loc[dups_list > 1]
   for ID in dups_list.keys():
       # If a person has both partial and complete entries, drop the partial entries
       if len(qualtrics.loc[(qualtrics["Student_id"]==ID) & (qualtrics["Finished"] == True)]) > 0 and len(qualtrics.loc[(qualtrics["Student_id"]==ID) & (qualtrics["Finished"] == False)]) > 0:
           drop_index = qualtrics[(qualtrics["Student_id"]==ID) & (qualtrics["Finished"] == False)].index
           qualtrics.drop(drop_index, axis=0, inplace = True)
           qualtrics.reset_index(drop = True, inplace = True) # Reset the index       
       drop_index = qualtrics[qualtrics["Student_id"]==ID].index
       if len(drop_index) > 1:
           qualtrics.drop(drop_index[:-1], axis = 0, inplace = True)
           qualtrics.reset_index(drop = True, inplace = True) # Reset the index   

   # Count the number of peers for each person; also turn User IDs into integers
   n_peers = []
   for y in range(0,qualtrics.shape[0]):
       qualtrics.loc[y, 'Student_id'] = int(float(qualtrics.loc[y, "Student_id"]))
       n_peers.append(0)
       for peer in range(1,10):
           if not(pd.isna(qualtrics.loc[y, 'Peer' + str(peer) + "_id"])):
               n_peers[y] += 1
               qualtrics.loc[y, 'Peer' + str(peer) + "_id"] = int(float(qualtrics.loc[y, 'Peer' + str(peer) + "_id"]))
   qualtrics['NumPeers'] = n_peers
   
   # Eliminate response if no peer ratings were given at all
   for ID in qualtrics["Student_id"]:
       a = ratings_given(ID)
       if (pd.isnull(a) == False).sum() == 0:    
           y = qualtrics[qualtrics["Student_id"] == ID].index
           qualtrics.drop(y, axis = 0, inplace = True)
           qualtrics.reset_index(drop = True, inplace = True)
   
   # Replace strings with numbers
   for item in tw_list:
       for peer in range(1,11):
           for y in range(0,qualtrics.shape[0]):
               if (peer-1 > qualtrics.loc[y, 'NumPeers']): # Put in "NaN" if we have run out of peers
                   qualtrics.loc[y, item + '_' + str(peer)] = np.NaN
               else: # Replace strings with numbers
                   if (type(qualtrics.loc[y, item + '_' + str(peer)]) is str):
                       qualtrics.loc[y, item + '_' + str(peer)] = int(float(qualtrics.loc[y, item + '_' + str(peer)][0]))
                   qualtrics.loc[y, item + '_' + str(peer)] += -1 # Make the lowest number 0       

   # Upload marks to Canvas
   print("Uploading marks to Canvas...")
   iter = 0
   tstamp = datetime.now().timestamp()
   while (iter==0 and (datetime.now().timestamp() - tstamp < 10)): # Keep looping until at least one iteration has been completed. Timeout after 10 seconds.
       for submission in new_assignment.get_submissions():
           iter = iter + 1
           ID = submission.user_id
           try:
               print("> " + orig_data.loc[orig_data["ID"]==ID, "Student"].values[0] + ": " + str(round(adj_score(ID),2)))
           except:
               pass
           if type(adj_score(ID)) is not str:
               if ID in qualtrics["Student_id"].values:
                   submission.edit(submission={'posted_grade': adj_score(ID), 'posted_at': "null", 'submitted_at': datetime.strftime(submission_date(ID), '%Y-%m-%dT23:59:59')})
               else:
                   submission.edit(submission={'posted_grade': adj_score(ID), 'posted_at': "null"})
           submission.edit(comment={'text_comment': comments(ID)})
       
   new_assignment.edit(assignment = {"published": config["publish_assignment"]})
   
   if iter==0:
       messagebox.showinfo("Error", "Unable to upload marks to Canvas. Please try again.")
   
    # Save the data; add all of the penalties information
   if session["SaveData"] and filename != "":
        penalties = {}
        penalties["PartialComplete"] = []
        penalties["SelfPerfect"] = []
        penalties["PeersAllZero"] = []
        for ID in qualtrics["Student_id"]:
            if PartialComplete(ID): penalties["PartialComplete"].append(1)
            else: penalties["PartialComplete"].append(0)
            if SelfPerfect(ID): penalties["SelfPerfect"].append(1)
            else: penalties["SelfPerfect"].append(0)
            if PeersAllZero(ID): penalties["PeersAllZero"].append(1)
            else: penalties["PeersAllZero"].append(0)
        qualtrics["PartialComplete"] = penalties["PartialComplete"]
        qualtrics["SelfPerfect"] = penalties["SelfPerfect"]
        qualtrics["PeersAllZero"] = penalties["PeersAllZero"]
        other_list.extend(["PartialComplete","SelfPerfect","PeersAllZero"])
    
    # Create dataset (long form)
        print("Saving data in long format...")
        long_data = {}
        for section in session["sections"]: long_data[section] = []
        long_data["RaterName"] = []
        long_data["ReceiverName"] = []
        long_data["RaterID"] = []
        long_data["ReceiverID"] =  []
        for item in tw_list: long_data[item] = []
        long_data["FeedbackShared"] = []
        long_data["FeedbackInstructor"] = []
        if (config["ScoreType"] == 2):
            long_data["OriginalScore"] = []
            long_data["RatioAdjust"] = []
        long_data["PeerMark"] = []
        long_data["NonComplete"] = []
        for ID in orig_data["ID"]:
            y1 = orig_data[orig_data["ID"]==ID].index[0]
            if (qualtrics["Student_id"] == ID).any():
                long_data["NonComplete"].extend([0] * len(rater_peerID(ID)))
                y = qualtrics[qualtrics["Student_id"] == ID].index[0]
                #for a, item in enumerate(other_list):
                #    long_data[item].extend([qualtrics.loc[y, item]] * len(rater_peerID(ID)))
            else:
                long_data["NonComplete"].extend([1] * len(rater_peerID(ID)))
                #for a, item in enumerate(other_list):
                #    long_data[item].extend([np.nan] * len(rater_peerID(ID)))
            for a, rater in enumerate(rater_peerID(ID)):
                y2 = orig_data[orig_data["ID"]==rater].index[0]
                long_data["ReceiverName"].append(orig_data.loc[y1, "Student"])
                long_data["RaterName"].append(orig_data.loc[y2, "Student"])
                long_data["ReceiverID"].append(ID)
                long_data["RaterID"].append(rater)
                for section in session["sections"]: long_data[section].append(orig_data.loc[y1, section])
                for b, item in enumerate(tw_list):
                    long_data[item].append(ratings_received(ID, rater)[0][b])
                long_data["FeedbackShared"].append(feedback(ID)[0][a])
                long_data["FeedbackInstructor"].append(feedback(ID)[1][a])
                if (config["ScoreType"] == 2):
                    long_data["OriginalScore"].append(float(orig_data[orig_data["ID"]==ID]["orig_grade"]))
                    long_data["RatioAdjust"].append(ratee_ratio(ID)[a])
            long_data["PeerMark"].extend([adj_score(ID)] * len(rater_peerID(ID)))
        long_data = pd.DataFrame(long_data)
    
    # Create dataset (wide form) if requested
        print("Saving data in wide format...")
        wide_data = {}
        wide_data["StudentID"] = []
        wide_data["StudentName"] = []
        for section in session["sections"]: wide_data[section] = []
        if (config["ScoreType"] == 2):
            wide_data["OriginalScore"] = []
        wide_data["PeerMark"] = []
        wide_data["NonComplete"] = []
        for item in other_list: wide_data[item] = []

        for ID in orig_data["ID"]:
            y1 = orig_data[orig_data["ID"]==ID].index[0]
            if (qualtrics["Student_id"] == ID).any():
                wide_data["NonComplete"].append(0)
                y = qualtrics[qualtrics["Student_id"] == ID].index[0]
                for a, item in enumerate(other_list):
                    wide_data[item].append(qualtrics.loc[y, item])
            else:
                wide_data["NonComplete"].append(1)
                for a, item in enumerate(other_list):
                    wide_data[item].append(np.nan)
            for section in session["sections"]: wide_data[section].append(orig_data.loc[y1, section])
            wide_data["StudentName"].append(orig_data.loc[y1, "Student"])
            wide_data["StudentID"].append(ID)
            if (config["ScoreType"] == 2): wide_data["OriginalScore"].append(float(orig_data[orig_data["ID"]==ID]["orig_grade"]))
            wide_data["PeerMark"].append(adj_score(ID))
            
        global temp_data
        temp_data = wide_data
        wide_data = pd.DataFrame(wide_data)
        
        with pd.ExcelWriter(filename) as writer:
            wide_data.to_excel(writer, sheet_name='Student Data', index = False)
            long_data.to_excel(writer, sheet_name='Rating Data', index = False)
        
        messagebox.showinfo("Upload Complete", 'The peer marks have been successfully uploaded to Canvas. An assignment, "Peer Mark", has been created. When it is ready to be released to students, you should publish the assignment and unmute it from the Gradebook. Peer rater data saved to "' + filename + '".')
   else:
        messagebox.showinfo("Upload Complete", 'The peer marks have been successfully uploaded to Canvas. An assignment, "Peer Mark", has been created. When it is ready to be released to students, you should publish the assignment and unmute it from the Gradebook.')
   pm.master.destroy()
   
#%% Policies
class Policies():
    def __init__(self, master):
        self.master = master
        self.master.title("Calculate Peer Marks")
        self.master.resizable(0, 0)
        self.master.grab_set()
        
        self.policies = tkinter.Frame(self.master)
        self.policies.pack()
        self.validation = self.policies.register(only_numbers)
        

                
        labelframe1 = tkinter.LabelFrame(self.policies, text = "Peer Mark Calculation Policy", fg = "black")
        labelframe1.grid(row = 0, column = 0, pady = 5, padx = 5, sticky = "W")
        
        self.policies_lab1 = tkinter.Label(labelframe1, text = "Scoring method:", fg = "black").grid(row = 0, column = 0, sticky = "E", pady = 5)
        self.policies_scoring = ttk.Combobox(labelframe1, width = 40, values = ["Average rating received", "Adjusted score based on ratings received"], state="readonly")
        self.policies_scoring.grid(row = 0, column = 1, columnspan = 2, sticky = "W", padx = 5)
        self.policies_scoring.current(config["ScoreType"]-1)
        
        tkinter.Label(labelframe1, text = "Self-scoring policy:", fg = "black").grid(row = 1, column = 0, sticky = "E", pady = 5)
        self.policies_selfrat = ttk.Combobox(labelframe1, width = 40, values = ["Do not allow self-scoring", "Allow self-scoring", "Subsitute self-score with average rating given", "Cap self-score at average rating given"], state="readonly")
        self.policies_selfrat.grid(row = 1, column = 1, columnspan = 2, sticky = "W", padx = 5)
        self.policies_selfrat.current(config["SelfVote"])
        
        tkinter.Label(labelframe1, text = "Minimum peer responses:", fg = "black").grid(row = 2, column = 0, sticky = "E", pady = 5)
        self.policies_minimum = tkinter.Entry(labelframe1, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.policies_minimum.grid(row = 2, column = 1, padx = 5, sticky = "W")
        self.policies_minimum.insert(0, str(config["MinimumPeers"]))
        
        labelframe2 = tkinter.LabelFrame(self.policies, text = "Penalties Policy", fg = "black")
        labelframe2.grid(row = 1, column = 0, pady = 5, padx = 5, sticky = "W")
        tkinter.Label(labelframe2, text = "Apply % penalty", fg = "black").grid(row = 0, column = 2, sticky = "W")
        
        tkinter.Label(labelframe2, text = "Students do not complete the peer evaluation", fg = "black").grid(row = 1, rowspan = 2, column = 0, sticky = "NE")
        tkinter.Label(labelframe2, text = "\n", fg = "black").grid(row = 1, rowspan = 2, column = 1, sticky = "NE")
        self.penalty_noncomplete = tkinter.Entry(labelframe2, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.penalty_noncomplete.grid(row = 1, column = 2, sticky="W", padx = 5, pady = 5)
        self.penalty_noncomplete.insert(0, config["Penalty_NonComplete"])
        
        tkinter.Label(labelframe2, text = "Students submit the peer evaluation late", fg = "black").grid(row = 2, rowspan = 2, column = 0, sticky = "NE")
        tkinter.Label(labelframe2, text = "(per day)", fg = "black").grid(row = 2, column = 2, sticky = "E")
        self.penalty_late = tkinter.Entry(labelframe2, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.penalty_late.grid(row = 2, column = 2, sticky="W", padx = 5, pady = 5)
        self.penalty_late.insert(0, config["Penalty_PerDayLate"])
        
        tkinter.Label(labelframe2, text = "Students give themselves perfect scores", fg = "black").grid(row = 3, rowspan = 2, column = 0, sticky = "NE")
        self.perfect_score = tkinter.IntVar()
        ttk.Radiobutton(labelframe2, text = "Retain scores", variable = self.perfect_score, value = 0).grid(row = 3, column = 1, sticky = "W")
        ttk.Radiobutton(labelframe2, text = "Exclude scores", variable = self.perfect_score, value = 1).grid(row = 4, column = 1, sticky = "W")
        self.perfect_score.set(config["Exclude_SelfPerfect"])
        self.penalty_selfperfect = tkinter.Entry(labelframe2, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.penalty_selfperfect.grid(row = 3, column = 2, sticky="W", padx = 5)
        self.penalty_selfperfect.insert(0, config["Penalty_SelfPerfect"])
        
        tkinter.Label(labelframe2, text = "Students give all peers the bottom score", fg = "black").grid(row = 5, rowspan = 2, column = 0, sticky = "NE")
        self.bottom_score = tkinter.IntVar()
        ttk.Radiobutton(labelframe2, text = "Retain scores", variable = self.bottom_score, value = 0).grid(row = 5, column = 1, sticky = "W")
        ttk.Radiobutton(labelframe2, text = "Exclude scores", variable = self.bottom_score, value = 1).grid(row = 6, column = 1, sticky = "W")
        self.bottom_score.set(config["Exclude_PeersAllZero"])
        self.penalty_allzero = tkinter.Entry(labelframe2, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.penalty_allzero.grid(row = 5, column = 2, sticky="W", padx = 5)
        self.penalty_allzero.insert(0, config["Penalty_PeersAllZero"])
        
        labelframe3 = tkinter.LabelFrame(self.policies, text = "Additional Settings", fg = "black")
        labelframe3.grid(row = 2, column = 0, pady = 5, padx = 5, sticky = "W")

        tkinter.Label(labelframe3, text = "Peer mark scale ranges from 1 to:", fg = "black").grid(row = 0, column = 0, sticky = "NE")
        self.points_possible = tkinter.Entry(labelframe3, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.points_possible.grid(row = 0, column = 1, sticky="W", padx = 5)
        self.points_possible.insert(0, config["PointsPossible"])

        tkinter.Label(labelframe3, text = "Rescale peer mark to score out of:", fg = "black").grid(row = 1, column = 0, sticky = "NE")
        self.rescale_to = tkinter.Entry(labelframe3, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.rescale_to.grid(row = 1, column = 1, sticky="W", padx = 5)
        self.rescale_to.insert(0, config["RescaleTo"])
        
        tkinter.Label(labelframe3, text = "Assignment name:", fg = "black").grid(row = 2, column = 0, sticky = "E", pady = 5)
        self.pm_name = tkinter.Entry(labelframe3, width = 40)
        self.pm_name.grid(row = 2, column = 1, padx = 5, sticky = "W")
        self.pm_name.insert(0, str(config["PeerMark_Name"]))

        self.publish_assignment = tkinter.IntVar()
        self.publish_assignment.set(int(config["publish_assignment"]))
        tkinter.Label(labelframe3, text = "Publish Peer Mark column in Grade Book?", fg = "black").grid(row = 3, column = 0, sticky = "NE")
        ttk.Checkbutton(labelframe3, text = "Make the column visible but hide grades", variable = self.publish_assignment).grid(row = 3, column = 1, sticky = "W")

        self.weekdays_only = tkinter.IntVar()
        self.weekdays_only.set(int(config["weekdays_only"]))
        tkinter.Label(labelframe3, text = "Exclude weekends?", fg = "black").grid(row = 4, column = 0, sticky = "NE")
        ttk.Checkbutton(labelframe3, text = "Only count weekdays when calculating late penalty", variable = self.publish_assignment).grid(row = 4, column = 1, sticky = "W")
        
        self.save_policies = tkinter.IntVar()
        self.save_policies.set(1)
        tkinter.Label(labelframe3, text = "Retain policies?", fg = "black").grid(row = 5, column = 0, sticky = "NE")
        ttk.Checkbutton(labelframe3, text = "Save policies for future assessments", variable = self.save_policies).grid(row = 5, column = 1, sticky = "W")
        
        tkinter.Button(self.policies, text = "Continue", fg = "black", command = self.save_settings).grid(row = 13, column = 0, columnspan = 3, pady = 5)
        
        self.policies_scoring.bind("<<ComboboxSelected>>", self.change_assnname)
    
    def change_assnname(self, event_list = ""):
        self.pm_name.delete(0, 'end')
        if self.policies_scoring.current()==0:
            self.pm_name.insert(0, "Peer Mark (average mark received)") # Average mark received
        else:
            self.pm_name.insert(0, "Team Assignment (peer-rating adjusted)")
    
    def save_settings(self):
        config["ScoreType"] = int(self.policies_scoring.current()+1)
        config["SelfVote"] = int(self.policies_selfrat.current())
        config["MinimumPeers"] = int(self.policies_minimum.get())
        config["PeerMark_Name"] = str(self.pm_name.get().strip())
        config["Penalty_NonComplete"] = int(self.penalty_noncomplete.get())
        config["Penalty_PerDayLate"] = int(self.penalty_late.get())
        config["Penalty_SelfPerfect"] = int(self.penalty_selfperfect.get())
        config["Penalty_PeersAllZero"] = int(self.penalty_allzero.get())
        config["PointsPossible"] = int(self.points_possible.get())
        config["RescaleTo"] = int(self.rescale_to.get())
        config["publish_assignment"] = bool(self.publish_assignment.get())
        config["weekdays_only"] = bool(self.weekdays_only.get())
        
        if self.save_policies.get() == 1: save_config()
        self.master.destroy()
        
def Policies_call():
    policies = tkinter.Toplevel()
    pol = Policies(policies)

#%% Bulk mail
class BulkMail:
    def __init__(self, master, preset = "", subject = "", message = ""):
        self.preset = preset
        self.master = master
        self.master.title("Send Bulk Message via Canvas")
        self.master.resizable(0, 0)
        
        self.master.bind("<FocusIn>", self.check_focus)
        
        self.window = tkinter.Frame(self.master)
        self.window.pack()
        
        tkinter.Label(self.window, text = "Unit:", fg = "black").grid(row = 1, column = 0, sticky = "E", pady = 5)
        self.window_unit = ttk.Combobox(self.window, width = 60, state="readonly")
        self.window_unit.grid(row = 1, column = 1, sticky = "W", padx = 5)
        self.window_unit["values"] = tuple(session["course.names"])
        self.window_unit.current(0)

        self.sec_text = tkinter.StringVar()
        self.sec_text.set("Distribution list file")
        self.secondary_lab = tkinter.Label(self.window, textvariable = self.sec_text, fg = "black").grid(row = 3, column = 0, sticky = "E", pady = 5)        

        # Distribution list file
        self.window_distlist = tkinter.Button(self.window, text = "Select...", fg = "black", command = self.select_file)
        self.window_distlist.grid(row = 3, column = 1, pady = 5, padx = 5, sticky = "W")
        self.datafile = tkinter.Label(self.window, text = "", fg = "black")
        self.datafile.grid(row = 4, column = 1, pady = 5, sticky = "W")

        # Group sets
        self.window_groupset = ttk.Combobox(self.window, width = 60, state="readonly")
        self.group_sets(self)
        self.window_unit.bind("<<ComboboxSelected>>", self.group_sets)
        #self.groupset.grid(row = 3, column = 1, sticky = "W", padx = 5)

        tkinter.Label(self.window, text = "Send to (via):", fg = "black").grid(row = 0, column = 0, sticky = "E", pady = 5)
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
        
        tkinter.Label(self.window, text = "Message:", fg = "black").grid(row = 5, column = 0, columnspan = 2, pady = 5, padx = 5)

        tkinter.Label(self.window, text = "Subject:", fg = "black").grid(row = 6, column = 0, sticky = "E", pady = 5)
        self.subject_text = tkinter.StringVar()
        self.subject = tkinter.Entry(self.window, textvariable = self.subject_text)
        self.subject.grid(row = 6, column = 1, pady = 5, padx = 5, sticky = "WE")
        self.subject_text.set(subject)

        self.message = tkinter.Text(self.window, width = 60, height = 15, wrap = tkinter.WORD)
        self.message.grid(row = 7, column = 0, columnspan = 2, sticky = "W", pady = 5, padx = 5)
        self.message.bind("<KeyRelease>", self.check_status)
        self.message.insert(tkinter.END, message)
        
        self.button_field = tkinter.Button(self.window, text = "Fields...", fg = "black", command = self.field_picker)
        self.button_field.grid(row = 8, column = 0, columnspan = 2, sticky = "W", padx = 5)
        
        self.sendmessage = tkinter.Button(self.window, text = "Send Bulk Message", fg = "black", state = "disabled", command = self.MessageSend)
        self.sendmessage.grid(row = 8, column = 0, columnspan = 2, pady = 5)
       
        self.distlist = {}
        session["DistList"] = ""

        self.save_template = tkinter.IntVar()
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
        field_window = tkinter.Toplevel()
        fp = FieldPicker(field_window, self.distlist_keys())
    
    def insert_field(self, event_type = "", field = ""):
        self.message.insert(tkinter.INSERT, field)
    
    def check_status(self, event_type = ""):
        if (len(self.message.get(1.0, tkinter.END).strip()) < 1) or (self.window_sendto.current() in [4,5,6] and session["DistList"] == ""):
            self.sendmessage["state"] = "disabled"
        else:
            self.sendmessage["state"] = "normal"
        if (self.window_sendto.current() in [4,5,6] and session["DistList"] == ""):
            self.button_field["state"] = "disabled"
        else:
            self.button_field["state"] = "normal"
        
    def select_file(self):
       session["DistList"] = filedialog.askopenfilename(initialdir = ".", title = "Select file",filetypes = (("Compatible File Types",("*.csv")),("All Files","*.*")))
       self.datafile["text"] = session["DistList"]
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
        GQL.course_id = session["course.ids"][self.window_unit.current()]
        for group_set in GQL.group_sets():
            self.group_sets_id.append(group_set["_id"])
            self.group_sets_name.append(group_set["name"])
        if len(self.group_sets_id) > 0:
            self.window_groupset["values"] = tuple(self.group_sets_name)
            self.window_groupset.current(0)
        else:
            self.window_groupset["values"] = ()
            self.window_groupset.set('')            
        session["group_set"] = self.group_sets_id

    def create_distlist(self, event_type = ""):
        if self.window_sendto.current() in [1,2,3]: # Prepopulate the list of groups
            GQL.groupset_id = self.group_sets_id[self.window_groupset.current()]        
        if self.window_sendto.current() in [0,1,2]: # Download students from Canvas
            students = GQL.students_comprehensive()
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
                    self.distlist[student]["group_homepage"] = config["API_URL"] + "groups/" + str(students[student]["group_id"]) + "/"
        if self.window_sendto.current() == 2: # Drop students who have groups
            for student in students:
                if "group_id" not in students[student]:
                    self.distlist[student] = students[student]
        if self.window_sendto.current() == 3: # Get list of groups
            groups = GQL.groups()
            students = GQL.students_comprehensive()
            for group in groups:
                if len(groups[group]["users"]) > 0:
                    self.distlist[group] = groups[group]
                    self.distlist[group]["members_bulletlist"] = ""
                    self.distlist[group]["members_list"] = ""
                    self.distlist[group]["firstnames_list"] = ""
                    self.distlist[group]["group_homepage"] = config["API_URL"] + "groups/" + str(group) + "/"
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
            print(session["DistList"])
            if len(session["DistList"]) > 0:
                try:
                    df = pd.read_csv(session["DistList"])
                    students = {}
                    if any(a in df.columns for a in ["External Data Reference", "ExternalDataReference", "Student_id", "id", "ID"]):
                        #self.datafile["text"] = session["DistList"]
                        if "Student_id" in df.columns: session["ID_column"] = "Student_id"
                        if "External Data Reference" in df.columns: session["ID_column"] = "External Data Reference"
                        if "ExternalDataReference" in df.columns: session["ID_column"] = "ExternalDataReference"
                        if "id" in df.columns: session["ID_column"] = "id"
                        if "ID" in df.columns: session["ID_column"] = "ID"
                        for y, ID in enumerate(df[session["ID_column"]]):
                            students[str(ID)] = {}
                            for col_name in df.columns:
                                if col_name != session["ID_column"]:
                                    students[str(ID)][col_name] = df.loc[y, col_name]                        
                    else:
                        messagebox.showinfo("Error", "Cannot find user ID column in file. This column should be labelled 'id', 'Student_id' or 'External Data Reference'.")
                except:
                    messagebox.showinfo("Error", "Trouble parsing file.")
        if self.window_sendto.current() == 4:
            self.distlist = students
        if self.window_sendto.current() in [5,6] and "Status" not in df.columns:
            messagebox.showinfo("Error", "Cannot find 'Status' column in file used to indicate whether a student has completed the survey.")
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
            
        #print(self.distlist)
                
    def MessageSend(self):
        if self.preset!="":
            print("Saving message as template...")
            config["subject_" + self.preset] = self.subject_text.get()
            config["message_" + self.preset] = self.message.get(1.0, tkinter.END)
            save_config()
            
        print("Preparing distribution list...")                
        self.create_distlist(self)
        if len(self.distlist) == 0: 
            messagebox.showinfo("Error", "No students found.")
        else:
            print("Preparing messages...")
            messagelist = {}
            for case in self.distlist:
                messagelist[case] = {}
                message = self.message.get(1.0, tkinter.END)
                subject = self.subject_text.get()
                for field in self.distlist[case]:
                    subject = subject.replace("["+field+"]", str(self.distlist[case][field]))
                    message = message.replace("["+field+"]", str(self.distlist[case][field]))
                messagelist[case]["subject"] = subject
                messagelist[case]["message"] = message
            
            for k, case in enumerate(messagelist, 1):
                print("Sending message {} of {}".format(k, len(messagelist)))
                if self.window_sendto.current() == 3: # Send group message
                    bulk_message = canvas.create_conversation(recipients = ['group_' + str(case)],
                                                         body = messagelist[case]["message"],
                                                         subject = messagelist[case]["subject"],
                                                         group_conversation = True,
                                                         context_code = "course_" + str(GQL.course_id))            
                else: # Send individual message
                    bulk_message = canvas.create_conversation(recipients = [str(case)],
                                                              body = messagelist[case]["message"],
                                                              subject = messagelist[case]["subject"],
                                                              force_new = True,
                                                              context_code = "course_" + str(GQL.course_id))
            messagebox.showinfo("Finished", "All messages have been sent.")
        self.master.destroy()

def BulkMail_call(preset = ""):
    global bm
    window = tkinter.Toplevel()
    bm = BulkMail(window)
    
def BulkMail_confirm():
    global bm
    window = tkinter.Toplevel()
    bm = BulkMail(window,
                  preset = "confirm", 
                  subject = config["subject_confirm"],
                  message = config["message_confirm"])

def BulkMail_nogroup():
    global bm
    window = tkinter.Toplevel()
    bm = BulkMail(window,
                  preset = "nogroup", 
                  subject = config["subject_nogroup"],
                  message = config["message_nogroup"])
    
def BulkMail_invitation():
    global bm
    window = tkinter.Toplevel()
    bm = BulkMail(window,
                  preset = "invitation", 
                  subject = config["subject_invitation"],
                  message = config["message_invitation"])

def BulkMail_reminder():
    global bm
    window = tkinter.Toplevel()
    bm = BulkMail(window,
                  preset = "reminder", 
                  subject = config["subject_reminder"],
                  message = config["message_reminder"])

#%% Field picker
class FieldPicker:
    def __init__(self, master, fields = []):
        self.fields = fields
        self.master = master
        self.master.title("Fields")
        self.master.resizable(0,0)
        self.master.grab_set()
        
        self.window = tkinter.Frame(self.master)
        self.window.pack()
        
        self.field_list = tkinter.Listbox(self.window, width = 30, height = 20)
        self.field_list.pack(padx = 10, pady = 10)
        for field in fields:
            self.field_list.insert(tkinter.END, field)
            
        self.insert_field = tkinter.Button(self.window, text = "Insert Field", command = self.insert_text)
        self.insert_field.pack(padx = 10, pady = 10)
        
    def insert_text(self):
        self.master.grab_release()
        bm.insert_field(field = "[" + self.fields[self.field_list.curselection()[0]] + "]")
        self.master.destroy()  

#%% Qualtrics Template
def upload_template():
    subprocess.Popen(['explorer', os.path.abspath(os.path.dirname(sys.argv[0])) + "\\survey"])
    messagebox.showinfo("Create Survey", "To create the survey, upload the \"Peer_Evaluation_Qualtrics_Template.qsf\" template to Qualtrics. Read the \"Instructions.txt\" file for more information.")

#%% Main Menu

class MainMenu:
    def __init__(self, master):
        self.master = master
        self.master.title("PEER")
        #self.window.geometry("300x250") # size of the window width:- 500, height:- 375
        self.master.resizable(0, 0) # this prevents from resizing the window
        self.master.bind("<FocusIn>", self.check_status)

        self.window = tkinter.Frame(self.master)
        self.window.pack(side="top", pady = 0, padx = 50)

        # create a pulldown menu, and add it to the menu bar
        self.menubar = tkinter.Menu(self.window)
        self.master.config(menu = self.menubar)

        self.filemenu = tkinter.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="Settings", command = Settings_call)
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Exit", command = self.exit_program)
        self.menubar.add_cascade(label="File", menu=self.filemenu)              

        self.toolmenu = tkinter.Menu(self.menubar, tearoff = 0)
        self.toolmenu.add_command(label="Send Bulk Message", command = BulkMail_call)
        self.toolmenu.add_command(label="Download Class List", command = ExportGroups_classlist)
        self.menubar.add_cascade(label="Tools", menu=self.toolmenu)

        labelframe1 = tkinter.LabelFrame(self.window, text = "Preparation", fg = "black")
        labelframe1.pack(pady = 10)

        self.button_cgm = tkinter.Button(labelframe1, text = "Confirm group membership", fg = "black", command = BulkMail_confirm)
        self.button_cgm.pack(pady = 10, padx = 10)# 'text' is used to write the text on the Button
        CreateToolTip(self.button_cgm, text = 
                      'Contact students via the Canvas bulk mailer to confirm\n'+
                      'their membership in each group, as well as their peers.')

        self.button_ng = tkinter.Button(labelframe1, text = "Contact students\nwithout groups", fg = "black", command = BulkMail_nogroup)
        self.button_ng.pack(pady = 10, padx = 10)# 'text' is used to write the text on the Button
        CreateToolTip(self.button_ng, text = 
                      'Contact students via the Canvas bulk mailer\n'+
                      'who are not currently enrolled in groups.')

        labelframe2 = tkinter.LabelFrame(self.window, text = "Survey Launch", fg = "black")
        labelframe2.pack(pady = 10)
        
        self.button_tp = tkinter.Button(labelframe2, text = "Create Qualtrics survey\nfrom template", fg = "black", command = upload_template)
        self.button_tp.pack(pady = 10, padx = 10)# 'text' is used to write the text on the Button
        CreateToolTip(self.button_tp, text = 
                      'Upload the Peer Evaluation template\n'+
                      'to Qualtrics to create the survey.')
        
        self.button_dl = tkinter.Button(labelframe2, text = "Create Contacts list\nfrom group membership", fg = "black", command = ExportGroups_call)
        self.button_dl.pack(pady = 10, padx = 10)# 'text' is used to write the text on the Button
        CreateToolTip(self.button_dl, text = 
                      'Export group membership data from Canvas and create\n'+
                      'a Contacts list file to be uploaded to Qualtrics.')

        self.button_si = tkinter.Button(labelframe2, text = "Send survey invitations\nfrom distribution list", fg = "black", command = BulkMail_invitation)
        self.button_si.pack(pady = 10, padx = 10)# 'text' is used to write the text on the Button
        CreateToolTip(self.button_si, text = 
                      'Send students a unique link to the Peer Evaluation from a\n'+ 
                      'Qualtrics distribution list using the Canvas bulk mailer.')


        labelframe3 = tkinter.LabelFrame(self.window, text = "Peer Evaluation Finalisation", fg = "black")
        labelframe3.pack(pady = 10)        
    
        self.button_sr = tkinter.Button(labelframe3, text = "Send survey reminders\nfrom distribution list", fg = "black", command = BulkMail_reminder)
        self.button_sr.pack(pady = 10, padx = 10)# 'text' is used to write the text on the Button
        CreateToolTip(self.button_sr, text = 
                      'Send students who haven\'t completed the Peer Evaluation a reminder to\n'+
                      'complete it (from a Qualtrics distribution list via the Canvas bulk mailer).')

        self.button_pm = tkinter.Button(labelframe3, text = "Upload peer marks\nand comments", fg = "black", command = PeerMark_call)
        self.button_pm.pack(pady = 10)# 'text' is used to write the text on the Button     
        CreateToolTip(self.button_pm, text = 
                      'Calculate peer marks from the Qualtrics survey data and upload\n'+ 
                      'these marks and peer feedback to the Canvas grade book.')

        self.check_status()
        
        if os.path.abspath(os.path.dirname(sys.argv[0]))[-4:].lower() == ".zip":
            messagebox.showinfo("Error", "PEER cannot be run within a ZIP folder. Please extract all of the files and try again.")
            mm.destroy()
            sys.exit()
        
        if config["API_TOKEN"] == "" or config["API_URL"] == "":
            messagebox.showinfo("Warning", "No configuration found. Please go to File -> Settings to set up a link to Canvas.")            
    
    def check_status(self, event_type = ""):
        global canvas, GQL
        if config["API_TOKEN"] == "" or config["API_URL"] == "":
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
            GQL = GraphQL()
            if "course.names" not in session:
                session["course.names"] = []
                session["course.ids"] = []
                courses = GQL.courses()
                for course in courses:
                    session["course.names"].append(courses[course]["name"])
                    session["course.ids"].append(course)
            canvas = Canvas(cleanupURL(config["API_URL"]), config["API_TOKEN"])
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
            
temp_data = {}
canvas = None
bm = None
pm = None

top_frame = tkinter.Tk()
mm = MainMenu(top_frame)
top_frame.mainloop()
