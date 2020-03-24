# -*- coding: utf-8 -*-

updates = """
Version 0.1.0 (2020-02-22): Default bulk messages are now editable in the 'default.txt' and 'config.txt' files.
Version 0.0.5 (2020-02-21): Fixed code to automatically hide grades from the students. In the SaveData file, the rater appears before the receiver.
Version 0.0.4 (2020-02-20): Only weekdays are now counted when calculating late penalties.
Version 0.0.3: Fixed a bug that caused the program to crash if there was an empty group (under "Export Groups").
Version 0.0.2: Fixed bug with repeated "other" list, and made the uploading of marks more robust.
Version 0.0.1: Fixed bug, regarding "PointsPossible" property of the "config" dictionary

(C) Tim Bednall 2019-2020."""

print("Loading Peer Evaluation Enhancement Resource (PEER)...\n\nVersion history:")
print(updates)

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


# Load the configuration file
# First of all, load the hard defaults
hard_defaults = {
    'API_URL': "",
    'API_TOKEN': "",
    'EmailSuffix': "",
    'ScoreType': 2,
    'SelfVote': 2,
    'AdjustedMark_Name': "Peer Mark",
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
    'subject_confirm': "Confirming Group Membership",
    'message_confirm': "Dear all,\n\nI am writing to confirm your membership in: [Group Name].\n\nAccording to our records, the members of this group include:\n[Group Members]\nIf there are any errors in our records -- such as missing group members or people listed who are not in your group -- please contact me as soon as possible. You may ignore this message if these records are correct.",
    'subject_nogroup' : "Lack of Group Membership",
    'message_nogroup' : "Dear [First Name],\n\nI am writing because according to our records, you are not currently listed as a member of any student team for this unit.\n\nAs you would be aware, there is a group assignment due later in the semester. If you do not belong to a team, you will be unable to complete the assignment.\n\nIf there is another circumstance that would prevent you from joining a group project, please let me know as soon as possible.\n\nPlease let me know if there is anything I can do to support you in the meantime.",
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
        if days_late(ID) > 0: pen = pen - (config["Penalty_PerDayLate"]/100)*days_late(ID)
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
        if days_late(ID) > 0 and config["Penalty_PerDayLate"] > 0: penalty_text = penalty_text + "-" + str(config["Penalty_PerDayLate"] * days_late(ID)) + "% because you submitted your evaluation " + str(days_late(ID)) + " day(s) late.\n"
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
#%% Settings
    
class Settings:
    def __init__(self, window):
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
        tkinter.Label(window, text = "\nPlease enter the student email suffix:", fg = "black").grid(row = 4, column = 1, sticky = "W")
        tkinter.Label(window, text = "Suffix:", fg = "black").grid(row = 5, column = 0, sticky = "E")
        self.settings_suffix = tkinter.Entry(window, width = 50)
        self.settings_suffix.grid(row = 5, column = 1, sticky="W", padx = 5)
        tkinter.Label(window, text = "You will only need to enter these settings once.", fg = "black").grid(row = 6, columnspan = 2)
        self.settings_default = tkinter.Button(window, text = "Restore Default Settings", fg = "black", command = self.restore_defaults).grid(row = 7, columnspan = 2, pady = 5)
        self.settings_save = tkinter.Button(window, text = "Save Settings", fg = "black", command = self.save_settings).grid(row = 8, columnspan = 2, pady = 5)
        self.update_fields()

    def update_fields(self):
        self.settings_URL.delete(0, tkinter.END)
        self.settings_URL.insert(0, config["API_URL"])
        self.settings_Token.delete(0, tkinter.END)
        self.settings_Token.insert(0, config["API_TOKEN"])
        self.settings_suffix.delete(0, tkinter.END)
        self.settings_suffix.insert(0, config["EmailSuffix"])
        
    def restore_defaults(self):
        global config
        config = {}
        for item in defaults: config[item] = defaults[item]
        self.update_fields()
        
    def save_settings(self):
        config["API_URL"] = self.settings_URL.get()
        config["API_TOKEN"] = self.settings_Token.get()
        config["EmailSuffix"] = self.settings_suffix.get()
        f = open("config.txt", "w")
        f.write(json.dumps(config, sort_keys=True, separators=(',\n', ':')))
        f.close()
        self.window.destroy()

def Settings_call():
    settings_win = tkinter.Toplevel()
    set_win = Settings(settings_win)
    settings_win.mainloop()

#%% Export Groups
class ExportGroups():
    def __init__(self, exportgroups):
        self.exportgroups = exportgroups
        exportgroups.title("Export Groups")
        exportgroups.resizable(0, 0)
        exportgroups.grab_set()
        
        tkinter.Label(exportgroups, text = "\nDownload student groups and group membership", fg = "black").grid(row = 0, column = 0, columnspan = 2)
        tkinter.Label(exportgroups, text = "Unit:", fg = "black").grid(row = 1, column = 0, sticky = "E", pady = 5)           
        self.exportgroups_unit = ttk.Combobox(exportgroups, width = 60, values = [], state="readonly")
        self.exportgroups_unit.bind("<<ComboboxSelected>>", self.group_sets)
        self.exportgroups_unit.grid(row = 1, column = 1, sticky = "W", padx = 5)

        tkinter.Label(exportgroups, text = "Group Set:", fg = "black").grid(row = 2, column = 0, sticky = "E", pady = 5)
        self.exportgroups_groupset = ttk.Combobox(exportgroups, width = 60, state="readonly")
        self.exportgroups_groupset.grid(row = 2, column = 1, sticky = "W", padx = 5)

        tkinter.Label(exportgroups, text = "Confirmation:", fg = "black").grid(row = 3, column = 0, sticky = "E", pady = 5)
        self.contactgroups = tkinter.BooleanVar()
        self.contactgroups.set(False)
        self.exportgroups_contact = tkinter.Checkbutton(exportgroups, variable = self.contactgroups, text = "Automatically contact students via Canvas to confirm group membership")
        self.exportgroups_contact.grid(row = 3, column = 1, sticky = "W")
        
        self.exportgroups_export = tkinter.Button(exportgroups, text = "Download Groups", fg = "black", command = self.begin_export)
        self.exportgroups_export.grid(row = 4, columnspan = 2, pady = 5)

        self.group_sets_id = []
        self.group_sets_name = []          
        
        if "courses" not in session:
            session["courses"] = []
            session["course_ids"] = []
            for course in canvas.get_courses():
                session["courses"].append(course.name)
                session["course_ids"].append(course.id)
        self.exportgroups_unit["values"] = tuple(session["courses"])
        self.exportgroups_unit.current(0)
        self.group_sets()
        
        self.statusbar = tkinter.Label(exportgroups, text="", bd=1, relief=tkinter.SUNKEN, anchor=tkinter.W)
        self.statusbar.grid(row = 5, columnspan = 2, sticky = "ew")
        
        self.status("")
    
    def group_sets(self, event_type = ""):
        self.group_sets_id = []
        self.group_sets_name = []
        course = canvas.get_course(session["course_ids"][self.exportgroups_unit.current()])
        group_categories = course.get_group_categories()
        for group_set in group_categories:
            self.group_sets_id.append(group_set.id)
            self.group_sets_name.append(group_set.name)
        self.exportgroups_groupset["values"] = tuple(self.group_sets_name)
        self.exportgroups_groupset.current(0)
    
    def status(self, text): 
        self.statusbar["text"] = text
        self.statusbar.update_idletasks()
    
    def begin_export(self, event_type = ""):
        foldername = ""
        foldername = tkinter.filedialog.askdirectory(title = "Choose Folder to Save Group Membership Information...")
        
        course = canvas.get_course(session["course_ids"][self.exportgroups_unit.current()])
        group_set = self.group_sets_id[self.exportgroups_groupset.current()]
        print("Accessing " + course.name + "...")
        
        users = course.get_users(enrollment_type=['student'])
        user_list = {"Email":[],
                     "ExternalDataReference":[],
                     "FirstName":[],
                     "LastName":[],
                     "Student_name":[],
                     "Group_name":[],
                     "Peer1_name":[],
                     "Peer2_name":[],
                     "Peer3_name":[],
                     "Peer4_name":[],
                     "Peer5_name":[],
                     "Peer6_name":[],
                     "Peer7_name":[],
                     "Peer8_name":[],
                     "Peer9_name":[],
                     "Student_id":[],
                     "Group_id":[],
                     "Peer1_id":[],
                     "Peer2_id":[],
                     "Peer3_id":[],
                     "Peer4_id":[],
                     "Peer5_id":[],
                     "Peer6_id":[],
                     "Peer7_id":[],
                     "Peer8_id":[],
                     "Peer9_id":[],
                     }
        
        print("Downloading student details...")
        for user in users:
            print("> " + user.name)
            if config["EmailSuffix"] != np.nan and config["EmailSuffix"] != "":
                user_list["Email"].append(str(user.sis_user_id) + config["EmailSuffix"])
            else:
                user_list["Email"].append("")
            user_list["ExternalDataReference"].append(user.id)
            user_list["FirstName"].append(user.name[0:user.name.find(" ")])
            user_list["LastName"].append(user.name[user.name.rfind(" ")+1:])
            user_list["Student_name"].append(user.name)
            user_list["Group_name"].append(np.NaN)
            user_list["Peer1_name"].append(np.NaN)
            user_list["Peer2_name"].append(np.NaN)
            user_list["Peer3_name"].append(np.NaN)
            user_list["Peer4_name"].append(np.NaN)
            user_list["Peer5_name"].append(np.NaN)
            user_list["Peer6_name"].append(np.NaN)
            user_list["Peer7_name"].append(np.NaN)
            user_list["Peer8_name"].append(np.NaN)
            user_list["Peer9_name"].append(np.NaN)
            user_list["Student_id"].append(user.id)   
            user_list["Group_id"].append(np.NaN)
            user_list["Peer1_id"].append(np.NaN)
            user_list["Peer2_id"].append(np.NaN)
            user_list["Peer3_id"].append(np.NaN)
            user_list["Peer4_id"].append(np.NaN)
            user_list["Peer5_id"].append(np.NaN)
            user_list["Peer6_id"].append(np.NaN)
            user_list["Peer7_id"].append(np.NaN)
            user_list["Peer8_id"].append(np.NaN)
            user_list["Peer9_id"].append(np.NaN)
        
        user_list = pd.DataFrame(user_list)
        
        groups = course.get_groups()
        group_list = {"Group_name":[], "Group_id":[], "Student_id": []}
        print("Downloading group information...")
        for a, group in enumerate(groups):
            if group.group_category_id == group_set:
                print("Found group: " + group.name)
                group_list["Group_name"].append(clean_text(str(group.name)))
                group_list["Group_id"].append(group.id)
                group_members = group.get_users()
                group_list["Student_id"].append([])
                for group_member in group_members:
                    group_list["Student_id"][a].append(group_member.id)
        
        # Write CSV of group members
        print("Creating list of students and their group peers...")
        for a in range(0, len(group_list["Group_id"])):
            for group_member in group_list["Student_id"][a]:
                user_list.loc[user_list["Student_id"] == group_member, "Group_name"] = group_list["Group_name"][a]
                user_list.loc[user_list["Student_id"] == group_member, "Group_id"] = group_list["Group_id"][a]
              
        for user_id in user_list.loc[np.isnan(user_list["Group_id"]) == False, "Student_id"]:
            group_id = user_list.loc[user_list["Student_id"]==user_id, "Group_id"].values[0]
            a = group_list["Group_id"].index(group_id)
            counter = 1
            for group_mem in group_list["Student_id"][a]:
                if group_mem != user_id:
                    user_list.loc[user_list["Student_id"]==user_id, "Peer" + str(counter) + "_id"] = group_mem
                    user_list.loc[user_list["Student_id"]==user_id, "Peer" + str(counter) + "_name"] = user_list.loc[user_list["Student_id"]==group_mem, "Student_name"].values[0]
                    counter = counter + 1
        
        
        user_list.loc[np.isnan(user_list["Group_id"]) == False].to_csv(os.path.join(foldername,"GroupMembers.csv"), index = False)
        
        # Write CSV of students lacking groups
        print("Creating list of students without groups...")
        user_list.loc[np.isnan(user_list["Group_id"]), ["Email", "ExternalDataReference", "FirstName", "Student_id", "Student_name"]].to_csv(os.path.join(foldername,"NoGroups.csv"), index = False)

        info_text = "Group membership details were saved in \"GroupMembers.csv\" and a list of students without groups has been saved in \"NoGroups.csv\"."
        
        if self.contactgroups.get():
            print("Sending confirmation message to groups via Canvas...")
            for a in range(0, len(group_list["Group_id"])):
                conf_message = config["message_confirm"].replace("[Group Name]", group_list["Group_name"][a])
                grp_list = ""
                iter = 0
                for student in group_list["Student_id"][a]:
                    iter = iter + 1
                    grp_list = grp_list + "* " + user_list.loc[user_list["Student_id"]==student, "Student_name"].values[0] + "\n"
                if iter > 0: # Only send messages to groups with 1 or more members
                    conf_message = conf_message.replace("[Group Members]", grp_list)
                    message = canvas.create_conversation(recipients = ['group_' + str(group_list["Group_id"][a])],
                                                         body = conf_message,
                                                         subject = config["subject_confirm"],
                                                         force_new = True,
                                                         group_conversation = True,
                                                         context_code = "course_" + str(course.id))
                
            for student in user_list.loc[np.isnan(user_list["Group_id"]), "Student_id"]:
                conf_message = config["message_nogroup"].replace("[First Name]", user_list.loc[user_list["Student_id"] == student, "FirstName"].values[0])
                message = canvas.create_conversation(recipients = [str(student)],
                                                     body = conf_message,
                                                     subject = config["subject_nogroup"],
                                                     force_new = True,
                                                     context_code = "course_" + str(course.id))
            info_text = info_text + " A message has been sent to students asking for confirmation of their group membership."

        messagebox.showinfo("Download complete", info_text)
        self.exportgroups.destroy()

def ExportGroups_call():
    exportgroups = tkinter.Toplevel()
    exp_grp = ExportGroups(exportgroups)
    exportgroups.mainloop()

#%% Peer Mark section


class PeerMark():
   def __init__(self):
       self.peermark = tkinter.Toplevel()
       self.sentinel = None

       self.peermark.title("Calculate Peer Marks")
       self.peermark.resizable(0, 0)
       self.peermark.grab_set()
    
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
       
       if "courses" not in session:
           session["courses"] = []
           session["course_ids"] = []
           for course in canvas.get_courses():
               session["courses"].append(course.name)
               session["course_ids"].append(course.id)
       self.peermark_unit["values"] = tuple(session["courses"])
       self.peermark_unit.current(0)
       self.group_assns()
        
       self.status_text = tkinter.StringVar()
       self.statusbar = tkinter.Label(self.peermark, bd=1, relief=tkinter.SUNKEN, anchor=tkinter.W, textvariable = self.status_text)
       self.statusbar.grid(row = 6, columnspan = 3, sticky = "ew")       
       
       self.peermark.mainloop()

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
        course = canvas.get_course(session["course_ids"][self.peermark_unit.current()])
        assns = course.get_assignments()
        for assn in assns:
            self.assn_id.append(assn.id)
            self.assn_name.append(assn.name)
        self.peermark_assn["values"] = tuple(self.assn_name)
        self.peermark_assn.current(0)
    
   def status(self, stat_text):
       self.status_text.set(stat_text)
       time.sleep(0.01)
       
   def start_calculate(self):
       self.status("Loading, please be patient...")
       session["DueDate"] = self.peermark_duedate.get_date()
       session["GroupAssignment_ID"] = self.assn_id[self.peermark_assn.current()]
       session["SaveData"] = self.UploadPeerMark.get()
       course = canvas.get_course(session["course_ids"][self.peermark_unit.current()])
       calculate_marks(course, self)

def PeerMark_call():
    global pm
    pm = PeerMark()

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
           pm.peermark.destroy()
   except:
       messagebox.showinfo("Error", "Unable to parse data from this file.")
       pm.peermark.peermark.destroy()
   
   # Create new assignment
   print("Creating new assignment for Peer Mark")
   new_assignment = course.create_assignment({'name': config["AdjustedMark_Name"],
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
       
   new_assignment.edit(assignment = {"published": False})
   
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
   pm.peermark.destroy()
   
#%% Policies
class Policies():
    def __init__(self):
        self.policies = tkinter.Toplevel()
        self.validation = self.policies.register(only_numbers)
        
        self.policies.title("Calculate Peer Marks")
        self.policies.resizable(0, 0)
        self.policies.grab_set()
                
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

        tkinter.Label(labelframe3, text = "Peer mark scale ranges from 1 to:", fg = "black").grid(row = 1, column = 0, sticky = "NE")
        self.points_possible = tkinter.Entry(labelframe3, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.points_possible.grid(row = 1, column = 1, sticky="W", padx = 5)
        self.points_possible.insert(0, config["PointsPossible"])

        tkinter.Label(labelframe3, text = "Rescale peer mark to score out of:", fg = "black").grid(row = 2, column = 0, sticky = "NE")
        self.rescale_to = tkinter.Entry(labelframe3, width = 5, validate="key", validatecommand=(self.validation, '%S'))
        self.rescale_to.grid(row = 2, column = 1, sticky="W", padx = 5)
        self.rescale_to.insert(0, config["RescaleTo"])
        
        self.save_policies = tkinter.IntVar()
        self.save_policies.set(1)
        tkinter.Label(labelframe3, text = "Retain policies?", fg = "black").grid(row = 3, column = 0, sticky = "NE")
        ttk.Checkbutton(labelframe3, text = "Save policies for future assessments", variable = self.save_policies).grid(row = 3, column = 1, sticky = "W")
        
        tkinter.Button(self.policies, text = "Continue", fg = "black", command = self.save_settings).grid(row = 13, column = 0, columnspan = 3, pady = 5)
        self.policies.mainloop()
    
    def save_settings(self):
        config["ScoreType"] = int(self.policies_scoring.current()+1)
        config["SelfVote"] = int(self.policies_selfrat.current())
        config["MinimumPeers"] = int(self.policies_minimum.get())
        config["Penalty_NonComplete"] = int(self.penalty_noncomplete.get())
        config["Penalty_PerDayLate"] = int(self.penalty_late.get())
        config["Penalty_SelfPerfect"] = int(self.penalty_selfperfect.get())
        config["Penalty_PeersAllZero"] = int(self.penalty_allzero.get())
        config["PointsPossible"] = int(self.points_possible.get())
        config["RescaleTo"] = int(self.rescale_to.get())
        
        if self.save_policies.get() == 1:
            f = open("config.txt", "w")
            f.write(json.dumps(config, sort_keys=True, separators=(',\n', ':')))
            f.close()
        
        self.policies.destroy()
        
def Policies_call():
    pol = Policies()

#%% Bulk mail
class BulkMail:
    def __init__(self):
        self.window = tkinter.Toplevel()
        self.window.title("Send Bulk Message via Canvas")
        self.window.resizable(0, 0)
        self.window.grab_set()
        
        tkinter.Label(self.window, text = "Unit:", fg = "black").grid(row = 0, column = 0, sticky = "E", pady = 5)
        self.window_unit = ttk.Combobox(self.window, width = 60, state="readonly")
        self.window_unit.grid(row = 0, column = 1, sticky = "W", padx = 5)
        
        tkinter.Label(self.window, text = "Distribution list file:", fg = "black").grid(row = 2, column = 0, sticky = "E", pady = 5)
        self.window_distlist = tkinter.Button(self.window, text = "Select...", fg = "black", command = self.select_file)
        self.window_distlist.grid(row = 2, column = 1, pady = 5, padx = 5, sticky = "W")
        self.datafile = tkinter.Label(self.window, text = "", fg = "black")
        self.datafile.grid(row = 3, column = 1, pady = 5, sticky = "W")
        
        tkinter.Label(self.window, text = "Send message to:", fg = "black").grid(row = 4, column = 0, sticky = "E", pady = 5)
        self.window_sendto = ttk.Combobox(self.window, width = 60, state="readonly")
        self.window_sendto.grid(row = 4, column = 1, sticky = "W", pady = 5, padx = 5)
        self.window_sendto["values"] = ["All students in list"]
        self.window_sendto.current(0)
        self.window_sendto.bind("<<ComboboxSelected>>", self.change_message)
        
        tkinter.Label(self.window, text = "Message:", fg = "black").grid(row = 5, column = 0, columnspan = 2, pady = 5, padx = 5)

        tkinter.Label(self.window, text = "Subject:", fg = "black").grid(row = 6, column = 0, sticky = "E", pady = 5)
        self.subject_text = tkinter.StringVar()
        self.subject = tkinter.Entry(self.window, textvariable = self.subject_text)
        self.subject.grid(row = 6, column = 1, pady = 5, padx = 5, sticky = "WE")

        self.message = tkinter.Text(self.window, width = 60, height = 15, wrap = tkinter.WORD)
        self.message.grid(row = 7, column = 0, columnspan = 2, sticky = "W", pady = 5, padx = 5)
        
        self.sendmessage = tkinter.Button(self.window, text = "Send Bulk Message", fg = "black", state = "disabled", command = self.MessageSend)
        self.sendmessage.grid(row = 8, column = 0, columnspan = 2, pady = 5)
       
        if "courses" not in session:
            session["courses"] = []
            session["course_ids"] = []
            for course in canvas.get_courses():
                session["courses"].append(course.name)
                session["course_ids"].append(course.id)
        self.window_unit["values"] = tuple(session["courses"])
        self.window_unit.current(0)
        
        self.window.mainloop()
        
    def select_file(self):
       session["DistList"] = filedialog.askopenfilename(initialdir = ".", title = "Select file",filetypes = (("Compatible File Types",("*.csv")),("All Files","*.*")))
       if len(session["DistList"]) > 0:
           try:
               self.distlist = pd.read_csv(session["DistList"])
               if any(a in self.distlist.columns for a in ["External Data Reference", "ExternalDataReference", "Student_id"]):
                   self.datafile["text"] = session["DistList"]
                   if "Student_id" in self.distlist.columns: session["ID_column"] = "Student_id"
                   if "External Data Reference" in self.distlist.columns: session["ID_column"] = "External Data Reference"
                   if "ExternalDataReference" in self.distlist.columns: session["ID_column"] = "ExternalDataReference"
                   self.sendmessage["state"] = "normal"
                   if "Status" in self.distlist.columns: 
                       self.window_sendto["values"] = ["All students in list", "Students who have not completed the peer evaluation", "Students who have completed the peer evalaution"]
                       self.change_message()
               else:
                   messagebox.showinfo("Error", "Cannot find user ID column in file. This column should be labelled 'Student_id' or 'External Data Reference'.")
           except:
               messagebox.showinfo("Error", "Trouble parsing file.")
       
    def change_message(self, event_type = ""):
        if len(self.window_sendto["values"]) > 1:
            self.message.delete(1.0, tkinter.END)
            self.subject_text.set("")
            if self.window_sendto.current() == 0:
                self.subject_text.set(config["subject_invitation"])
                self.message.insert(tkinter.END, config["message_invitation"])
            if self.window_sendto.current() == 1:
                self.subject_text.set(config["subject_reminder"])
                self.message.insert(tkinter.END, config["message_reminder"])            
            if self.window_sendto.current() == 2:
                self.subject_text.set(config["subject_thankyou"])
                self.message.insert(tkinter.END, config["message_thankyou"])
                
    def MessageSend(self):
        # Filter participants if requested
        print("Sending messages...")
        if self.window_sendto.current() == 1: self.distlist = self.distlist[self.distlist["Status"] != "Finished Survey"]
        if self.window_sendto.current() == 2: self.distlist = self.distlist[self.distlist["Status"] == "Finished Survey"]
        if self.window_sendto.current() > 0: self.distlist = self.distlist.reset_index(drop = True) # Reset the index
        for y, ID in enumerate(self.distlist[session["ID_column"]]):
            temp_text = self.message.get(1.0, tkinter.END)
            for col_name in self.distlist.columns:
                temp_text = temp_text.replace("["+col_name+"]", str(self.distlist.loc[y, col_name]))
            print("Sending message " + str(y+1) + " of " + str(len(self.distlist[session["ID_column"]])))
            #print(temp_text + "\n\n")
            bulk_message = canvas.create_conversation(recipients = [str(ID)],
                                                 body = temp_text,
                                                 subject = self.subject_text.get(),
                                                 force_new = True,
                                                 context_code = "course_" + str(session["course_ids"][self.window_unit.current()]))
            
        messagebox.showinfo("Finished", "All messages have been sent.")
        self.window.destroy()

def BulkMail_call():
    bm = BulkMail()

#%% Main Menu
class MainMenu:
    def __init__(self):
        self.window = tkinter.Tk()
        self.window.title("PEER")
        self.window.geometry("300x250") # size of the window width:- 500, height:- 375
        self.window.resizable(0, 0) # this prevents from resizing the window
        self.window.bind("<FocusIn>", self.check_status)
        top_frame = tkinter.Frame(self.window).pack(side="top", pady = 10)
        tkinter.Button(top_frame, text = "Modify user settings", fg = "black", command = Settings_call).pack(pady = 10)# 'fg - foreground' is used to color the contents
        self.button_dl = tkinter.Button(top_frame, text = "Download student groups\nfrom Canvas", fg = "black", command = ExportGroups_call)
        self.button_dl.pack(pady = 10)# 'text' is used to write the text on the Button
        self.button_bulk = tkinter.Button(top_frame, text = "Send bulk message\nvia Canvas", fg = "black", command = BulkMail_call)
        self.button_bulk.pack(pady = 10)
        self.button_pm = tkinter.Button(top_frame, text = "Calculate peer marks", fg = "black", command = PeerMark_call)
        self.button_pm.pack(pady = 10)# 'text' is used to write the text on the Button
        self.check_status()
        self.window.mainloop()
    
    def check_status(self, event_type = ""):
        global canvas
        if config["API_TOKEN"] == "" or config["API_URL"] == "" or config["EmailSuffix"] == "":
            self.button_dl["state"] ='disabled'
            self.button_bulk["state"] = 'disabled'
            self.button_pm["state"] ='disabled'
        else:
            canvas = Canvas(config["API_URL"], config["API_TOKEN"])
            self.button_dl["state"] ='normal'
            self.button_bulk["state"] = 'normal'
            self.button_pm["state"] ='normal'

temp_data = {}
canvas = None
mm = MainMenu()

