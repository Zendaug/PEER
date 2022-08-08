# -*- coding: utf-8 -*-
"""
Created on Fri Mar 11 13:56:23 2022

@author: tbednall
"""

import json
import os
from canvasapi import Canvas

from modules.functions import *

# Global variables
version = 0
canvas = None
canvas_data = ""
hard_defaults = {}

tw_list = [] # Questions that make up the teamwork score
tw_names = [] # The names of the teamwork variables 
other_list = [] # The names of other variables

# Set up dictionary for ratios and adjusted scores for fast lookup
adj_dict = {}

config = {}

hard_defaults = {
    'API_URL': "",
    'API_TOKEN': "",
    'EmailSuffix': "",
    'ScoreType': 1,
    'SelfVote': 0,
    'PeerMark_Name': "Peer Mark",
    'MinimumPeers': 2,
    'MinimumPeersPolicy': 1,
    'MinimumAdjustment': 2,
    'Penalty_NonComplete': 100,
    'Penalty_PartialComplete': 0,
    'Penalty_SelfPerfect': 0,
    'Penalty_PeersAllZero': 0,
    'Penalty_PerDayLate': 0,
    'Penalty_Late_Custom': True,
    'Exclude_PartialComplete': 0,
    'Exclude_SelfPerfect': 0,
    'Exclude_PeersAllZero': 0,
    'PointsPossible': 5,
    'RescaleTo': 5,
    'publish_assignment': True,
    'weekdays_only': True,
    'subject_confirm': "Confirming Group Membership",
    'message_confirm': "Dear [firstnames_list],\n\nI am writing to confirm your membership in: [name].\n\nAccording to our records, the members of this group include:\n\n[members_bulletlist]\n\nIf there are any errors in our records -- such as missing group members or people incorrectly listed as group members -- please contact me as soon as possible. You may ignore this message if these records are correct.\n\nWe encourage you to visit your group's homepage, a platform where you can collaborate with your fellow members. Your group's homepage is located at: [group_homepage]",
    'subject_nogroup' : "Lack of Group Membership",
    'message_nogroup' : "Dear [first_name],\n\nI am writing because according to our records, you are not currently listed as a member of any student team for this unit.\n\nAs you would be aware, there is a group assignment due later in the semester. If you do not belong to a team, you will be unable to complete the assignment.\n\nIf there is another circumstance that would prevent you from joining a group project, please let me know as soon as possible.\n\nPlease let me know if there is anything I can do to support you in the meantime.",
    'subject_invitation': "Peer Evaluation now available",
    'message_invitation': {"Qualtrics": "Dear [First Name],\n\nThe Peer Evaluation survey is now available.\n\nYou may access it via this link: [Link]\n\nThis link is unique to yourself, so please do not share it with other people.",
                           "LimeSurvey": "Dear [firstname],\n\nThe Peer Evaluation survey is now available.\n\nYou may access it via this link: (obtain this link from LimeSurvey)\n\nYour access code is: [token]."},
    'subject_reminder': "Reminder: Peer Evaluation due",
    'message_reminder': {"Qualtrics": "Dear [First Name],\n\nThis is a courtesy reminder to complete the Peer Evaluation survey. According to our records, you have not yet completed this survey.\n\nYou may access the survey via this link: [Link]\n\nThis link is unique to yourself, so please do not share it with other people.",
                         "LimeSurvey": "Dear [firstname],\n\nThis is a courtesy reminder to complete the Peer Evaluation survey. According to our records, you have not yet completed this survey.\n\nYou may access the survey via this link: (obtain this link from LimeSurvey)\n\nYour access code is: [token]."},
    'subject_thankyou' : "Thank you for completing the Peer Evaluation",
    'message_thankyou' : {"Qualtrics": "Dear [First Name],\n\nThank you for completing the Peer Evaluation. There is no need to take any further action.",
                          "LimeSurvey": "Dear [firstname],\n\nThank you for completing the Peer Evaluation. There is no need to take any further action."},
    'firsttime_export' : True,
    'firsttime_upload' : True,
    'ExportData': True,
    'UploadData': True,
    "feedback_only": False,
    "SurveyPlatform": "Qualtrics",
    "EmailFormat": ""
    }

config = hard_defaults

# Overwrite hard defaults with soft defaults
try:
    f = open("defaults.txt", "r")
    defaults = json.loads(f.read())
    f.close()
    for item in defaults:
        config[item] = defaults[item]
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
    try:
        os.remove("update_completed.txt")
    except:
        pass
    os.rename("update.txt", "update_completed.txt")
    save_config(config)
except:
    pass

session = {}

user_data = None

pm = None

GQL = None