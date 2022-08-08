# -*- coding: utf-8 -*-
"""
Created on Fri Mar 11 15:20:35 2022

@author: tbednall
"""
import requests

from tkinter import messagebox

import modules.globals as globals
import forms.Settings as Settings
import modules.functions as func

class GraphQL():
    def __init__(self, course_id = 0, groupset_id = 0):
        self.course_id = course_id
        self.groupset_id = groupset_id
        self.cache = {}
    
    def query(self, query_text):
        global mm
        query = {'access_token': globals.config["API_TOKEN"], 'query': query_text}
        for attempt in range(5):
            try:
                result = requests.post(url = "{}api/graphql".format(globals.config["API_URL"]), data = query).json()
            except:
                print("Trying to connect to Canvas via GraphQL. Attempt {} of 5.".format(attempt+1))
            else:
                return result
        messagebox.showinfo("Error", "Cannot connect to Canvas via GraphQL. There may be an Internet connectivity problem or the Canvas base URL may be incorrect.")
        Settings.Settings_call()
    
    def courses(self):
        if "courses" in self.cache: return(self.cache["courses"])
        temp = self.query('query MyQuery {allCourses {_id\nname}}')["data"]["allCourses"]
        temp2 = {}
        for course in temp:
            temp2[int(course["_id"])] = {}
            temp2[int(course["_id"])]["name"] = func.clean_text(course["name"])
        df = {}
        for a in sorted(temp2.keys(), reverse = True):
            df[str(a)] = temp2[a]
        self.cache["courses"] = df
        return(df)
    
    def assignments(self):
        if "assn_" + str(self.course_id) not in self.cache:
            self.cache["assn_" + str(self.course_id)] = {}
            try:
                query_result = self.query('query MyQuery {course(id: "' + str(self.course_id) + '") {assignmentsConnection {nodes {_id name pointsPossible assignmentGroup {_id name} groupSet {_id}}}}}')["data"]["course"]["assignmentsConnection"]["nodes"]
                for assn in query_result:
                    self.cache["assn_" + str(self.course_id)][assn["_id"]] = {}
                    self.cache["assn_" + str(self.course_id)][assn["_id"]]["name"] = func.clean_text(assn["name"])
                    self.cache["assn_" + str(self.course_id)][assn["_id"]]["pointsPossible"] = assn["pointsPossible"]
                    self.cache["assn_" + str(self.course_id)][assn["_id"]]["assignmentGroup"] = assn["assignmentGroup"]
                    self.cache["assn_" + str(self.course_id)][assn["_id"]]["groupSet"] = assn["groupSet"]
            except: 
                pass                
        return(self.cache["assn_" + str(self.course_id)])
    
    def group_assignments(self):
        if "groupassn_" + str(self.course_id) not in self.cache:
            self.cache["groupassn_" + str(self.course_id)] = {}
            for assn in self.assignments():
                if self.assignments()[assn]["groupSet"] is not None: 
                    self.cache["groupassn_" + str(self.course_id)][assn] = self.assignments()[assn]
        return self.cache["groupassn_" + str(self.course_id)]
    
    def assignment_groups(self):
        if "assn_group_" + str(self.course_id) not in self.cache:
            self.cache["assn_group_" + str(self.course_id)] = {}
            assns = self.assignments()
            for assn in assns:
                assn_group = assns[assn]["assignmentGroup"]
                if assn_group["_id"] not in self.cache["assn_group_" + str(self.course_id)]:
                    self.cache["assn_group_" + str(self.course_id)][assn_group["_id"]] = {}
                    self.cache["assn_group_" + str(self.course_id)][assn_group["_id"]]["name"] = func.clean_text(assn_group["name"])
        return(self.cache["assn_group_" + str(self.course_id)])        

    def group_sets(self):
        if "groupset_" + str(self.course_id) not in self.cache:
            try:
                self.cache["groupset_" + str(self.course_id)] = self.query('query MyQuery {course(id:"' + str(self.course_id) + '") {groupSetsConnection {nodes {_id\nname}}}}')["data"]["course"]["groupSetsConnection"]["nodes"]
            except: 
                self.cache["groupset_" + str(self.course_id)] = {}
        return(self.cache["groupset_" + str(self.course_id)])

    def sections(self):
        if "sections_" + str(self.course_id) not in self.cache:
            query_result = self.query('query MyQuery {course(id: "' + str(self.course_id) + '") {sectionsConnection {nodes {_id\nname}}}}')["data"]["course"]["sectionsConnection"]["nodes"]
            self.cache["sections_" + str(self.course_id)] = {}
            for section in query_result:
                self.cache["sections_" + str(self.course_id)][section["_id"]] = {}
                self.cache["sections_" + str(self.course_id)][section["_id"]]["name"] = func.clean_text(section["name"])
        return(self.cache["sections_" + str(self.course_id)])
    
    def students(self):
        if "students_" + str(self.course_id) not in self.cache:
            query_result = self.query('query MyQuery {course(id: "' + str(self.course_id) + '") {enrollmentsConnection {nodes {type\nuser {_id\nname\nemail}}}}}')
            self.cache["students_" + str(self.course_id)] = {"id": [], "name": [], "email": []}
            for student in query_result["data"]["course"]["enrollmentsConnection"]["nodes"]:
                if (student["type"] == "StudentEnrollment"):
                    if (student["user"]["_id"] not in self.cache["students_" + str(self.course_id)]["id"]):
                        self.cache["students_" + str(self.course_id)]["id"].append(student["user"]["_id"])
                        self.cache["students_" + str(self.course_id)]["name"].append(func.clean_text(student["user"]["name"]))
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
                        df[group["_id"]] = {"name": func.clean_text(group["name"]), "users": []}
                        for user in group["membersConnection"]["nodes"]:
                            df[group["_id"]]["users"].append(user["user"]["_id"])
                    self.cache["groups_" + str(self.course_id)] = df
                    return(df)
                    break
        except:
            return({})
        return({})
    
    def students_comprehensive(self):
        print("Course ID is: " + str(self.course_id))
        if "students_" + str(self.course_id) + "_" + str(self.groupset_id) in self.cache:
            return(self.cache["students_" + str(self.course_id) + "_" + str(self.groupset_id)])
        query_result = self.query('query MyQuery {course(id: "' + str(self.course_id) + '") {enrollmentsConnection {nodes {type\nuser {_id\nname\nemail\nsisId}\nsection {_id}}}}}')["data"]["course"]["enrollmentsConnection"]["nodes"]
        student_series = {}
        # Set up dictionary for students
        for student in query_result:
            if (student["type"] == "StudentEnrollment"):
                if (student["user"]["_id"] not in student_series):
                    student_series[student["user"]["_id"]] = {"name": func.clean_text(student["user"]["name"]), "email": student["user"]["email"], "sisId": student["user"]["sisId"]}
                    student_series[student["user"]["_id"]]["first_name"] = func.clean_text(student["user"]["name"][0:student["user"]["name"].find(" ")])
                    student_series[student["user"]["_id"]]["last_name"] = func.clean_text(student["user"]["name"][student["user"]["name"].rfind(" ")+1:])
                    student_series[student["user"]["_id"]]["sections"] = []

                    # Recreate the student's email address from the sis ID, if the email field is blank
                    if (student_series[student["user"]["_id"]]["email"] is None or student_series[student["user"]["_id"]]["email"] == "") and globals.config["EmailFormat"] != "":
                        student_series[student["user"]["_id"]]["email"] = globals.config["EmailFormat"]
                        student_series[student["user"]["_id"]]["email"] = student_series[student["user"]["_id"]]["email"].replace("[sisId]", str(student_series[student["user"]["_id"]]["sisId"]) or "") # If the replacement is None, replace with ""
                        student_series[student["user"]["_id"]]["email"] = student_series[student["user"]["_id"]]["email"].replace("[first_name]", str(student_series[student["user"]["_id"]]["first_name"]) or "")
                        student_series[student["user"]["_id"]]["email"] = student_series[student["user"]["_id"]]["email"].replace("[last_name]", str(student_series[student["user"]["_id"]]["last_name"]) or "")
                
                student_series[student["user"]["_id"]]["sections"].append(student["section"]["_id"])
        
        # Identify which group the student is in
        if self.groupset_id != 0: #Only look for groups if a group set has been specified
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
                    student_series[student]["group_name"] = func.clean_text(groups[group]["name"])
                            
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
                    teacher_series[teacher["user"]["_id"]] = {"name": func.clean_text(teacher["user"]["name"]), "email": teacher["user"]["email"]}
                    teacher_series[teacher["user"]["_id"]]["first_name"] = func.clean_text(teacher["user"]["name"][0:teacher["user"]["name"].find(" ")])
                    teacher_series[teacher["user"]["_id"]]["last_name"] = func.clean_text(teacher["user"]["name"][teacher["user"]["name"].rfind(" ")+1:])
                    teacher_series[teacher["user"]["_id"]]["sections"] = []
                teacher_series[teacher["user"]["_id"]]["sections"].append(teacher["section"]["_id"])
                
        self.cache["teachers_" + str(self.course_id) + "_" + str(self.groupset_id)] = teacher_series                
        return(teacher_series)

    def user_details(self):
        query_result = self.query('query MyQuery {course(id: "' + str(self.course_id) + '") {enrollmentsConnection {nodes {type\nuser {_id\nname\nemail}\nsection {_id}}}}}')["data"]["course"]["enrollmentsConnection"]["nodes"]
        return(query_result)
    
    def whoami(self):
        result = requests.post(url = globals.config["API_URL"] + "api/v1/inst_access_tokens", 
                      headers = {'Accept': 'application/json',
                                 'Authorization': 'Bearer ' + globals.config["API_TOKEN"]}).json()
        #print(result["token"])
        print(result)
        result2 = requests.post(url = "https://swinburnesarawak.api.instructure.com/graphql", 
                      headers = {'Content-Type': 'application/json',
                                 'Accept': 'application/json',
                                 'Authorization': 'Bearer ' + result["token"]},
                      data = '{"query":"{ whoami { userUuid } }"}').json()
        
        #userUuid
        
        print(result2)
        result3 = requests.post(url = "https://swinburnesarawak.api.instructure.com/graphql", 
                                headers = {'Content-Type': 'application/json',
                                           'Accept': 'application/json',
                                           'Authorization': 'Bearer ' + result["token"]},
                                data = '{"query":"{ account(id="' + result2["data"]["whoami"]["userUuid"] + '") {id name} }}"').json()

        print(result3)
        return(result).json()