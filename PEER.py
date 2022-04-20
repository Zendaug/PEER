# -*- coding: utf-8 -*-
version = "0.10.0"

print("The Peer Evaluation Enhancement Resource (PEER)\n")

try:
    f = open("Version History.txt", "r")
    updates = f.read()
    f.close()
    print(updates)
except:
    print("Unable to locate Version History.")

from datetime import datetime

print("\n(C) Tim Bednall 2019-{}.".format(datetime.now().year))

import tkinter as tk

import os
import sys
from tkinter import messagebox
from modules.functions import *
import canvasapi

import modules.globals as globals
import modules.server as server

if os.path.abspath(os.path.dirname(sys.argv[0]))[-4:].lower() == ".zip":
    messagebox.showinfo("Error", "PEER cannot be run within a ZIP folder. Please extract all of the files from the ZIP and try again.")
    sys.exit()

import forms.MainMenu as MainMenu
top_frame = tk.Tk()
mm = MainMenu.MainMenu(top_frame)

if globals.config["API_TOKEN"] != "" and globals.config["API_URL"] != "":
    globals.canvas = canvasapi.Canvas(cleanupURL(globals.config["API_URL"]), globals.config["API_TOKEN"])
    globals.user_data = globals.canvas.get_user("self")
    if globals.user_data is not None:
        print("\nWelcome {}.".format(globals.user_data.short_name))
        try:
            server_data = server.user_login() # Logs the user, and returns the new version number
            if version != server_data["version"]:
                resp = messagebox.askquestion("Announcement","A new version of PEER is available! You should download the latest version for the most recent bug fixes and new features.\n\nCurrent version: {}\nNew version: {}\n\nDo you want to download it now?".format(version, server_data["version"]))
                if resp == "yes": open_file(server_data["url"])
        except:
            pass

top_frame.mainloop()