# -*- coding: utf-8 -*-
version = "0.11.0"

from datetime import datetime

import tkinter as tk

import os
import sys
from tkinter import messagebox
import modules.functions as func
import canvasapi

import modules.globals as globals
import modules.server as server
globals.version = version

if os.path.abspath(os.path.dirname(sys.argv[0]))[-4:].lower() == ".zip":
    messagebox.showinfo("Error", "PEER cannot be run within a ZIP folder. Please extract all of the files from the ZIP and try again.")
    sys.exit()

if __name__ == "__main__":
    import forms.MainMenu as MainMenu
    top_frame = tk.Tk()
    mm = MainMenu.MainMenu(top_frame)
    top_frame.resizable(True, True)

print("The Peer Evaluation Enhancement Resource (PEER)")
print("(C) Tim Bednall 2019-{}.".format(datetime.now().year))

if globals.config["API_TOKEN"] != "" and globals.config["API_URL"] != "":
    globals.canvas = canvasapi.Canvas(func.cleanupURL(globals.config["API_URL"]), globals.config["API_TOKEN"])
    globals.user_data = globals.canvas.get_user("self")
    if globals.user_data is not None:
        print("\nWelcome {}.".format(globals.user_data.short_name))
        try:
            server_data = server.user_login() # Logs the user, and returns the new version number
            github_data = server.check_version()
            if version != github_data["tag_name"]: #server_data["version"]:
                resp = messagebox.askquestion("Announcement","A new version of PEER is available! You should download the latest version for the most recent bug fixes and new features.\n\nCurrent version: {}\nNew version: {}\n\nDo you want to download it now?".format(version, github_data["tag_name"]))
                if resp == "yes": func.open_file(github_data["html_url"])
        except:
            pass

if globals.config["API_TOKEN"] == "" or globals.config["API_URL"] == "":
	print("\nNo configuration found. Please go to \"File -> Settings\" to set up a link to Canvas.")  
else:
	print("\nPlease click on the \"Peer Evaluation\" menu to begin the evaluation process.")

top_frame.mainloop()

sys.stdout = sys.__stdout__
sys.stderr = sys.__stderr__