# Peer Evaluation Enhancement Resource
 Peer Evaluation App

Version 0.0.1: Fixed bug, regarding "PointsPossible" property of the "config" dictionary
Version 0.0.2: Fixed bug with repeated "other" list, and made the uploading of marks more robust.
Version 0.0.3: Fixed a bug that caused the program to crash if there was an empty group (under "Export Groups").
Version 0.0.4 (2020-02-20): Only weekdays are now counted when calculating late penalties.
Version 0.0.5 (2020-02-21): Fixed code to automatically hide grades from the students. In the SaveData file, the rater appears before the receiver.

Version 0.1.0 (2020-02-22): Default bulk messages are now editable in the 'default.txt' and 'config.txt' files.
Version 0.1.1 (2020-04-14): Fixed a bug in which the program crashed when downloading groups with a member who had unenrolled from a course.
Version 0.1.2 (2020-05-13): Fixed a bug causing the program to crash while exporting students when dealing with multiple group sets.

Version 0.2.0 (2020-05-27): PEER now comes with its own installer. Introduced an "Open App Folder" button under Settings.

Version 0.3.0 (2020-06-19): new features:
* Major revision, including faster uploading/downloading of data from Canvas using the GraphQL API, and revamping of the bulk mailer feature.
* New menu system in the main window. Simplification of the "Settings" section (email address no longer required.)
* Courses/units are now presented from most recent to oldest.
* Added hover-over captions to describe different parts of the interface.

Version 0.4.0 (2020-06-20), new features: 
* Added Policy settings, allowing the uploaded Peer Mark column to be published/unpublished, and to count only weekdays when calculating late penalties. 
* Fixed a bug causing the Group message feature to send out messages to individual members, rather than the entire group.

Version 0.5.0 (2020-06-22), new features:
* Added a field picker to the Bulk Mailer.
* Added an error handling system to the GraphQL query system to make it more robust to Internet connectivity issues.
* Added the Peer survey template file as part of the distribution.

Version 0.5.1 (2020-06-23): PEER issues a warning message if you try to run the application within a ZIP folder.

Version 0.6.0 (2020-07-19), new features: 
* Removed PEER's dependence on pandas and numpy packages, significantly reducing the size of the program. 
* Added the option to apply a penalty for partial completion of the peer evaluation survey.
* Added hover-over text to many of the options.

Version 0.6.1 (2020-08-01): Revised code to remove incompatibilities with MacOS.

Version 0.6.2 (2020-09-15): Fixed a bug causing PEER to crash in the bulk mailer if a unit contained no group sets.

Version 0.6.3 (2020-10-11): Fixed a bug that caused PEER to add students without groups to the Contact list.

Version 0.6.4 (2020-10-21): Fixed a bug that caused PEER to crash on the "Upload Peer Marks" screen if a unit did not have any assignments.

Version 0.6.5 (2020-11-11): Fixed a bug that caused PEER to crash when uploading Peer Marks if students had used extended character sets when writing their feedback.

Version 0.7.0 (2020-11-26), new features: 
* Provided options for how PEER should handle cases where the minimum number of peer raters is not met, including "Excuse", substitute with average rating, or assign score of 0. 
* PEER now uses Canvas' native late penalties function.
* In the Bulk Mailer, the "Fields" button is now disabled if the user has not yet nominated a distribution list file (if that is the method of sending). The bulk mailer now reports an error if you try to use a distribution list with duplicate IDs.
* Changed terminology for the scoring method ("Moderated Mark" instead of "Adjusted Mark").
* The exported XLSX file now includes the group details in the "Ratings" spreadsheet.
* PEER now adds the Peer Evaluation to the same assignment group as the team assessment
* Using the moderated scoring option, PEER now includes an option to allow you to stop a score being adjusted if the adjustment falls under a nominated %.
* PEER now allows you to set the Points associated with the PEER evaluation on the "Calculate Peer Marks" window.

Version 0.7.1 (2021-03-30): Fixed a bug causing PEER to crash when attempting to send a bulk message to a student who has unenrolled from a course.

Version 0.7.2 (2021-04-22): Fixed a bug causing PEER to crash when creating the exported spreadsheet of peer marks when a student had unenrolled from a course.

Version 0.8.0 (2021-05-07): new features:
* When a Class List is exported, the list now contains the URL of the group home page for each student.
* A new online logging feature has been established that monitors usage of PEER.
* Automatic notifications when a new version of PEER is available.
Bug fixes:
* Feedback to the students now correctly reflects the rescaled scores.
* Resolved a problem when an email distribution list had blank rows in it, which were falsely detected as duplicates.

Version 0.9.0 (2021-05-14)
New features:
* The scoring options now allow you to create a feedback-only peer mark (i.e., if enabled, the peer mark does not count towards students' final grades).
Bug fixes:
* PEER does not allow users to upload the peer marks unless the total score for the peer mark is greater than zero.
* PEER reports an error message if you try to upload peer marks in a unit in which you are not the convenor.
* PEER reports an error message if the peer rating data cannot be found.

Version 0.9.1 (2021-05-31)
Bug fixes:
* Made PEER more robust when producing the XLSX output file.
* Fixed a minor bug causing a rater's name not to appear in the XLSX file.
* Fixed a minor bug in which the "Upload Marks" button was greyed out.
* The default scoring policy is now set to the Average Mark method.

Version 0.9.2 (2021-06-02)
Bug fixes:
* Made PEER more robust against inconsistencies in group member when uploading data.
* Fixed the "Upload Marks" page so that only group assignments can be selected from the dropdown box.
* Provided an update to the survey template that extends the survey expiry time.

Version 0.9.3 (2021-06-24)
Bug fixes:
* Fixed a bug causing PEER to crash when exporting groups with only a single member in them. PEER now issues a warning if this situation is detected.
* Fixed a bug that occurs causing PEER not to overwrite an exported Contacts List file.

Version 0.9.4 (2021-11-02)
Bug fixes:
* Fixed a bug causing PEER to crash when uploading student results if a group assignment had not been set up in Canvas.
* Revised the description of what happens to students when they are Excused. The description now reads: "[Y]ou were excused from this assessment. Please note, you have not been penalised. Your final grade will be based on the weighted average of your other assignments."
* Adjusted the terminology for scoring methods. The "Teamwork Mark" now refers to the average rating received scoring method. The "Moderated Group Mark" now refers to the adjusted method.

Version 0.9.5 (2021-11-10)
Bug fixes:
* Fixed a small bug in PEER causing it to crash on some installations of Canvas.
* Revised code to incorporate GraphQL in the "Upload Marks" section, making PEER faster.
* The upper limit for the Points the peer mark is worth has been removed.

Version 0.9.6 (2022-02-09)
Bug fixes:
* Fixed a bug in PEER causing it not to read the first row of an XLSX file exported from Qualtrics.

Version 0.10.0 (2022-03-15)
New features:
* PEER now correctly downloads the email addresses of students
* Added compatibility for LimeSurvey
* Greater modularisation of PEER
Bug fixes:
* Fixed a bug causing PEER to crash while parsing a TSV/CSV file.
