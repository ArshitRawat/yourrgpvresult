===============================
📘 RGPV LOCAL RESULT FETCHING
===============================

To use this system for LOCAL result fetching (without live scraping), follow the instructions below:

---------------------------------------
🔹 STEP 1: Prepare Your CSV File
---------------------------------------

Create a CSV file with the following **exact format and order**:

ROLL.NO,NAME,BT101,BT102,BT103,BT104,BT105,BT101,BT103,BT104,BT105,BT106,BT108,SGPA,CGPA,RESULT  
0105CS211001,ENTERNAMEABCDEF,B+,B,B+,A+,A,B+,A,A,B+,B+,B+,8.19,8.19,PASS  
0105CS211002,ENTERNAMEUVXYZ,A,B,A,A+,A,A,A,A,A,B+,A,8.67,8.67,PASS  
...  
(continue for all students in the batch)

✔️ Make sure:
- Column headers are present in the first row.
- There are no missing columns or data mismatches.
- SGPA and CGPA are numeric (e.g., 8.19), and RESULT is usually "PASS" or "FAIL".

---------------------------------------
🔹 STEP 2: Name Your CSV File
---------------------------------------

The CSV file should be saved with the following filename format:

  {collegecode}{branch}{year}sem{sem}.csv


📌 Examples:
- For college code **0105**, branch **CS**, year **21**, semester **1**  
  ➤ `0105CS21sem1.csv`

- For college code **0413**, branch **IT**, year **20**, semester **3**  
  ➤ `0413IT20sem3.csv`

❗ This naming format is mandatory for the application to detect and process your file correctly.

---------------------------------------
🔹 STEP 3: File Location
---------------------------------------

Place the CSV file inside the following subdirectory:

  📂 {collegecode}{branch}

📌 For example:
- CSV: `0105CS21sem1.csv`
- Folder: `0105CS`
- Full path: `0105CS/0105CS21sem1.csv`

Create the folder manually if it doesn't already exist.

---------------------------------------
✅ You're Ready!

Now run the application and choose "Old Result" mode.
The system will read the local CSV file and generate all analysis, charts, and reports.

Enjoy automated grade analytics! 🎓

-----------------------------------------------------------------------------------------
Created by Arshit Rawat UNDER Department Of Computer Science & Engineering O.I.S.T, Bhopal
------------------------------------------------------------------------------------------
