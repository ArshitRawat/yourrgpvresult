===========================================
📘 SETTING UP CHROMEDRIVER FOR THE PROJECT
===========================================

To enable automated result fetching, this project uses **Selenium** with Google Chrome in headless mode.

You must download and configure **ChromeDriver** that matches your version of Google Chrome.

-----------------------------------------------
🔹 STEP 1: Check Your Chrome Version
-----------------------------------------------

1. Open Google Chrome.

2. Go to the address bar and type:
chrome://settings/help

3. Note the full version number (e.g., `114.0.5735.90`).

4. Only the **major version** matters (e.g., `114`).

-----------------------------------------------
🔹 STEP 2: Download the Right ChromeDriver
-----------------------------------------------

Go to the official ChromeDriver download page:

🔗 https://chromedriver.chromium.org/downloads

1. Find the version matching your Chrome (e.g., Chrome 114 → ChromeDriver 114).

2. Download the correct zip file for your operating system.

3. Extract the `chromedriver` (or `chromedriver.exe` on Windows).

-----------------------------------------------
🔹 STEP 3: Create a `driver` Folder
-----------------------------------------------

Inside your project directory, create a folder named:

🔹driver

Move the downloaded `chromedriver` file into this folder.

Your folder structure should now look like:

rgpv-result-scraper/
├── newscrapper.py
├── driver/
│ └── chromedriver.exe (on Windows)
│ └── chromedriver (on Mac/Linux)

------------------------------------------------------------------------------------------
Created by Arshit Rawat UNDER Department Of Computer Science & Engineering O.I.S.T, Bhopal
------------------------------------------------------------------------------------------
