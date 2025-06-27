# 🎓 RGPV Result Scraper & Analyzer

A full-stack Python web application that automates the extraction, filtering, and analysis of student results from the official [RGPV](http://result.rgpv.ac.in/) portal. It supports both **live scraping** and **local CSV fetching**. The app generates Excel and PDF reports with SGPA/CGPA statistics and pie charts.

> ⚠️ Intended for academic and educational purposes only.

---

## 📌 Key Features

- 🔍 **Automated Result Scraping** using Selenium.
- 🔐 **Captcha Solving** with Tesseract OCR.
- 📊 **SGPA/CGPA Analytics** and Pie Charts.
- 📁 **CSV, Excel, and PDF Export**.
- 🧾 **Old Result Fetching** from local CSV files.
- 🖥️ **Real-time Progress Tracker** via Flask web interface.

---

## 📁 Folder & File Requirements

### 🔹 For Local CSV-Based Fetching:

You must organize CSV files in the following structure:

project/
├── 0105CS/
│ └── 0105CS21sem1.csv
├── 0413IT/
│ └── 0413IT20sem3.csv

- Folder name = `{college}{branch}` (e.g., `0105CS`)
- File name = `{college}{branch}{year}sem{sem}.csv`  
  (e.g., `0105CS21sem1.csv`)

> ✅ Make sure the CSV contains header + student rows exactly like this:

```
ROLL.NO,NAME,BT101,BT102,...,SGPA,CGPA,RESULT
0105CS211001,AAKASH PATEL,B+,B,...,8.19,8.19,PASS
...
```
⚙️ Installation & Setup
1. Clone the Repository

git clone https://github.com/yourusername/rgpv-result-scraper.git
cd rgpv-result-scraper

2. Install Required Packages

pip install -r requirements.txt

3. Install Tesseract OCR

Download Tesseract

After installation, update the following line in newscrapper.py:


pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
🚗 ChromeDriver Setup (for Selenium)

🔹 Step 1: Check Chrome Version

Go to chrome://settings/help in your browser to find your Chrome version.

🔹 Step 2: Download Matching ChromeDriver

Visit: https://chromedriver.chromium.org/downloads

Download the version matching your browser.

Extract it into a folder named driver/ in your project.


rgpv-result-scraper/
├── driver/
│   └── chromedriver.exe (Windows) or chromedriver (Linux/Mac)

🔹 Step 3: (Optional) Hardcode ChromeDriver Path

In newscrapper.py, change:
driver = webdriver.Chrome(options=chrome_options)
to:
chrome_path = os.path.join(base_dir, 'driver', 'chromedriver.exe')  # Adjust if Linux/Mac
driver = webdriver.Chrome(executable_path=chrome_path, options=chrome_options)

▶️ Running the Application

python newscrapper.py

Then open your browser and go to:

http://localhost

**INPUT EXAMPLE**
![image](https://github.com/user-attachments/assets/bc28a50e-d091-45da-9319-c4db3339d964)

📤 Output Examples
**OUTPUT PAGE**
![image](https://github.com/user-attachments/assets/0f4440f2-813b-4f90-bed2-429eadbcaea7)

**EXCEL PAGE**
![image](https://github.com/user-attachments/assets/7e2cd103-60fb-42c6-bbdc-3277ad4cf69f)

.– With embedded charts

**PDF  PAGE**
![image](https://github.com/user-attachments/assets/1d156aec-ccd3-4530-b5ad-8730fb2427bc)
![image](https://github.com/user-attachments/assets/3bf57a03-8735-4b80-81d1-bca047a0e365)


.pdf – Tabular layout with pie charts

Output files are temporarily stored and auto-deleted after a few minutes.

📈 Visual Example
(Optional screenshots section – add your own)

📋 Web Form UI

📊 SGPA/CGPA Pie Charts

📁 Excel File with Charts

📄 Downloadable PDF Report

💡 Use Cases
Department result audits

Batch-wise academic performance

Automated data collection for reports

📄 License
This project is provided for academic, research, and educational use only. Unauthorized scraping of live data may violate website terms.

✨ Created By
<p align="center"> <b>Arshit Rawat UNDER Department of Computer Science & Engineering O.I.S.T, BHOPAL</b> <br/> 💻 <a href="https://github.com/arshitrawat" target="_blank">github.com/arshitrawat</a> <br/> 📧 arshitrawat2704@gmail.com </p>
