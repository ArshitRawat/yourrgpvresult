import os
import time
import uuid
import requests
import threading
import pandas as pd
from io import BytesIO
from flask import Flask, render_template, request, send_file, jsonify
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoAlertPresentException
from PIL import Image
import pytesseract
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
import csv
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image as PdfImage
from reportlab.lib.pagesizes import A4, landscape
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage

app = Flask(__name__)
base_dir = os.path.dirname(os.path.abspath(__file__))
pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe"

task_progress = {}
job_files = {}

def fail_counter(csv_path):
    df = pd.read_csv(f'{csv_path}.csv')
    subject_columns = df.columns[2:-3]
    f_counts = []
    pass_counts = []
    for column in subject_columns:
        f_count = (df[column] == 'F').sum()
        pass_count = len(df) - f_count
        f_counts.append(f_count)
        pass_counts.append(pass_count)
    f_count_row = [''] + ['F Count'] + f_counts + [''] * (len(df.columns) - len(f_counts) - 2)

    pass_count_row = [''] + ['PASS Count'] + pass_counts + [''] * (len(df.columns) - len(pass_counts) - 2)
    df.loc[len(df)] = f_count_row
    df.loc[len(df)] = pass_count_row
    df.to_csv(f'{csv_path}.csv', index=False)

    print("Updated CSV file with 'F' and 'PASS' counts has been saved.")


def safe_delete(path):
    try:
        if os.path.exists(path):
            os.remove(path)
            print(f"[✓] Deleted: {path}")
    except Exception as e:
        print(f"[!] Failed to delete {path}: {e}")

def writeCSV(enroll, name, *args, sgpa, cgpa, remark, filename):
    gradesString = [str(a) + "," for a in args]
    filepath = os.path.join(base_dir, filename)
    information = [enroll, ",", name, ","] + gradesString + [sgpa, ",", cgpa, ",", remark, "\n"]
    with open(filepath, 'a') as f:
        f.writelines(information)

def show_count(pct, allvals):
    total = sum(allvals)
    count = int(round(pct * total / 100.0))
    return f'{count}' if count > 0 else ''

def makeXslx(filename, job_id):
    csv_path = os.path.join(base_dir, f"{filename}.csv")
    excel_path = os.path.join(base_dir, f"{filename}.xlsx")

    # Read the CSV to a DataFrame and ensure it has the correct structure
    try:
        df = pd.read_csv(csv_path, skip_blank_lines=False)
        print("CSV loaded successfully!")
    except Exception as e:
        print(f"Error loading CSV: {e}")
        return None

    # Check if the DataFrame is correctly loaded and has no missing columns
    if df.empty:
        print("[!] DataFrame is empty!")
        return None
    print(df.head())  # Check first few rows to confirm the data structure

    # Save the DataFrame as an Excel file using pandas
    try:
        df.to_excel(excel_path, index=True)  # Save without the index column
        print(f"Excel file saved at: {excel_path}")
    except Exception as e:
        print(f"Error saving Excel file: {e}")
        return None

    # Now add images to the Excel file
    try:
        wb = load_workbook(excel_path)  # Open the Excel file
        ws = wb.active

        # Corrected image paths
        img_sgpa_path = os.path.join(base_dir, f'sgpa_pie_{job_id}.png')
        img_cgpa_path = os.path.join(base_dir, f'cgpa_pie_{job_id}.png')

        # Check if images exist before adding them
        if not os.path.exists(img_sgpa_path):
            print(f"[!] SGPA image not found: {img_sgpa_path}")
        if not os.path.exists(img_cgpa_path):
            print(f"[!] CGPA image not found: {img_cgpa_path}")
        
        # Add images if they exist
        if os.path.exists(img_sgpa_path):
            img_sgpa = OpenpyxlImage(img_sgpa_path)
            ws.add_image(img_sgpa, 'K2')  # Adjust the position as needed
        if os.path.exists(img_cgpa_path):
            img_cgpa = OpenpyxlImage(img_cgpa_path)
            ws.add_image(img_cgpa, 'K33')

        # Save the workbook with the images added
        wb.save(excel_path)
        print(f"Excel file with images saved at: {excel_path}")
    except Exception as e:
        print(f"Error adding images or saving workbook: {e}")
        return None

    return excel_path


def calculate_grades_and_averages(csv_path):
    s_grades = [0, 0, 0, 0, 0]
    c_grades = [0, 0, 0, 0, 0]
    s_total = 0
    s_count = 0
    c_total = 0
    c_count = 0

    try:
        with open(csv_path, 'r') as file:
            reader = csv.reader(file)
            for row in reader:
                try:
                    val = float(row[-3])
                    c_val = float(row[-2])
                except (IndexError, ValueError):
                    continue

                # For s_grades (using val)
                if val < 4.0:
                    s_grades[0] += 1
                elif val < 5:
                    s_grades[1] += 1
                elif val < 6.5:
                    s_grades[2] += 1
                elif val < 7.5:
                    s_grades[3] += 1 
                else:
                    s_grades[4] += 1

                # For c_grades (using c_val)
                if c_val < 4.0:
                    c_grades[0] += 1
                elif c_val < 5:
                    c_grades[1] += 1
                elif c_val < 6.5:
                    c_grades[2] += 1
                elif c_val < 7.5:
                    c_grades[3] += 1
                else:
                    c_grades[4] += 1

                s_total += val
                s_count += 1
                c_total += c_val
                c_count += 1
                
        return s_grades, c_grades, s_count, c_count
    except Exception as e:
        print("Error in grade calculation:", e)
        return None, None, None, 0, 0


def makePdfFromExcel(excel_path, pdf_path,job_id):

    df = pd.read_excel(excel_path)
    data = [df.columns.tolist()] + df.fillna('').astype(str).values.tolist()

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=landscape(A4),
        rightMargin=10, leftMargin=10, topMargin=15, bottomMargin=15
    )
    page_width = landscape(A4)[0] - 20
    num_columns = len(data[0]) if data else 1
    col_widths = []
    for col_idx in range(num_columns):
        max_content_length = 0
        for row in data:
            if col_idx < len(row):
                content_length = len(str(row[col_idx]))
                max_content_length = max(max_content_length, content_length)
        content_width = max(40, min(120, max_content_length * 4.5))
        col_widths.append(content_width)
    total_width = sum(col_widths)
    if total_width > page_width:
        scale_factor = page_width / total_width
        col_widths = [width * scale_factor for width in col_widths]

    table = Table(data, repeatRows=1, colWidths=col_widths)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),  
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),      # Default alignment CENTER for all cells
        ('ALIGN', (2, 0), (2, -1), 'LEFT'),         # Override first column to LEFT
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),            
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 6),                
        ('BOTTOMPADDING', (0, 0), (-1, 0), 2),              
        ('TOPPADDING', (0, 0), (-1, -1), 1),                
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),             
        ('LEFTPADDING', (0, 0), (-1, -1), 1),               
        ('RIGHTPADDING', (0, 0), (-1, -1), 1),              
        ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
        ('WORDWRAP', (0, 0), (-1, -1), True),               
    ]))
    # Add the table to the PDF
    elements = [table]

    # Add images (SGPA and CGPA pie charts) to the PDF
    img_sgpa_path = os.path.join(base_dir, f'sgpa_pie_{job_id}.png')
    img_cgpa_path = os.path.join(base_dir, f'cgpa_pie_{job_id}.png')

    # Add the SGPA image
    if os.path.exists(img_sgpa_path):
        sgpa_img = PdfImage(img_sgpa_path, width=3*inch, height=3*inch)
        elements.append(sgpa_img)

    # Add the CGPA image
    if os.path.exists(img_cgpa_path):
        cgpa_img = PdfImage(img_cgpa_path, width=3*inch, height=3*inch)
        elements.append(cgpa_img)

    # Build the PDF with the table and images
    doc.build(elements)

def generate_enrollment_numbers(college, branch, year, start, end):
    """Generate list of enrollment numbers for the given range"""
    enrollments = []
    for num in range(start, end + 1):
        num_str = f"{num:03d}"
        enroll = f"{college}{branch}{year}1{num_str}"
        print(f"Enroll: {enroll}")
        enrollments.append(enroll)
    return enrollments

def filter_csv_by_enrollments(source_csv, target_csv, enrollments):
    """Filter existing CSV file by specific enrollment numbers"""
    try:
        if not os.path.exists(source_csv):
            return False, "Source CSV file not found"
        df = pd.read_csv(source_csv)
        first_col = df.columns[0]
        filtered_df = df[df[first_col].isin(enrollments)]
        
        if filtered_df.empty:
            return False, "No matching enrollment numbers found in existing file"
        filtered_df.to_csv(target_csv, index=False)     
        return True, f"Found {len(filtered_df)} records out of {len(enrollments)} requested"        
    except Exception as e:
        return False, f"Error filtering CSV: {str(e)}"

def handle_old_result(college, branch, year, sem, start, end, job_id):
    """Handle old result functionality - extract specific roll numbers from existing CSV"""
    base_filename = f'{college}{branch}{year}sem{sem}'
    subfolder = f"{college}{branch}"
    source_csv = os.path.join(base_dir, subfolder, f"{base_filename}.csv")
    target_filename = f'{college}{branch}{year}sem{sem}_{job_id}'
    target_csv = os.path.join(base_dir, f"{target_filename}.csv")
    enrollments = generate_enrollment_numbers(college, branch, year, start, end)

    task_progress[job_id] = {
        "done": 0, 
        "total": len(enrollments), 
        "current_enroll": "Searching existing file..."
    }
    success, message = filter_csv_by_enrollments(source_csv, target_csv, enrollments)
    
    if success:
        task_progress[job_id]["done"] = len(enrollments)
        task_progress[job_id]["current_enroll"] = message
        try:
            fail_counter(target_filename)
            s_grades, c_grades, s_count, c_count = calculate_grades_and_averages(target_csv)
            with open(target_csv, 'a', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(['SGPA'])
                for i in range(len(s_grades)):
                    if (i == 0):
                        writer.writerow([f'Below 4.0', s_grades[i]])
        
                    elif (i == 1):
                        writer.writerow([f'Between 4.0-4.99', s_grades[i]])
                    elif (i == 2):
                        writer.writerow([f'Between 5.0-6.49', s_grades[i]])
                    elif (i == 3):
                        writer.writerow([f'Between 6.5-7.49', s_grades[i]])
                    else:
                        writer.writerow([f'Above 7.5', s_grades[i]])
                    
                writer.writerow(['CGPA'])
                for i in range(len(s_grades)):
                    if (i == 0):
                        writer.writerow([f'Below 4.0', s_grades[i]])
        
                    elif (i == 1):
                        writer.writerow([f'Between 4.0-4.99', s_grades[i]])
                    elif (i == 2):
                        writer.writerow([f'Between 5.0-6.49', s_grades[i]])
                    elif (i == 3):
                        writer.writerow([f'Between 6.5-7.49', s_grades[i]])
                    else:
                        writer.writerow([f'Above 7.5', s_grades[i]])

            # Make sure the labels only match the grades that have non-zero counts
            labels = ['Below 4.0', '4.0 - 4.99', '5.0 - 6.49', '6.5 - 7.49', 'Above 7.5']
            
            filtered_s_grades = [g for g in s_grades if g > 0]
            filtered_c_grades = [g for g in c_grades if g > 0]

            # Only include the labels corresponding to non-zero grades
            filtered_s_labels = [label for label, grade in zip(labels, s_grades) if grade > 0]
            filtered_c_labels = [label for label, grade in zip(labels, c_grades) if grade > 0]

            # Colors for the pie chart
            blue_shades = ['#004c6d', '#346888', '#5886a5', '#7aa6c2', '#9dc6e0']
            red_shades = ["#700D0D", "#9b3a3a", "#c54242", '#ff6666', '#ff9999']

            b_colors = blue_shades[:len(filtered_s_grades)]
            r_colors = red_shades[:len(filtered_c_grades)]

            # SGPA Pie Chart
            plt.figure(figsize=(4, 4))
            wedges, texts, autotexts = plt.pie(
                filtered_s_grades,
                labels=filtered_s_labels,
                colors=b_colors,
                autopct=lambda pct: show_count(pct, filtered_s_grades),
                startangle=90
            )
            for text in autotexts:
                text.set_color('white')
                text.set_fontsize(12)
            plt.title('SGPA Distribution (Student Counts)', fontsize=14)
            plt.axis('equal')
            plt.tight_layout()
            plt.savefig(f'sgpa_pie_{job_id}.png')
            plt.close()
            print("Pie chart saved as sgpa_pie.png.")

            # CGPA Pie Chart
            plt.figure(figsize=(4, 4))
            wedges, texts, autotexts = plt.pie(
                filtered_c_grades,
                labels=filtered_c_labels,
                colors=r_colors,
                autopct=lambda pct: show_count(pct, filtered_c_grades),
                startangle=90
            )
            for text in autotexts:
                text.set_color('white')
                text.set_fontsize(12)
            plt.title('CGPA Distribution (Student Counts)', fontsize=14)
            plt.axis('equal')
            plt.tight_layout()
            plt.savefig(f'cgpa_pie_{job_id}.png')
            plt.close()
            print("Pie chart saved as cgpa_pie.png.")

            # Create the Excel file after grade calculations
            xlsx_path = makeXslx(target_filename, job_id)
            job_files[job_id] = xlsx_path
            print(f"[+] Old result job {job_id} complete. Created: {xlsx_path}")
            return True
        except Exception as e:
            print(f"[!] Error creating Excel: {e}")
            return False
    else:
        print(f"[!] Old result failed for job {job_id}: {message}")
        task_progress[job_id]["error"] = message
        return False

def readFromImage(url: str) -> str:
    try:
        response = requests.get(url, timeout=10)
        if response.status_code != 200:
            return ""
        img = Image.open(BytesIO(response.content))
        text = pytesseract.image_to_string(img)
        return text.upper().replace(" ", "").strip()
    except Exception as e:
        print(f"OCR Error: {e}")
        return ""

def resultFound(start, end, college, branch, year, sem, job_id):
    if branch not in ["CS", "IT", "ME", "AL", "CD", "EC", "EX","AU"]:
        return

    task_progress[job_id] = {"done": 0, "total": end - start + 1, "current_enroll": ""}
    noResult = []
    print(f"[~] Starting job {job_id} for {start} to {end}")
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920x1080")

    driver = webdriver.Chrome(options=chrome_options)
    filename_base = f'{college}{branch}{year}sem{sem}_{job_id}'
    filename_csv = f'{filename_base}.csv'
    driver.implicitly_wait(0.5)
    driver.get("http://result.rgpv.ac.in/Result/ProgramSelect.aspx")
    driver.find_element(By.ID, "radlstProgram_1").click()

    firstRow = True
    while start <= end:
        num = f"{start:03d}"
        enroll = f"{college}{branch}{year}1{num}"
        task_progress[job_id]["current_enroll"] = enroll

        try:
            img_element = driver.find_element(By.XPATH, '//img[contains(@src, "CaptchaImage.axd")]')
            img_src = img_element.get_attribute("src")
            url = f'http://result.rgpv.ac.in/result/{img_src.split("Result/")[-1]}'
            captcha = readFromImage(url)

            Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_drpSemester")).select_by_value(str(sem))
            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_TextBox1").send_keys(captcha)
            time.sleep(1)
            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtrollno").send_keys(enroll)
            time.sleep(2)
            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnviewresult").send_keys(Keys.ENTER)
            time.sleep(2)
            try:
                alert = Alert(driver)
                alerttext = alert.text
                alert.accept()
            except NoAlertPresentException:
                alerttext = ""

            if "Total Credit" in driver.page_source:
                if firstRow:
                    headers = []
                    rows = driver.find_elements(By.CSS_SELECTOR, "table.gridtable tbody tr")
                    for row in rows:
                        cells = row.find_elements(By.TAG_NAME, "td")
                        if len(cells) >= 4:
                            headers.append(cells[0].text.strip('- [T]'))
                    writeCSV("Enrollment No.", "Name", *headers, sgpa="SGPA", cgpa="CGPA", remark="REMARK", filename=filename_csv)
                    firstRow = False

                name = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblNameGrading").text
                grades = []
                rows = driver.find_elements(By.CSS_SELECTOR, "table.gridtable tbody tr")
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 4:
                        grades.append(cells[3].text.strip())

                sgpa = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblSGPA").text
                cgpa = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcgpa").text
                result = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblResultNewGrading").text
                writeCSV(enroll, name, *grades, sgpa=sgpa, cgpa=cgpa, remark=result.replace(",", " "), filename=filename_csv)

                driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnReset").send_keys(Keys.ENTER)
                task_progress[job_id]["done"] += 1
                start += 1
            else:
                if "Result" in alerttext:
                    driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnReset").send_keys(Keys.ENTER)
                    task_progress[job_id]["done"] += 1
                    start += 1
                    noResult.append(enroll)
                else:
                    driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_TextBox1").clear()
                    driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtrollno").clear()
                    continue
        except Exception as e:
            print(f"Error: {e}")
            start += 1

    task_progress[job_id]["current_enroll"] = ""
    driver.quit()
    columns_to_remove = [2, 3, 4]
    df = pd.read_csv(f'{filename_csv}')
    df = df.drop(df.columns[columns_to_remove], axis=1)
    df.to_csv(f'{filename_csv}', index=False)
    fail_counter(filename_base)
    s_grades, c_grades, s_count, c_count = calculate_grades_and_averages(filename_csv)
    with open(filename_csv, 'a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['SGPA'])  # Blank row
        for i in range(len(s_grades)):
            if (i == 0):
                writer.writerow([f'Below 4.0', s_grades[i]])
        
            elif (i == 1):
                writer.writerow([f'Between 4.0-4.99', s_grades[i]])
            elif (i == 2):
                writer.writerow([f'Between 5.0-6.49', s_grades[i]])
            elif (i == 3):
                writer.writerow([f'Between 6.5-7.49', s_grades[i]])
            else:
                writer.writerow([f'Above 7.5', s_grades[i]])
                    
        writer.writerow(['CGPA'])
        for i in range(len(s_grades)):
            if (i == 0):
                writer.writerow([f'Below 4.0', s_grades[i]])
        
            elif (i == 1):
                writer.writerow([f'Between 4.0-4.99', s_grades[i]])
            elif (i == 2):
                writer.writerow([f'Between 5.0-6.49', s_grades[i]])
            elif (i == 3):
                writer.writerow([f'Between 6.5-7.49', s_grades[i]])
            else:
                writer.writerow([f'Above 7.5', s_grades[i]])

            # Make sure the labels only match the grades that have non-zero counts
    labels = ['Below 4.0', '4.0 - 4.99', '5.0 - 6.49', '6.5 - 7.49', 'Above 7.5']
            
    filtered_s_grades = [g for g in s_grades if g > 0]
    filtered_c_grades = [g for g in c_grades if g > 0]

            # Only include the labels corresponding to non-zero grades
    filtered_s_labels = [label for label, grade in zip(labels, s_grades) if grade > 0]
    filtered_c_labels = [label for label, grade in zip(labels, c_grades) if grade > 0]

            # Colors for the pie chart
    blue_shades = ['#004c6d', '#346888', '#5886a5', '#7aa6c2', '#9dc6e0']
    red_shades = ["#700D0D", "#9b3a3a", "#c54242", '#ff6666', '#ff9999']

    b_colors = blue_shades[:len(filtered_s_grades)]
    r_colors = red_shades[:len(filtered_c_grades)]

            # SGPA Pie Chart
    plt.figure(figsize=(4, 4))
    wedges, texts, autotexts = plt.pie(
                filtered_s_grades,
                labels=filtered_s_labels,
                colors=b_colors,
                autopct=lambda pct: show_count(pct, filtered_s_grades),
                startangle=90
            )
    for text in autotexts:
        text.set_color('white')
        text.set_fontsize(12)
    plt.title('SGPA Distribution (Student Counts)', fontsize=14)
    plt.axis('equal')
    plt.tight_layout()
    plt.savefig(f'sgpa_pie_{job_id}.png')
    plt.close()
    print("Pie chart saved as sgpa_pie.png.")

            # CGPA Pie Chart
    plt.figure(figsize=(4, 4))
    wedges, texts, autotexts = plt.pie(
                filtered_c_grades,
                labels=filtered_c_labels,
                colors=r_colors,
                autopct=lambda pct: show_count(pct, filtered_c_grades),
                startangle=90
            )
    for text in autotexts:
        text.set_color('white')
        text.set_fontsize(12)
    plt.title('CGPA Distribution (Student Counts)', fontsize=14)
    plt.axis('equal')
    plt.tight_layout()
    plt.savefig(f'cgpa_pie_{job_id}.png')
    plt.close()
    print("Pie chart saved as cgpa_pie.png.")

    xlsx_path = makeXslx(filename_base,job_id)
    job_files[job_id] = xlsx_path
    print(f"[+] Job {job_id} complete. Writing XLSX to: {xlsx_path}")


@app.route('/')
def form():
    return render_template('1st.html')

@app.route('/submit', methods=['POST'])
def submit():
    job_id = str(uuid.uuid4())
    data = request.form
    use_old_result = 'old_result' in data and data['old_result'] == 'on' 
    college = data['college']
    branch = data['branch'].upper()
    year = data['year']
    sem = int(data['sem'])
    if use_old_result:
        if 'start' not in data or 'end' not in data or not data['start'] or not data['end']:
            return jsonify({"error": "Start and end roll numbers are required even for existing results"})
        
        start = int(data['start'])
        end = int(data['end'])
        def old_result_thread():
            handle_old_result(college, branch, year, sem, start, end, job_id)
        
        threading.Thread(target=old_result_thread).start()
        return jsonify({"job_id": job_id, "old_result": True})
    else:
        start = int(data['start'])
        end = int(data['end'])
        
        threading.Thread(target=resultFound, args=(
            start, end, college, branch, year, sem, job_id)).start()
        return jsonify({"job_id": job_id})

def schedule_file_deletion(job_id):
    """ Schedule the deletion of files after 2 minutes """
    pie_sgpa_path = os.path.join(base_dir, f'sgpa_pie_{job_id}.png')
    pie_cgpa_path = os.path.join(base_dir, f'cgpa_pie_{job_id}.png')
    threading.Timer(120.0, lambda: safe_delete(pie_sgpa_path)).start()
    threading.Timer(120.0, lambda: safe_delete(pie_cgpa_path)).start()

@app.route('/progress/<job_id>')
def get_progress(job_id):
    progress = task_progress.get(job_id, {"done": 0, "total": 1, "current_enroll": ""})
    if "error" in progress:
        progress["error"] = progress["error"]
    
    return jsonify(progress)

@app.route('/download/<job_id>')
def download(job_id):
    file_type = request.args.get('type', 'xlsx')
    xlsx_path = job_files.get(job_id)

    if not xlsx_path or not os.path.exists(xlsx_path):
        return "File not found or not ready.", 404

    base_name = os.path.splitext(xlsx_path)[0]
    csv_path = base_name + ".csv"
    pdf_path = base_name + ".pdf"

    # Schedule deletion of all 3 files after 2 minutes
    threading.Timer(120.0, lambda: safe_delete(xlsx_path)).start()
    threading.Timer(120.0, lambda: safe_delete(csv_path)).start()
    threading.Timer(120.0, lambda: safe_delete(pdf_path)).start()
    schedule_file_deletion(job_id)
    
    if file_type == 'pdf':
        if not os.path.exists(pdf_path):
            makePdfFromExcel(xlsx_path, pdf_path,job_id)
        return send_file(
            pdf_path,
            as_attachment=True,
            download_name=os.path.basename(pdf_path),
            mimetype='application/pdf'
        )
    return send_file(
        xlsx_path,
        as_attachment=True,
        download_name=os.path.basename(xlsx_path),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(debug=False, host="0.0.0.0", port=80) 
