# RGPV Result WebScraper

A comprehensive web application for fetching and analyzing RGPV (Rajiv Gandhi Proudyogiki Vishwavidyalaya) student results with support for both regular and diploma students.

## 🌟 Features

### Core Functionality
- **Batch Result Fetching**: Download results for multiple students in one go
- **Individual SGPA Search**: Search SGPA data across all semesters for a specific enrollment number
- **Diploma Student Support**: Special handling for diploma students with different enrollment formats
- **Old Result Access**: Access previously stored results from local files
- **Multi-format Export**: Export results in both Excel (.xlsx) and PDF formats

### Data Analysis & Visualization
- **Automated Grade Analysis**: Calculates pass/fail counts for each subject
- **SGPA/CGPA Distribution**: Visual pie charts showing grade distributions
- **Statistical Reports**: Comprehensive analysis with student performance metrics
- **Excel Integration**: Embedded charts and formatted spreadsheets

### User Interface
- **Responsive Design**: Modern, mobile-friendly interface
- **Real-time Progress**: Live progress tracking during result fetching
- **Interactive Forms**: Dynamic form fields based on user selections
- **Error Handling**: Comprehensive error messages and validation

## 📸 Screenshots

### Main Interface
<img width="1349" height="639" alt="image" src="https://github.com/user-attachments/assets/a35291aa-b801-4c6a-b452-9c4a9fa8a7f3" />



### SGPA Search Interface
<img width="1366" height="638" alt="image" src="https://github.com/user-attachments/assets/c9433d61-ecc2-41f2-be7e-51e4b9bd6ddd" />



### Results Dashboard
<img width="1352" height="436" alt="image" src="https://github.com/user-attachments/assets/db56cef1-0fc3-4d1b-bdba-23498c22dc28" />



## 🚀 Getting Started

### Prerequisites

- Python 3.7+
- Chrome browser (for Selenium WebDriver)
- Tesseract OCR

### Required Dependencies

```bash
pip install flask selenium pandas matplotlib reportlab openpyxl requests pytesseract pillow
```

### System Requirements

1. **Tesseract OCR**: Download and install from [GitHub](https://github.com/tesseract-ocr/tesseract)
   - Update the path in `app.py` (line 38):
   ```python
   pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
   ```

2. **ChromeDriver**: Ensure Chrome browser is installed (ChromeDriver will be automatically managed)

### Installation

1. Clone or download the project
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Update Tesseract path in `app.py`
4. Run the application:
   ```bash
   python app.py
   ```
5. Open your browser and navigate to `http://localhost`

## 📚 Usage Guide

### Main Result Fetching

1. **Basic Information**:
   - Enter College Code (e.g., `0105`)
   - Select Branch (CSE, IT, ME, etc.)
   - Choose Batch (2021-25, 2022-26, etc.)
   - Select Semester (1-8)

2. **Student Type Selection**:
   - **Regular Students**: Standard enrollment format
   - **Diploma Students**: Check the diploma checkbox and select admission semester (3 or 5)

3. **Result Source**:
   - **New Results**: Fetch fresh data from RGPV servers
   - **Old Results**: Access previously stored local files

4. **Roll Number Range**:
   - Enter starting and ending roll numbers
   - Example: Start: 1, End: 50

### SGPA Search

1. Navigate to SGPA Search page
2. Enter 12-digit enrollment number
3. Check diploma student if applicable
4. System will search across all available semesters

### Understanding Enrollment Formats

- **Regular Students**: `0133CS231047`
  - Format: `[College][Branch][Year]1[Roll]`
- **Diploma Students**: `0133CS243D21`
  - Format: `[College][Branch][Year][AdmissionSem]D[Roll]`

## 🏗️ Project Structure

```
WebScraper/
├── app.py          # Main Flask application and core logic
├── templates/
│   ├── 1st.html           # Main result fetching interface
│   └── sgpa.html          # SGPA search interface
├── [College][Branch]/      # Directory structure for stored results
│   └── *.csv              # Semester-wise result files
├── *.xlsx                 # Generated Excel files
├── *.pdf                  # Generated PDF files
├── *_pie_*.png           # Generated chart images
└── README.md             # This file
```

## 🔧 Key Functions

### Core Processing Functions

- **`resultFound()`**: Main function for fetching new results from RGPV servers
- **`handle_old_result()`**: Processes locally stored result files
- **`fetch_sgpa_by_enrollment()`**: Retrieves SGPA data for individual students
- **`generate_enrollment_numbers()`**: Creates enrollment numbers for both regular and diploma students

### Data Processing Functions

- **`makeXslx()`**: Converts CSV data to Excel format with embedded charts
- **`makePdfFromExcel()`**: Generates PDF reports from Excel files
- **`calculate_grades_and_averages()`**: Analyzes grade distributions
- **`fail_counter()`**: Calculates pass/fail statistics

### Utility Functions

- **`readFromImage()`**: OCR-based captcha solving
- **`writeCSV()`**: Structured CSV data writing
- **`safe_delete()`**: Secure file cleanup

## 📊 Output Formats

### Excel Files (.xlsx)
- Formatted spreadsheets with student data
- Embedded SGPA and CGPA distribution charts
- Pass/fail statistics for each subject
- 1-based row indexing for user convenience

### PDF Files (.pdf)
- Professional report layout
- Tabular data presentation
- Integrated statistical charts
- Landscape orientation for better readability

### CSV Files (.csv)
- Raw data in comma-separated format
- Includes grade analysis and statistics
- Compatible with external data analysis tools

## 🎯 Special Features

### Diploma Student Support
- Automatic enrollment format adjustment
- Year calculation correction (year - 1 for SGPA search)
- Admission semester handling (3rd and 5th semester options)

### Progress Tracking
- Real-time progress updates during batch processing
- Current enrollment number display
- Percentage completion indicators

### File Management
- Automatic file cleanup after 2 minutes
- Secure temporary file handling
- Organized directory structure

## ⚙️ Configuration

### Branch Codes
- `CS`: Computer Science Engineering
- `IT`: Information Technology
- `ME`: Mechanical Engineering
- `al`: CS-AIML (Artificial Intelligence & Machine Learning)
- `cd`: CS-DS (Computer Science & Data Science)
- `EC`: Electronics & Communication
- `EX`: Electronics Engineering
- `au`: Automobile Engineering

### Batch Years
- `21`: 2021-25
- `22`: 2022-26
- `23`: 2023-27
- `24`: 2024-28

## 🛠️ Troubleshooting

### Common Issues

1. **Tesseract Path Error**:
   - Verify Tesseract installation
   - Update path in `app.py`

2. **ChromeDriver Issues**:
   - Ensure Chrome browser is updated
   - Check internet connectivity

3. **File Not Found Errors**:
   - Verify directory structure
   - Check file permissions

4. **Captcha Recognition Failures**:
   - Network connectivity issues
   - OCR accuracy limitations (system will retry)

## 📝 API Endpoints

- `GET /`: Main result fetching interface
- `GET /sgpa`: SGPA search interface
- `POST /submit`: Process result fetching requests
- `POST /submit-sgpa`: Process SGPA search requests
- `GET /progress/<job_id>`: Get processing progress
- `GET /download/<job_id>`: Download generated files

## 🚨 Disclaimer

This tool is developed for educational purposes only. It does not misuse any data and is not officially affiliated with RGPV (Rajiv Gandhi Proudyogiki Vishwavidyalaya).

## 👥 Developed By:

**Developed and Maintained by:**
Arshit Rawat, Department of Computer Science & Engineering, O.I.S.T, Bhopal

## 📄 License

This project is intended for educational use and internal purposes within the academic institution.

## 🔄 Version History

- **Latest Version**: Enhanced diploma student support with dynamic form fields
- **Previous Updates**: 
  - Added SGPA search functionality
  - Implemented real-time progress tracking
  - Enhanced data visualization with charts
  - Improved error handling and user experience

---

**Note**: Ensure you have proper permissions and follow your institution's guidelines when using this tool for accessing student result data.
