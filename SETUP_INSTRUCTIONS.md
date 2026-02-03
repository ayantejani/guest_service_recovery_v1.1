# Holiday Inn Express Markham - Guest Service Recovery Report Generator

## Complete Setup Guide

### System Requirements
- Python 3.7 or higher
- pip (Python package manager)
- 500MB free disk space
- Modern web browser (Chrome, Firefox, Safari, Edge)

### Installation Steps

#### 1. Extract the Application
```bash
unzip guest-report-app-complete.zip
cd guest-report-app-python
```

#### 2. Create Virtual Environment (Recommended)
```bash
# On Windows
python -m venv venv
venv\Scripts\activate

# On macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

#### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

**Note:** If you encounter issues with WeasyPrint on Windows, you may need to install additional dependencies:
- Download and install GTK+ from: https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer
- Then retry: `pip install -r requirements.txt`

#### 4. Run the Application
```bash
python run.py
```

You should see output like:
```
 * Running on http://127.0.0.1:5000
 * Press CTRL+C to quit
```

#### 5. Access the Application
Open your web browser and go to:
```
http://localhost:5000
```

### Usage Instructions

#### Uploading Excel File
1. Click the upload area or drag-and-drop your Excel file
2. Supported format: `.xlsx` (Excel 2007+)
3. File must have headers in row 3 with the following columns:
   - Date
   - Time
   - Guest Name (First Name Last Name)
   - Room
   - Confirmation no
   - Membership Tier
   - Problem Area
   - Complaint Details
   - Action Taken
   - FD Staff
   - Follow-Up-Required
   - Follow-Up Date
   - Follow up Staff
   - Follow Up Comments

#### Selecting Date Range
- **Option 1:** Use "From Date" and "To Date" pickers to select a custom date range
- **Option 2:** Select a specific month from the "Select Month" dropdown

#### Generating Report
1. After uploading the file, select your preferred date range or month
2. Click "Generate Report"
3. The report will display with:
   - Executive summary (Total, Completed, Active, Overdue incidents)
   - Staff performance analytics
   - Problem area breakdown
   - Member tier analysis
   - Active and Overdue cases (if any exist)
   - Detailed incident report with all complaints

#### Downloading PDF
1. Click "Download PDF" button
2. The PDF will be named: `Guest Service Recovery Report [DATE_RANGE].pdf`
3. PDF includes all report sections with professional formatting

### Report Sections Explained

#### Executive Summary
- **Total Incidents:** All complaints in the selected period
- **Completed:** Cases with Follow-Up-Required = "No" OR Follow-Up Date in past with staff assigned
- **Active:** Cases with Follow-Up-Required = "Yes" AND Follow-Up Date in future OR not set
- **Overdue:** Cases with Follow-Up-Required = "Yes" AND Follow-Up Date in past with NO staff assigned

#### Staff Performance
Shows count of complaints handled by each FD Staff member

#### Problem Area Breakdown
Lists all problem categories with:
- Count of incidents
- Percentage of total
- Room numbers affected

#### Member Tier Analysis
Shows complaints by membership tier (Diamond, Platinum, Gold, Silver, Club, Non Members)
- Tiers with >10 complaints marked as "Needs Attention!" in red

#### Active and Overdue Cases
Filtered view showing only:
- Active cases (pending follow-up)
- Overdue cases (follow-up past due)
Sorted by status (Overdue first)

#### Detailed Incident Report
Complete table with all complaints including:
- Date & Details (Date, CRN#, Room, Tier, Problem Area as colored pills)
- Guest Name
- Complaint Details
- Action Taken
- Staff & Status (Staff name and status badge)

### Troubleshooting

#### Port Already in Use
If port 5000 is already in use, edit `run.py` and change:
```python
app.run(debug=True, port=5000)
```
to:
```python
app.run(debug=True, port=5001)  # or any available port
```

#### WeasyPrint Installation Issues
If WeasyPrint fails to install on Windows:
1. Download GTK+ Runtime: https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer
2. Install GTK+ to default location
3. Retry pip install

#### Excel File Not Parsing
- Ensure headers are in row 3
- Check that required columns exist (Date, Time, Guest Name, Room, Problem Area)
- Verify date format is recognized (YYYY-MM-DD or MM/DD/YYYY)

#### PDF Generation Fails
- Ensure all required Python packages are installed: `pip install -r requirements.txt`
- Try restarting the application
- Check that you have sufficient disk space for temporary files

### File Structure
```
guest-report-app-python/
├── app/
│   ├── __init__.py
│   ├── app.py                 # Flask application
│   ├── excel_parser.py        # Excel file parsing
│   ├── pdf_generator.py       # PDF generation
│   └── report_utils.py        # Data processing utilities
├── templates/
│   └── index.html             # Web interface
├── static/
│   ├── css/                   # Stylesheets
│   └── js/                    # JavaScript files
├── requirements.txt           # Python dependencies
├── run.py                     # Application entry point
└── README.md                  # Documentation
```

### Key Features

✅ Drag-and-drop Excel file upload
✅ Date range and month-based filtering
✅ Executive summary with key metrics
✅ Staff performance analytics
✅ Problem area breakdown with room numbers
✅ Member tier analysis with alerts
✅ Active and Overdue cases tracking
✅ Professional PDF reports with color coding
✅ Alternating row colors for readability
✅ Responsive web interface
✅ No data stored - all processing is local

### Support

For issues or questions:
1. Check the troubleshooting section above
2. Verify your Excel file format matches the requirements
3. Ensure all dependencies are installed: `pip install -r requirements.txt`
4. Restart the application: `python run.py`

### License
This application is provided for Holiday Inn Express Markham internal use.

---
**Version:** 1.0
**Last Updated:** January 2026
