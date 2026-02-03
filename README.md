# Guest Service Recovery Report Generator - Python Version

A complete Python conversion of the Holiday Inn Express Markham complaint management system. This application allows staff to upload Excel complaint data and generate professional PDF reports with analytics.

## Features

- **Excel File Upload** - Drag-and-drop interface for uploading complaint data
- **Date Range Filtering** - Select custom date ranges or specific months
- **Executive Summary** - Total incidents, closed cases, active cases, and overdue cases
- **Staff Performance Analytics** - Aggregated complaint handling counts by staff member
- **Problem Area Breakdown** - Categorized complaints with percentages and room numbers
- **Detailed Incident Report** - Complete table of all complaints with status badges
- **PDF Export** - Professional PDF reports with Holiday Inn Express branding

## Project Structure

```
guest-report-app-python/
├── app/
│   ├── __init__.py
│   ├── app.py                 # Flask application and API routes
│   ├── excel_parser.py        # Excel file parsing logic
│   ├── pdf_generator.py       # PDF generation from HTML
│   └── report_utils.py        # Data processing utilities
├── templates/
│   └── index.html             # Main web interface
├── static/
│   ├── css/
│   └── js/
├── requirements.txt           # Python dependencies
├── run.py                     # Application entry point
└── README.md                  # This file
```

## Installation

### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

### Setup

1. **Clone or extract the project**
   ```bash
   cd guest-report-app-python
   ```

2. **Create a virtual environment (recommended)**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

1. **Start the Flask server**
   ```bash
   python3 run.py
   ```

2. **Open your browser**
   Navigate to `http://localhost:5000`

3. **Upload an Excel file**
   - Click the upload area or drag-and-drop your Excel file
   - The file must contain columns: Date, Time, Guest Name, Room, Problem Area, etc.

4. **Generate a report**
   - Select a date range or specific month
   - Click "Generate Report"
   - The PDF will download automatically

## API Endpoints

### POST /api/upload
Upload and parse an Excel file.

**Request:**
- Form data with `file` field containing the Excel file

**Response:**
```json
{
  "success": true,
  "complaints": [...],
  "count": 48,
  "errors": []
}
```

### POST /api/generate-report
Generate a PDF report from complaints data.

**Request:**
```json
{
  "complaints": [...],
  "startDate": "2026-01-01T00:00:00",
  "endDate": "2026-01-31T23:59:59"
}
```

**Response:**
- Binary PDF file

### GET /api/months
Get available months for 2026.

**Response:**
```json
[
  {"value": 1, "label": "January 2026"},
  {"value": 2, "label": "February 2026"},
  ...
]
```

### GET /api/health
Health check endpoint.

**Response:**
```json
{"status": "ok"}
```

## Excel File Format

The application expects an Excel file with the following columns:

| Column | Required | Description |
|--------|----------|-------------|
| Date | Yes | Complaint date |
| Time | Yes | Complaint time |
| Guest Name | Yes | Guest's full name |
| Room | Yes | Room number |
| Membership Tier | No | Guest membership level |
| Problem Area | Yes | Category of complaint |
| Complaint Details | No | Detailed description |
| Action Taken | No | Action taken to resolve |
| FD Staff | No | Staff member who handled it |
| Follow-Up-Required | No | Yes/No for follow-up needed |
| Follow-Up Date | No | Date for follow-up |
| Follow up Staff | No | Staff assigned for follow-up |
| Follow Up Comments | No | Additional notes |

## Report Output

The generated PDF includes:

1. **Header** - Holiday Inn Express Markham branding
2. **Report Period** - Date range of the report
3. **Executive Summary** - Key metrics with color-coded cards
4. **Staff Performance** - Table of complaints handled by each staff member
5. **Problem Area Breakdown** - Categorized complaints with percentages
6. **Detailed Incident Report** - Complete table with all complaint details and status badges

## Technologies Used

- **Flask** - Web framework
- **openpyxl** - Excel file parsing
- **WeasyPrint** - PDF generation
- **Jinja2** - HTML templating
- **CORS** - Cross-origin resource sharing

## Troubleshooting

### Port already in use
If port 5000 is already in use, modify the port in `run.py`:
```python
app.run(debug=True, host='0.0.0.0', port=5001)
```

### PDF generation fails
Ensure WeasyPrint dependencies are installed:
```bash
pip install --upgrade weasyprint
```

### Excel parsing errors
Check that your Excel file:
- Has headers in the first row
- Contains valid dates in the Date column
- Has required fields: Date, Time, Guest Name, Room, Problem Area

## Development

To modify the application:

1. **Update data processing** - Edit `app/report_utils.py`
2. **Modify Excel parsing** - Edit `app/excel_parser.py`
3. **Change PDF styling** - Edit `app/pdf_generator.py`
4. **Add API routes** - Edit `app/app.py`
5. **Update UI** - Edit `templates/index.html`

## License

This application is provided as-is for Holiday Inn Express Markham.

## Support

For issues or questions, please contact the development team.
