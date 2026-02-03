"""
Flask application for Guest Service Recovery Report Generator.
Equivalent to the entire Node.js/TypeScript application.
"""

from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
import os
import tempfile
from io import BytesIO

from app.excel_parser import parse_excel_from_bytes
from app.pdf_generator import generate_pdf
from app.report_utils import (
    ParsedComplaint,
    get_month_name,
)

# Initialize Flask app
app = Flask(__name__, template_folder='../templates', static_folder='../static')
CORS(app)

# Configuration
UPLOAD_FOLDER = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE


def allowed_file(filename):
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def filter_complaints_by_date(
    complaints, start_date: datetime, end_date: datetime
):
    """Filter complaints by date range."""
    end_date = end_date.replace(hour=23, minute=59, second=59)
    return [
        c for c in complaints
        if start_date <= c.date_time <= end_date
    ]


@app.route('/')
def index():
    """Serve the main page."""
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle Excel file upload and parsing."""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Please upload an Excel file.'}), 400

        # Read file bytes
        file_bytes = file.read()

        # Parse Excel file
        complaints, errors = parse_excel_from_bytes(file_bytes)

        if not complaints:
            return jsonify({
                'error': 'Failed to parse Excel file',
                'details': errors
            }), 400

        # Store complaints in session (in production, use database)
        # For now, return the parsed data
        return jsonify({
            'success': True,
            'complaints': [c.to_dict() for c in complaints],
            'count': len(complaints),
            'errors': errors if errors else []
        }), 200

    except Exception as e:
        return jsonify({'error': f'Upload failed: {str(e)}'}), 500


@app.route('/api/generate-report', methods=['POST'])
def generate_report():
    """Generate report with filtered complaints."""
    try:
        data = request.get_json()

        if not data or 'complaints' not in data:
            return jsonify({'error': 'No complaints data provided'}), 400

        # Parse complaints from request
        complaints_data = data.get('complaints', [])
        complaints = []

        for c in complaints_data:
            try:
                complaint = ParsedComplaint(
                    date_time=datetime.fromisoformat(c['dateTime']),
                    guest_name=c['guestName'],
                    room=c['room'],
                    problem_area=c['problemArea'],
                    confirmation_no=c.get('confirmationNo'),
                    membership_tier=c.get('membershipTier'),
                    complaint_details=c.get('complaintDetails'),
                    action_taken=c.get('actionTaken'),
                    fd_staff=c.get('fdStaff'),
                    follow_up_required=c.get('followUpRequired'),
                    follow_up_date=datetime.fromisoformat(c['followUpDate']) if c.get('followUpDate') else None,
                    follow_up_staff=c.get('followUpStaff'),
                    follow_up_comments=c.get('followUpComments'),
                )
                complaints.append(complaint)
            except Exception as e:
                print(f"Error parsing complaint: {e}")
                continue

        if not complaints:
            return jsonify({'error': 'No valid complaints to process'}), 400

        # Get date range
        start_date_str = data.get('startDate')
        end_date_str = data.get('endDate')

        if not start_date_str or not end_date_str:
            return jsonify({'error': 'Date range is required'}), 400

        start_date = datetime.fromisoformat(start_date_str)
        end_date = datetime.fromisoformat(end_date_str)

        # Filter complaints by date
        filtered_complaints = filter_complaints_by_date(complaints, start_date, end_date)

        if not filtered_complaints:
            return jsonify({'error': 'No complaints found in the specified date range'}), 400

        # Generate PDF
        pdf_bytes = generate_pdf(filtered_complaints, start_date, end_date)

        # Return PDF as file
        pdf_filename = f'Guest Service Recovery Report {start_date.strftime("%m-%d-%Y")} to {end_date.strftime("%m-%d-%Y")}.pdf'
        return send_file(
            BytesIO(pdf_bytes),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=pdf_filename
        )

    except Exception as e:
        return jsonify({'error': f'Report generation failed: {str(e)}'}), 500


@app.route('/api/months', methods=['GET'])
def get_months():
    """Get available months for 2026."""
    current_date = datetime.now()
    months = []

    for month in range(1, 13):
        month_date = datetime(2026, month, 1)
        if month_date <= current_date:
            months.append({
                'value': month,
                'label': f'{get_month_name(month)} 2026'
            })

    return jsonify(months), 200


@app.route('/api/health', methods=['GET'])
def health():
    """Health check endpoint."""
    return jsonify({'status': 'ok'}), 200


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
