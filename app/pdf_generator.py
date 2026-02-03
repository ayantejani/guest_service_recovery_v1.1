"""
PDF generator for complaint reports.
Equivalent to server/pdfGenerator.ts
"""

from typing import List, Dict, Any
from datetime import datetime
from weasyprint import HTML, CSS
from io import BytesIO
from app.report_utils import (
    ParsedComplaint,
    calculate_metrics,
    calculate_staff_performance,
    calculate_problem_areas,
    calculate_membership_tiers,
    get_complaint_status,
    format_date,
    get_report_period_label,
    get_staff_color,
)


def generate_report_html(
    complaints: List[ParsedComplaint],
    start_date: datetime,
    end_date: datetime,
) -> str:
    """Generate HTML report from complaints data."""

    metrics = calculate_metrics(complaints, end_date)
    staff_performance = calculate_staff_performance(complaints)
    problem_areas = calculate_problem_areas(complaints)
    membership_tiers = calculate_membership_tiers(complaints)

    # Sort complaints by date descending
    sorted_complaints = sorted(complaints, key=lambda x: x.date_time, reverse=True)
    
    # Filter and sort active and overdue cases (Overdue first, then Active)
    active_overdue_complaints = [c for c in complaints if get_complaint_status(c, end_date)[0] in ["Active", "Overdue"]]
    active_overdue_complaints.sort(key=lambda x: (get_complaint_status(x, end_date)[0] != "Overdue", x.date_time), reverse=True)
    
    # Calculate active and overdue cases per staff member
    staff_case_status = {}
    for complaint in complaints:
        staff = complaint.fd_staff or "Unassigned"
        if staff not in staff_case_status:
            staff_case_status[staff] = {"active": False, "overdue": False}
        
        status, _, _ = get_complaint_status(complaint, end_date)
        if status == "Active":
            staff_case_status[staff]["active"] = True
        elif status == "Overdue":
            staff_case_status[staff]["overdue"] = True

    # Build HTML
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            * {{
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }}
            
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                color: #333;
                background-color: white;
                line-height: 1.3;
                font-size: 11px;
            }}
            
            .container {{
                width: 100%;
                margin: 0;
                padding: 20px;
                background-color: white;
            }}
            
            .header {{
                text-align: center;
                margin-bottom: 20px;
                border-bottom: 3px solid #0078d4;
                padding-bottom: 10px;
            }}
            
            .header h1 {{
                font-size: 26px;
                color: #0078d4;
                margin-bottom: 2px;
                font-weight: bold;
            }}
            
            .header p {{
                font-size: 13px;
                color: #666;
                margin: 0;
            }}
            
            .report-period {{
                text-align: center;
                font-size: 11px;
                color: #666;
                margin-bottom: 20px;
            }}
            
            .metrics {{
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 12px;
                margin-bottom: 25px;
            }}
            
            .metric-card {{
                background-color: #f5f5f5;
                border-left: 4px solid #0078d4;
                padding: 12px;
                border-radius: 2px;
            }}
            
            .metric-card.closed {{
                border-left-color: #28a745;
            }}
            
            .metric-card.active {{
                border-left-color: #ffc107;
            }}
            
            .metric-card.overdue {{
                border-left-color: #dc3545;
            }}
            
            .metric-label {{
                font-size: 9px;
                color: #666;
                text-transform: uppercase;
                margin-bottom: 5px;
                font-weight: 600;
            }}
            
            .metric-value {{
                font-size: 22px;
                font-weight: bold;
                color: #0078d4;
            }}
            
            .metric-card.closed .metric-value {{
                color: #28a745;
            }}
            
            .metric-card.active .metric-value {{
                color: #ffc107;
            }}
            
            .metric-card.overdue .metric-value {{
                color: #dc3545;
            }}
            
            .section {{
                margin-bottom: 25px;
                page-break-inside: auto;
            }}
            
            .section:last-of-type {{
                page-break-inside: auto;
            }}
            
            .section-title {{
                font-size: 14px;
                font-weight: bold;
                color: #0078d4;
                margin-bottom: 10px;
                padding-bottom: 8px;
                border-bottom: 2px solid #0078d4;
            }}
            
            table {{
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 15px;
            }}
            
            tr {{
                page-break-inside: avoid;
            }}
            
            th {{
                background-color: #0078d4;
                color: white;
                padding: 8px;
                text-align: left;
                font-weight: 600;
                font-size: 12px;
                border: 1px solid #0078d4;
                page-break-inside: avoid;
            }}
            
            td {{
                padding: 12px;
                border: 1px solid #ddd;
                font-size: 11px;
                vertical-align: top;
            }}
            
            tbody tr:nth-child(even) {{
                background-color: #f0f4f8;
            }}
            
            tbody tr:nth-child(odd) {{
                background-color: #ffffff;
            }}
            
            .status-badge {{
                display: inline-block;
                padding: 4px 8px;
                border-radius: 14px;
                font-size: 10px;
                font-weight: 700;
                text-transform: uppercase;
                white-space: normal;
                word-break: break-word;
                max-width: 100%;
            }}
            
            .status-completed {{
                background-color: #28a745;
                color: white;
            }}
            
            .status-active {{
                background-color: #ffc107;
                color: #333;
            }}
            
            .status-overdue {{
                background-color: #dc3545;
                color: white;
            }}
            
            .percentage {{
                font-weight: bold;
                color: #0078d4;
            }}
            
            .page-break {{
                page-break-after: auto;
            }}
            
            .overdue-details {{
                font-size: 9px;
                color: #666;
                margin-top: 3px;
            }}
            
            .attention-badge {{
                display: inline-block;
                background-color: #dc3545;
                color: white;
                padding: 2px 5px;
                border-radius: 3px;
                font-size: 9px;
                font-weight: bold;
                margin-left: 5px;
            }}
            
            /* Pill styles */
            .detail-pills {{
                display: flex;
                flex-direction: column;
                gap: 8px;
            }}
            
            .pill {{
                display: inline-block;
                padding: 4px 8px;
                border-radius: 14px;
                font-size: 10px;
                font-weight: 600;
                color: #333;
                width: fit-content;
            }}
            
            .pill-date {{
                background-color: #e8e8e8;
                color: #333;
            }}
            
            .pill-crn {{
                background-color: transparent;
                border: 2px solid #0078d4;
                color: #0078d4;
            }}
            
            .pill-room {{
                background-color: transparent;
                border: 2px solid #6f42c1;
                color: #6f42c1;
            }}
            
            .pill-tier {{
                background-color: transparent;
                border: 2px solid #e83e8c;
                color: #e83e8c;
            }}
            
            .pill-problem {{
                background-color: transparent;
                border: 2px solid #ffc107;
                color: #ffc107;
            }}
            
            .pill-staff {{
                color: #333;
                background-color: transparent;
                border: none;
            }}
            
            .pill-overdue-status {{
                background-color: #dc3545;
                color: white;
                border: none;
            }}
            
            .pill-active-status {{
                background-color: #ffc107;
                color: #333;
                border: none;
            }}
            
            /* Column widths for detailed incidents */
            .detailed-table th:nth-child(1),
            .detailed-table td:nth-child(1) {{ width: 14%; }}
            
            .detailed-table th:nth-child(2),
            .detailed-table td:nth-child(2) {{ width: 18%; }}
            
            .detailed-table th:nth-child(3),
            .detailed-table td:nth-child(3) {{ width: 24%; }}
            
            .detailed-table th:nth-child(4),
            .detailed-table td:nth-child(4) {{ width: 24%; }}
            
            .detailed-table th:nth-child(5),
            .detailed-table td:nth-child(5) {{ width: 20%; }}
        </style>
    </head>
    <body>
        <div class="container">
            <!-- Header -->
            <div class="header">
                <h1>Holiday Inn Express Markham</h1>
                <p>Guest Service Recovery Report</p>
            </div>
            
            <div class="report-period">
                Report Period: {get_report_period_label(start_date, end_date)}
            </div>
            
            <!-- Metrics -->
            <div class="metrics">
                <div class="metric-card">
                    <div class="metric-label">Total Incidents</div>
                    <div class="metric-value">{metrics['total']}</div>
                </div>
                <div class="metric-card closed">
                    <div class="metric-label">Completed</div>
                    <div class="metric-value">{metrics['completed']} ({metrics['completed_percentage']}%)</div>
                </div>
                <div class="metric-card active">
                    <div class="metric-label">Active</div>
                    <div class="metric-value">{metrics['active']}</div>
                </div>
                <div class="metric-card overdue">
                    <div class="metric-label">Overdue</div>
                    <div class="metric-value">{metrics['overdue']}</div>
                </div>
            </div>
            
            <!-- Staff Performance -->
            <div class="section">
                <h2 class="section-title">Staff Performance</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Staff Name</th>
                            <th>Number of Complaints</th>
                        </tr>
                    </thead>
                    <tbody>
    """

    for staff in staff_performance:
        staff_name = staff['name']
        status_pills = ""
        
        if staff_name in staff_case_status:
            if staff_case_status[staff_name]["overdue"]:
                status_pills += '<span class="pill pill-overdue-status">has overdue cases</span>'
            if staff_case_status[staff_name]["active"]:
                status_pills += '<span class="pill pill-active-status">has active cases</span>'
        
        pills_html = f'<span style="margin-left: 12px;">{status_pills}</span>' if status_pills else ""
        
        html += f"""
                        <tr>
                            <td>{staff_name}</td>
                            <td style="display: flex; align-items: center; gap: 8px;">{staff['count']}{pills_html}</td>
                        </tr>
        """

    html += """
                    </tbody>
                </table>
            </div>
            
            <!-- Problem Area Breakdown -->
            <div class="section">
                <h2 class="section-title">Problem Area Breakdown</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Problem Area</th>
                            <th>Count</th>
                            <th>Percentage</th>
                            <th>Rooms</th>
                        </tr>
                    </thead>
                    <tbody>
    """

    for area_data in problem_areas:
        html += f"""
                        <tr>
                            <td>{area_data['area']}</td>
                            <td>{area_data['count']}</td>
                            <td><span class="percentage">{area_data['percentage']}%</span></td>
                            <td>{area_data['rooms']}</td>
                        </tr>
        """

    html += """
                    </tbody>
                </table>
            </div>
            
            <!-- Member Tier Analysis -->
            <div class="section">
                <h2 class="section-title">Member Tier Analysis</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Member Tier</th>
                            <th>Number of Complaints</th>
                        </tr>
                    </thead>
                    <tbody>
    """

    for tier_data in membership_tiers:
        tier = tier_data['tier']
        count = tier_data['count']
        needs_attention = ' <span class="attention-badge">Needs Attention!</span>' if tier_data['needsAttention'] else ''
        html += f"""
                        <tr>
                            <td>{tier}{needs_attention}</td>
                            <td>{count}</td>
                        </tr>
        """

    html += """
                    </tbody>
                </table>
            </div>
    """

    # Active and Overdue Cases (only if there are any)
    if active_overdue_complaints:
        html += """
            <!-- Active and Overdue Cases -->
            <div class="section">
                <h2 class="section-title">Active and Overdue Cases</h2>
                <table class="detailed-table">
                    <thead>
                        <tr>
                            <th>Date & Details</th>
                            <th>Guest Name</th>
                            <th>Complaint Details</th>
                            <th>Action Taken</th>
                            <th>Staff & Status</th>
                        </tr>
                    </thead>
                    <tbody>
        """

        for complaint in active_overdue_complaints:
            status, days_overdue, message = get_complaint_status(complaint, end_date)
            status_class = f"status-{status.lower()}"
            
            status_text = status
            follow_up_date_display = ""
            if status == "Overdue" and days_overdue is not None:
                status_text = f"Overdue by {days_overdue} Days<br/>"
                follow_up_date_str = format_date(complaint.follow_up_date) if complaint.follow_up_date else "-"
                follow_up_date_display = f'<div class="overdue-details">Follow Up date was {follow_up_date_str}</div>'
            
            # Build date & details pills
            date_details_html = '<div class="detail-pills">'
            date_details_html += f'<span class="pill pill-date">{format_date(complaint.date_time)}</span>'
            date_details_html += f'<span class="pill pill-crn">CRN# {complaint.confirmation_no or "-"}</span>'
            date_details_html += f'<span class="pill pill-room">Room {complaint.room}</span>'
            if complaint.membership_tier:
                date_details_html += f'<span class="pill pill-tier">{complaint.membership_tier}</span>'
            date_details_html += f'<span class="pill pill-problem">{complaint.problem_area}</span>'
            date_details_html += '</div>'
            
            # Build staff & status section
            staff_status_html = '<div class="detail-pills">'
            staff_status_html += f'<div style="font-weight: 600; margin-bottom: 4px;">{complaint.fd_staff or "-"}</div>'
            staff_status_html += f'<span class="status-badge {status_class}">{status_text}</span>'
            if message and status != "Overdue":
                staff_status_html += f'<div class="overdue-details">{message}</div>'
            if follow_up_date_display:
                staff_status_html += follow_up_date_display
            staff_status_html += '</div>'
            
            html += f"""
                        <tr>
                            <td>{date_details_html}</td>
                            <td>{complaint.guest_name}</td>
                            <td>{complaint.complaint_details or '-'}</td>
                            <td>{complaint.action_taken or '-'}</td>
                            <td>{staff_status_html}</td>
                        </tr>
            """

        html += """
                    </tbody>
                </table>
            </div>
        """
            
    html += """
            
            <!-- Detailed Incidents -->
            <div class="section">
                <h2 class="section-title">Detailed Incident Report</h2>
                <table class="detailed-table">
                    <thead>
                        <tr>
                            <th>Date & Details</th>
                            <th>Guest Name</th>
                            <th>Complaint Details</th>
                            <th>Action Taken</th>
                            <th>Staff & Status</th>
                        </tr>
                    </thead>
                    <tbody>
    """

    for complaint in sorted_complaints:
        status, days_overdue, message = get_complaint_status(complaint, end_date)
        status_class = f"status-{status.lower()}"
        
        status_text = status
        follow_up_date_display = ""
        if status == "Overdue" and days_overdue is not None:
            status_text = f"Overdue by {days_overdue} Days<br/>"
            follow_up_date_str = format_date(complaint.follow_up_date) if complaint.follow_up_date else "-"
            follow_up_date_display = f'<div class="overdue-details">Follow Up date was {follow_up_date_str}</div>'
        
        # Build date & details pills
        date_details_html = '<div class="detail-pills">'
        date_details_html += f'<span class="pill pill-date">{format_date(complaint.date_time)}</span>'
        date_details_html += f'<span class="pill pill-crn">CRN# {complaint.confirmation_no or "-"}</span>'
        date_details_html += f'<span class="pill pill-room">Room {complaint.room}</span>'
        if complaint.membership_tier:
            date_details_html += f'<span class="pill pill-tier">{complaint.membership_tier}</span>'
        date_details_html += f'<span class="pill pill-problem">{complaint.problem_area}</span>'
        date_details_html += '</div>'
        
        # Build staff & status section
        staff_status_html = '<div class="detail-pills">'
        staff_status_html += f'<div style="font-weight: 600; margin-bottom: 4px;">{complaint.fd_staff or "-"}</div>'
        staff_status_html += f'<span class="status-badge {status_class}">{status_text}</span>'
        if message and status != "Overdue":
            staff_status_html += f'<div class="overdue-details">{message}</div>'
        if follow_up_date_display:
            staff_status_html += follow_up_date_display
        staff_status_html += '</div>'
        
        html += f"""
                        <tr>
                            <td>{date_details_html}</td>
                            <td>{complaint.guest_name}</td>
                            <td>{complaint.complaint_details or '-'}</td>
                            <td>{complaint.action_taken or '-'}</td>
                            <td>{staff_status_html}</td>
                        </tr>
        """

    html += """
                    </tbody>
                </table>
            </div>
        </div>
    </body>
    </html>
    """
    
    return html


def generate_pdf(
    complaints: List[ParsedComplaint],
    start_date: datetime,
    end_date: datetime,
) -> bytes:
    """Generate PDF from complaints data."""
    html_content = generate_report_html(complaints, start_date, end_date)
    
    # Convert HTML to PDF using WeasyPrint
    pdf_bytes = HTML(string=html_content).write_pdf()
    return pdf_bytes
