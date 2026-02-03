"""
Report utilities for parsing Excel data and generating analytics.
Equivalent to server/reportUtils.ts
"""

from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any
import re


class ParsedComplaint:
    """Represents a parsed complaint record."""
    
    def __init__(
        self,
        date_time: datetime,
        guest_name: str,
        room: str,
        problem_area: str,
        confirmation_no: Optional[str] = None,
        membership_tier: Optional[str] = None,
        complaint_details: Optional[str] = None,
        action_taken: Optional[str] = None,
        fd_staff: Optional[str] = None,
        follow_up_required: Optional[str] = None,
        follow_up_date: Optional[datetime] = None,
        follow_up_staff: Optional[str] = None,
        follow_up_comments: Optional[str] = None,
    ):
        self.date_time = date_time
        self.guest_name = guest_name
        self.room = room
        self.problem_area = problem_area
        self.confirmation_no = confirmation_no
        self.membership_tier = membership_tier
        self.complaint_details = complaint_details
        self.action_taken = action_taken
        self.fd_staff = fd_staff
        self.follow_up_required = follow_up_required
        self.follow_up_date = follow_up_date
        self.follow_up_staff = follow_up_staff
        self.follow_up_comments = follow_up_comments

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return {
            'dateTime': self.date_time.isoformat(),
            'guestName': self.guest_name,
            'room': self.room,
            'confirmationNo': self.confirmation_no,
            'membershipTier': self.membership_tier,
            'problemArea': self.problem_area,
            'complaintDetails': self.complaint_details,
            'actionTaken': self.action_taken,
            'fdStaff': self.fd_staff,
            'followUpRequired': self.follow_up_required,
            'followUpDate': self.follow_up_date.isoformat() if self.follow_up_date else None,
            'followUpStaff': self.follow_up_staff,
            'followUpComments': self.follow_up_comments,
        }


def parse_excel_row(row: Dict[str, Any]) -> Optional[ParsedComplaint]:
    """
    Parse Excel row data into complaint object.
    Handles both the template format and actual HIEX format.
    """
    try:
        # Try to find date and time columns
        date_value = find_value(row, ["Date", "date", "DateTime", "dateTime"])
        time_value = find_value(row, ["Time", "time"])

        date_time = None
        if date_value is not None:
            if time_value is not None:
                date_time = combine_date_time(date_value, time_value)
            else:
                date_time = parse_date(date_value)

        if not date_time:
            return None

        # Find guest name - handle newlines in header
        guest_name = find_value_with_newlines(
            row,
            [
                "Guest Name",
                "guestName",
                "Guest Name (First Name Last Name)",
                "Guest Name\n(First Name Last Name)",
            ],
        )

        # Find room
        room = find_value(row, ["Room", "room"])

        # Find problem area
        problem_area = find_value(row, ["Problem Area", "problemArea", "Problem area"])

        if not guest_name or not room or not problem_area:
            return None

        # Find follow-up date
        follow_up_date_value = find_value(
            row, ["Follow-Up Date", "followUpDate", "Follow Up Date"]
        )
        follow_up_date = parse_date(follow_up_date_value) if follow_up_date_value else None

        # Find follow-up required field
        follow_up_required = find_value(
            row,
            [
                "Follow-Up-Required",
                "followUpRequired",
                "Follow-Up Required",
            ],
        )

        # Find FD Staff (primary staff tracking)
        fd_staff = find_value(
            row,
            [
                "FD Staff",
                "fdStaff",
                "FD Staff ",
            ],
        )

        return ParsedComplaint(
            date_time=date_time,
            guest_name=str(guest_name).strip(),
            room=str(room).strip(),
            confirmation_no=(
                str(find_value(row, ["Confirmation no", "confirmationNo"])).strip()
                if find_value(row, ["Confirmation no", "confirmationNo"])
                else None
            ),
            membership_tier=(
                str(find_value(row, ["Membership Tier", "membershipTier"])).strip()
                if find_value(row, ["Membership Tier", "membershipTier"])
                else None
            ),
            problem_area=str(problem_area).strip(),
            complaint_details=(
                str(find_value(row, ["Complaint Details", "complaintDetails"])).strip()
                if find_value(row, ["Complaint Details", "complaintDetails"])
                else None
            ),
            action_taken=(
                str(find_value(row, ["Action Taken", "actionTaken"])).strip()
                if find_value(row, ["Action Taken", "actionTaken"])
                else None
            ),
            fd_staff=(
                str(fd_staff).strip() if fd_staff else None
            ),
            follow_up_required=(
                str(follow_up_required).strip() if follow_up_required else None
            ),
            follow_up_date=follow_up_date,
            follow_up_staff=(
                str(
                    find_value(
                        row,
                        ["Follow up Staff", "followUpStaff", "Follow-up Staff"],
                    )
                ).strip()
                if find_value(row, ["Follow up Staff", "followUpStaff", "Follow-up Staff"])
                else None
            ),
            follow_up_comments=(
                str(find_value(row, ["Follow Up Comments", "followUpComments"])).strip()
                if find_value(row, ["Follow Up Comments", "followUpComments"])
                else None
            ),
        )
    except Exception as e:
        print(f"Error parsing Excel row: {e}")
        return None


def find_value(row: Dict[str, Any], possible_keys: List[str]) -> Any:
    """Find a value in a row by checking multiple possible key variations."""
    for key in possible_keys:
        if key in row and row[key] is not None:
            return row[key]

    # Try case-insensitive search
    row_keys = list(row.keys())
    for possible_key in possible_keys:
        for actual_key in row_keys:
            # Normalize both for comparison
            normalized_possible = (
                possible_key.replace("\r\n", " ")
                .replace("\n", " ")
                .replace("\r", " ")
                .replace(r"\s+", " ")
                .strip()
                .lower()
            )
            normalized_actual = (
                actual_key.replace("\r\n", " ")
                .replace("\n", " ")
                .replace("\r", " ")
                .replace(r"\s+", " ")
                .strip()
                .lower()
            )

            if normalized_possible == normalized_actual:
                return row[actual_key]

    return None


def find_value_with_newlines(row: Dict[str, Any], possible_keys: List[str]) -> Any:
    """Find a value handling keys with newlines."""
    row_keys = list(row.keys())

    for possible_key in possible_keys:
        # Try exact match first
        if possible_key in row:
            return row[possible_key]

        # Try normalized match
        normalized = (
            possible_key.replace("\r\n", " ")
            .replace("\n", " ")
            .replace("\r", " ")
            .replace(r"\s+", " ")
            .strip()
            .lower()
        )

        for actual_key in row_keys:
            actual_normalized = (
                actual_key.replace("\r\n", " ")
                .replace("\n", " ")
                .replace("\r", " ")
                .replace(r"\s+", " ")
                .strip()
                .lower()
            )

            if normalized == actual_normalized:
                return row[actual_key]

    return None


def combine_date_time(date_value: Any, time_value: Any) -> Optional[datetime]:
    """Combine date and time values into a single DateTime."""
    try:
        from datetime import time as time_type
        
        date = parse_date(date_value)
        if not date:
            return None

        # Parse time value
        hours = 0
        minutes = 0

        if isinstance(time_value, time_type):
            # Python time object from openpyxl
            hours = time_value.hour
            minutes = time_value.minute
        elif isinstance(time_value, (int, float)):
            # Excel stores time as a decimal fraction of a day
            total_minutes = round(time_value * 24 * 60)
            hours = total_minutes // 60
            minutes = total_minutes % 60
        elif isinstance(time_value, str):
            try:
                time_str = str(time_value).strip().replace('PM', '').replace('AM', '').strip()
                time_parts = time_str.split(":")
                hours = int(time_parts[0]) if len(time_parts) > 0 and time_parts[0].isdigit() else 0
                minutes = int(time_parts[1]) if len(time_parts) > 1 and time_parts[1].isdigit() else 0
            except:
                return date

        # Validate hours and minutes
        if hours > 23 or minutes > 59:
            return date

        date = date.replace(hour=hours, minute=minutes, second=0, microsecond=0)
        return date
    except Exception as e:
        print(f"Error combining date and time: {e}")
        return parse_date(date_value)


def parse_date(date_value: Any) -> Optional[datetime]:
    """Parse date string or number in various formats."""
    if not date_value:
        return None

    # If already a datetime object
    if isinstance(date_value, datetime):
        return date_value if is_valid_date(date_value) else None

    # If it's a number (Excel serial date)
    if isinstance(date_value, (int, float)):
        # Excel serial date: 1 = 1900-01-01
        excel_epoch = datetime(1900, 1, 1)
        days_offset = int(date_value) - 1
        date = excel_epoch + timedelta(days=days_offset)
        return date if is_valid_date(date) else None

    # Parse string dates
    try:
        date_str = str(date_value).strip()
        date = datetime.fromisoformat(date_str)
        return date if is_valid_date(date) else None
    except:
        try:
            # Try common date formats
            for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"]:
                try:
                    date = datetime.strptime(date_str, fmt)
                    return date if is_valid_date(date) else None
                except:
                    continue
        except:
            pass

    return None


def is_valid_date(date: datetime) -> bool:
    """Check if a date is valid."""
    return isinstance(date, datetime)


def get_complaint_status(complaint: ParsedComplaint, report_end_date: datetime) -> tuple[str, Optional[int], Optional[str]]:
    """
    Determine status badge for a complaint.
    Returns: (status_string, days_overdue_or_none, message_or_none)
    
    Logic:
    - Completed: Follow-up required is No
    - Active: Follow-up required is Yes AND Follow-up date is past AND follow-up staff assigned
      Message: "No Follow up comments recorded"
    - Active: Follow-up required is Yes AND Follow-up date is not set
      Message: "No Follow up Date Assigned"
    - Active: Follow-up required is blank
      Message: "No Follow up Details Recorded"
    - Active: Follow-up required is Yes AND Follow-up date is in the future
    - Overdue: Follow-up required is Yes AND Follow-up date is past AND NO follow-up staff
    """
    # Check if follow-up required is blank/None
    if complaint.follow_up_required is None or str(complaint.follow_up_required).strip() == "":
        return "Active", None, "No Follow up Details Recorded"
    
    follow_up_required = str(complaint.follow_up_required).lower().strip() in ["yes", "y"]

    # Completed: Follow-up required is No
    if not follow_up_required:
        return "Completed", None, None

    # Active: Follow-up required is Yes AND Follow-up date is not set
    if not complaint.follow_up_date:
        return "Active", None, "No Follow up Date Assigned"

    now = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    follow_up_date = complaint.follow_up_date.replace(hour=0, minute=0, second=0, microsecond=0)

    # Active: Follow-up required is Yes AND Follow-up date is in the future
    if follow_up_date > now:
        return "Active", None, None

    # Follow-up date is in the past or today
    # Active: Follow-up date is past AND follow-up staff name is assigned
    if complaint.follow_up_staff:
        return "Active", None, "No Follow up comments recorded"

    # Overdue: Follow-up required is Yes AND Follow-up date is past AND NO follow-up staff
    days_overdue = (now - follow_up_date).days
    return "Overdue", days_overdue, None


def calculate_metrics(complaints: List[ParsedComplaint], report_end_date: datetime) -> Dict[str, Any]:
    """Calculate report metrics from complaints."""
    total_complaints = len(complaints)

    closed_count = 0
    active_count = 0
    overdue_count = 0

    for complaint in complaints:
        status, _, _ = get_complaint_status(complaint, report_end_date)
        
        if status == "Completed":
            closed_count += 1
        elif status == "Active":
            active_count += 1
        elif status == "Overdue":
            overdue_count += 1

    return {
        "total": total_complaints,
        "completed": closed_count,
        "active": active_count,
        "overdue": overdue_count,
        "completed_percentage": (
            round((closed_count / total_complaints) * 100)
            if total_complaints > 0
            else 0
        ),
    }


def calculate_staff_performance(complaints: List[ParsedComplaint]) -> List[Dict[str, Any]]:
    """Calculate staff performance from FD Staff (primary staff tracking)."""
    staff_map: Dict[str, int] = {}

    for complaint in complaints:
        staff = complaint.fd_staff or "Unassigned"
        staff_map[staff] = staff_map.get(staff, 0) + 1

    # Sort by count descending
    sorted_staff = sorted(staff_map.items(), key=lambda x: x[1], reverse=True)
    return [{"name": name, "count": count} for name, count in sorted_staff]


def calculate_problem_areas(complaints: List[ParsedComplaint]) -> List[Dict[str, Any]]:
    """Calculate problem area breakdown."""
    problem_map: Dict[str, Dict[str, Any]] = {}

    for complaint in complaints:
        area = complaint.problem_area
        if area not in problem_map:
            problem_map[area] = {"count": 0, "rooms": set()}
        problem_map[area]["count"] += 1
        problem_map[area]["rooms"].add(complaint.room)

    total = len(complaints)

    result = []
    for area, data in sorted(problem_map.items(), key=lambda x: x[1]["count"], reverse=True):
        result.append(
            {
                "area": area,
                "count": data["count"],
                "percentage": round((data["count"] / total) * 100) if total > 0 else 0,
                "rooms": ", ".join(sorted([str(r) for r in data["rooms"]])),
            }
        )

    return result


def format_date(date: Optional[datetime]) -> str:
    """Format date for display."""
    if not date:
        return "-"
    return date.strftime("%m/%d/%Y")


def get_report_period_label(start_date: datetime, end_date: datetime) -> str:
    """Generate report period label."""
    start = format_date(start_date)
    end = format_date(end_date)
    return f"{start} to {end}"


def get_month_name(month: int) -> str:
    """Get month name."""
    months = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ]
    return months[month - 1] if 1 <= month <= 12 else ""


def calculate_membership_tiers(complaints: List[ParsedComplaint]) -> List[Dict[str, Any]]:
    """Calculate membership tier breakdown with complaint counts."""
    tier_map: Dict[str, int] = {
        "Diamond": 0,
        "Platinum": 0,
        "Gold": 0,
        "Silver": 0,
        "Club": 0,
        "Non Members": 0,
    }

    for complaint in complaints:
        tier = complaint.membership_tier or "Non Members"
        # Normalize tier name
        tier_normalized = tier.strip().title()
        
        if tier_normalized in tier_map:
            tier_map[tier_normalized] += 1
        else:
            # If it doesn't match exactly, try to categorize
            if "non" in tier_normalized.lower():
                tier_map["Non Members"] += 1
            else:
                tier_map["Non Members"] += 1

    # Build result list in order with alert for high counts
    result = []
    for tier in ["Diamond", "Platinum", "Gold", "Silver", "Club", "Non Members"]:
        count = tier_map[tier]
        needs_attention = count > 10
        result.append({
            "tier": tier,
            "count": count,
            "needsAttention": needs_attention,
        })

    return result


def get_staff_color(staff_name: Optional[str]) -> str:
    """Get a consistent light color for a staff member name."""
    if not staff_name:
        return "#e8e8e8"  # Light gray for no staff
    
    # Color palette - light colors that don't overpower
    colors = [
        "#cce5ff",  # Light blue
        "#e6ccff",  # Light purple
        "#ffe6cc",  # Light orange
        "#ccffcc",  # Light green
        "#ffcccc",  # Light red
        "#ffffcc",  # Light yellow
        "#ccffff",  # Light cyan
        "#ffccff",  # Light magenta
        "#e6f2ff",  # Very light blue
        "#f0e6ff",  # Very light purple
    ]
    
    # Generate a consistent index based on staff name
    hash_value = sum(ord(c) for c in staff_name.lower())
    color_index = hash_value % len(colors)
    return colors[color_index]
