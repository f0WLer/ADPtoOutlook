"""
Excel to Outlook Calendar Importer

This module imports employee time-off requests from Excel into Outlook calendar events.
Supports filtering by date range, clearing existing calendars, and verbose event titles.

Author: Company IT
Date: December 2025
"""

import pandas as pd
import win32com.client
from datetime import datetime, timedelta, time
from typing import Optional, Tuple, List
import win32timezone
import argparse
import sys


# ============================================================================
# CONSTANTS
# ============================================================================

OUTLOOK_CALENDAR_FOLDER = 9  # olFolderCalendar
OUTLOOK_APPOINTMENT_ITEM = 1  # olAppointmentItem
OUTLOOK_FREE_STATUS = 0  # olFree
STANDARD_WORK_HOURS = 8
UTC_OFFSET_HOURS = 6

DATE_FORMATS = ['%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y', '%m-%d-%Y']
TIME_FORMATS = ['%I:%M %p', '%H:%M:%S', '%H:%M', '%I:%M:%S %p']

EXCEL_COLUMNS = {
    'STATUS': 'REQUEST STATUS',
    'NAME': 'NAME',
    'DATE': 'TIME OFF REQUEST DATE',
    'START_TIME': 'START TIME',
    'DURATION': 'DURATION',
    'DAYS_HOURS': 'DAYS/HOURS',
    'REASON': 'REASON CODE',
    'POLICY': 'POLICY NAME'
}

STATUS_APPROVED = 'Approved'
HOURS_INDICATOR = 'HOURS'
EVENT_CATEGORY = "Time Off"


# ============================================================================
# DATA MODELS
# ============================================================================

class TimeOffRequest:
    """Represents a single time-off request from the Excel file."""
    
    def __init__(self, row: pd.Series, row_index: int):
        """Initialize a TimeOffRequest from a pandas Series row."""
        self.row_index = row_index
        self.name = row.get(EXCEL_COLUMNS['NAME'], 'Unknown')
        self.status = str(row.get(EXCEL_COLUMNS['STATUS'], '')).strip()
        self.reason = row.get(EXCEL_COLUMNS['REASON'], 'Time Off')
        self.policy = row.get(EXCEL_COLUMNS['POLICY'], '')
        self.duration = row.get(EXCEL_COLUMNS['DURATION'], 1)
        self.days_hours_type = str(row.get(EXCEL_COLUMNS['DAYS_HOURS'], '')).strip().upper()
        
        # Parse dates and times
        self.start_date = parse_date(row.get(EXCEL_COLUMNS['DATE'], ''))
        self.start_time = parse_time(row.get(EXCEL_COLUMNS['START_TIME'], ''))
        
        # Validate and normalize duration
        if pd.isna(self.duration) or self.duration <= 0:
            self.duration = 1
    
    def is_approved(self) -> bool:
        """Check if the request is approved."""
        return self.status == STATUS_APPROVED
    
    def is_valid(self) -> bool:
        """Check if the request has valid date information."""
        return self.start_date is not None
    
    def is_partial_day(self) -> bool:
        """Determine if this is a partial day (hours) or full day request."""
        # If duration is in hours and less than a full work day
        if self.days_hours_type == HOURS_INDICATOR:
            return self.duration < 24  # Less than 24 hours = partial day
        # If duration is in days, only partial if less than 1 day
        return self.duration < 1
    
    def get_num_days(self) -> int:
        """Calculate the number of days for this request."""
        # If duration is in hours (type is HOURS), convert to days
        if self.days_hours_type == HOURS_INDICATOR and self.duration >= 1:
            # Duration is in hours, convert to days
            return max(1, int(self.duration / 24))
        # Otherwise duration is already in days
        return int(self.duration) if self.duration >= 1 else 1
    
    def is_in_date_range(self, date_range: Optional[Tuple[datetime, datetime]]) -> bool:
        """Check if the request falls within the specified date range."""
        if not date_range or self.start_date is None:
            return True
        
        range_start, range_end = date_range
        return range_start <= self.start_date <= range_end


class EventConfig:
    """Configuration for creating calendar events."""
    
    def __init__(self, verbose_titles: bool = False, include_descriptions: bool = False):
        self.verbose_titles = verbose_titles
        self.include_descriptions = include_descriptions


# ============================================================================
# PARSING UTILITIES
# ============================================================================

def parse_date(date_str) -> Optional[datetime]:
    """
    Parse date string in various formats.
    
    Args:
        date_str: Date string to parse
        
    Returns:
        datetime object or None if parsing fails
    """
    if pd.isna(date_str):
        return None
    
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(str(date_str), fmt)
        except (ValueError, TypeError):
            continue
    
    return None


def parse_time(time_str) -> Optional[time]:
    """
    Parse time string in various formats.
    
    Args:
        time_str: Time string to parse
        
    Returns:
        time object or None if parsing fails
    """
    if pd.isna(time_str):
        return None
    
    time_str = str(time_str).strip()
    
    for fmt in TIME_FORMATS:
        try:
            return datetime.strptime(time_str, fmt).time()
        except (ValueError, TypeError):
            continue
    
    return None


def parse_date_range_args(start_str: str, end_str: str) -> Tuple[datetime, datetime]:
    """
    Parse and validate date range arguments.
    
    Args:
        start_str: Start date in MM-DD-YYYY format
        end_str: End date in MM-DD-YYYY format
        
    Returns:
        Tuple of (start_date, end_date)
        
    Raises:
        ValueError: If dates are invalid or end is before start
    """
    try:
        range_start = datetime.strptime(start_str, '%m-%d-%Y')
        range_end = datetime.strptime(end_str, '%m-%d-%Y')
        
        if range_start > range_end:
            raise ValueError("Start date must be before or equal to end date")
        
        return range_start, range_end
    except ValueError as e:
        raise ValueError(f"Invalid date range: {e}")


# ============================================================================
# EXCEL DATA PROCESSING
# ============================================================================

def load_time_off_requests(excel_file: str) -> List[TimeOffRequest]:
    """
    Load time-off requests from Excel file.
    
    Args:
        excel_file: Path to the Excel file
        
    Returns:
        List of TimeOffRequest objects
    """
    print(f"Reading {excel_file}...")
    df = pd.read_excel(excel_file)
    
    requests = []
    for idx, (index, row) in enumerate(df.iterrows()):
        request = TimeOffRequest(row, idx)
        requests.append(request)
    
    return requests


def filter_requests(
    requests: List[TimeOffRequest], 
    date_range: Optional[Tuple[datetime, datetime]] = None
) -> List[TimeOffRequest]:
    """
    Filter requests to only approved and valid entries within date range.
    
    Args:
        requests: List of TimeOffRequest objects
        date_range: Optional tuple of (start_date, end_date) to filter by
        
    Returns:
        Filtered list of TimeOffRequest objects
    """
    filtered = []
    
    for request in requests:
        if not request.is_approved():
            continue
        
        if not request.is_valid():
            print(f"Skipping row {request.row_index}: Invalid date format")
            continue
        
        if not request.is_in_date_range(date_range):
            continue
        
        filtered.append(request)
    
    return filtered


def calculate_date_range_from_requests(requests: List[TimeOffRequest]) -> Tuple[Optional[datetime], Optional[datetime]]:
    """
    Calculate the earliest and latest dates from a list of requests.
    
    Args:
        requests: List of TimeOffRequest objects
        
    Returns:
        Tuple of (earliest_date, latest_date)
    """
    earliest = None
    latest = None
    
    for request in requests:
        if not request.is_approved() or not request.is_valid():
            continue
        
        if request.start_date is not None:
            if earliest is None or request.start_date < earliest:
                earliest = request.start_date
            
            if latest is None or request.start_date > latest:
                latest = request.start_date
    
    return earliest, latest


# ============================================================================
# OUTLOOK CONNECTION AND CALENDAR MANAGEMENT
# ============================================================================

class OutlookConnection:
    """Manages connection to Outlook application."""
    
    def __init__(self):
        """Initialize Outlook connection."""
        print("Connecting to Outlook...")
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self._display_account_info()
    
    def _display_account_info(self):
        """Display Outlook account information."""
        print(f"\nOutlook Account Information:")
        
        try:
            print(f"  Current User: {self.namespace.CurrentUser.Name}")
        except Exception as e:
            print(f"  Current User: Could not retrieve ({e})")
        
        try:
            accounts = self.outlook.Session.Accounts
            print(f"  Number of accounts: {accounts.Count}")
            for i in range(1, accounts.Count + 1):
                account = accounts.Item(i)
                print(f"    Account {i}: {account.DisplayName} ({account.SmtpAddress})")
        except Exception as e:
            print(f"  Could not retrieve account details: {e}")
    
    def get_calendar_folder(self):
        """Get the default Outlook calendar folder."""
        calendar_folder = self.namespace.GetDefaultFolder(OUTLOOK_CALENDAR_FOLDER)
        print(f"\nDefault Calendar Folder: {calendar_folder.Name}")
        
        try:
            print(f"Calendar Path: {calendar_folder.FolderPath}")
            print(f"Store Name: {calendar_folder.Store.DisplayName}")
        except Exception as e:
            print(f"Calendar Path: Could not retrieve ({e})")
        
        return calendar_folder
    
    def get_or_create_subfolder(self, parent_folder, folder_name: str):
        """
        Get an existing calendar subfolder or create a new one.
        
        Args:
            parent_folder: Parent folder object
            folder_name: Name of the subfolder
            
        Returns:
            Calendar folder object
        """
        try:
            subfolder = parent_folder.Folders(folder_name)
            print(f"Using existing calendar: {folder_name}")
        except:
            subfolder = parent_folder.Folders.Add(folder_name)
            print(f"Created new calendar: {folder_name}")
        
        return subfolder


def generate_calendar_name(
    base_name: str, 
    earliest_date: Optional[datetime], 
    latest_date: Optional[datetime]
) -> str:
    """
    Generate calendar name with date range.
    
    Args:
        base_name: Base name for the calendar
        earliest_date: Earliest date in the dataset
        latest_date: Latest date in the dataset
        
    Returns:
        Formatted calendar name
    """
    if earliest_date and latest_date:
        return f"{base_name}: {earliest_date.strftime('%b %Y')} - {latest_date.strftime('%b %Y')}"
    return base_name


def clear_calendar(calendar_name: str) -> None:
    """
    Delete all events from the specified calendar.
    
    Args:
        calendar_name: Name of the calendar folder to clear
    """
    print(f"Connecting to Outlook for clearing...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    calendar_folder = namespace.GetDefaultFolder(OUTLOOK_CALENDAR_FOLDER)
    
    try:
        target_calendar = calendar_folder.Folders(calendar_name)
        print(f"Found calendar: {calendar_name}")
    except:
        print(f"Calendar '{calendar_name}' not found, skipping clear.")
        return
    
    items = target_calendar.Items
    item_count = items.Count
    
    if item_count == 0:
        print("Calendar is already empty.")
        return
    
    print(f"Clearing {item_count} events from '{calendar_name}'...")
    
    # Collect all items first to avoid index shifting issues
    items_to_delete = []
    for i in range(1, item_count + 1):
        try:
            items_to_delete.append(items.Item(i))
        except:
            pass
    
    # Now delete them
    deleted = 0
    for item in items_to_delete:
        try:
            item.Delete()
            deleted += 1
            if deleted % 50 == 0:
                print(f"  Deleted {deleted}...")
        except Exception as e:
            # Item may have already been deleted or moved
            pass
    
    print(f"✓ Cleared {deleted} events from calendar\n")


# ============================================================================
# EVENT CREATION
# ============================================================================

def check_duplicate_event(calendar_folder, subject: str, start: datetime, end: datetime, is_all_day: bool = False):
    """
    Check if an event with the same subject on the same day already exists.
    
    Args:
        calendar_folder: Calendar folder to search
        subject: Event subject to match (employee name)
        start: Event start datetime (LOCAL time, not UTC)
        end: Event end datetime (LOCAL time, not UTC)
        is_all_day: Whether this is an all-day event
        
    Returns:
        Tuple of (existing_appointment, needs_update)
        - existing_appointment: The existing appointment object if found, None otherwise
        - needs_update: True if times are different and need updating, False otherwise
    """
    try:
        items = calendar_folder.Items
        items.IncludeRecurrences = False
        items.Sort("[Start]")
        
        # Filter items around the target date for efficiency
        filter_str = f"[Start] >= '{(start - timedelta(days=1)).strftime('%m/%d/%Y')}' AND [Start] <= '{(end + timedelta(days=1)).strftime('%m/%d/%Y')}'"
        filtered_items = items.Restrict(filter_str)
        
        for item in filtered_items:
            # Match subject (employee name) and date
            if item.Subject == subject:
                # Convert COM datetime to Python datetime for proper comparison
                # item.Start is in local timezone
                item_start_dt = datetime(item.Start.year, item.Start.month, item.Start.day,
                                        item.Start.hour, item.Start.minute, item.Start.second)
                item_end_dt = datetime(item.End.year, item.End.month, item.End.day,
                                      item.End.hour, item.End.minute, item.End.second)
                
                # Check if it's on the same day
                if item_start_dt.date() == start.date():
                    # Check if times are different (needs update)
                    if is_all_day:
                        # All-day events don't need time comparison
                        return (item, False)
                    else:
                        # For timed events, check if start/end times differ
                        time_diff_start = abs((item_start_dt - start).total_seconds())
                        time_diff_end = abs((item_end_dt - end).total_seconds())
                        if time_diff_start < 60 and time_diff_end < 60:
                            # Times are the same, no update needed
                            return (item, False)
                        else:
                            # Times are different, needs update
                            return (item, True)
    except Exception as e:
        # If there's an error checking, proceed with creation (fail-safe)
        pass
    
    return (None, False)


def create_event_title(request: TimeOffRequest, config: EventConfig) -> str:
    """
    Create event title based on configuration.
    
    Args:
        request: TimeOffRequest object
        config: EventConfig object
        
    Returns:
        Formatted event title
    """
    if config.verbose_titles:
        return f"{request.name} - {request.reason}"
    return request.name


def create_event_body(request: TimeOffRequest, day_offset: int = 0, num_days: int = 1, duration_hours: Optional[float] = None) -> str:
    """
    Create event body with detailed information.
    
    Args:
        request: TimeOffRequest object
        day_offset: Current day offset for multi-day events
        num_days: Total number of days
        duration_hours: Duration in hours for partial day events
        
    Returns:
        Formatted event body
    """
    body_parts = [
        f"Employee: {request.name}",
        f"Reason: {request.reason}",
        f"Policy: {request.policy}",
    ]
    
    if duration_hours:
        body_parts.append(f"Duration: {duration_hours:.1f} hour(s)")
    elif num_days > 1:
        body_parts.append(f"Day {day_offset + 1} of {num_days}")
    
    body_parts.append(f"Status: {request.status}")
    
    return "\n".join(body_parts)


def adjust_time_for_utc(dt: datetime, offset_hours: int = UTC_OFFSET_HOURS) -> datetime:
    """
    Adjust datetime for UTC offset.
    
    Args:
        dt: datetime object
        offset_hours: Hours to offset
        
    Returns:
        Adjusted datetime
    """
    return dt - timedelta(hours=offset_hours)


def configure_appointment_base(appointment, request: TimeOffRequest, config: EventConfig, outlook):
    """
    Configure common appointment properties.
    
    Args:
        appointment: Outlook appointment object
        request: TimeOffRequest object
        config: EventConfig object
        outlook: Outlook application object
    """
    appointment.Subject = create_event_title(request, config)
    appointment.Categories = EVENT_CATEGORY
    appointment.BusyStatus = OUTLOOK_FREE_STATUS
    appointment.StartTimeZone = outlook.TimeZones.CurrentTimeZone
    appointment.EndTimeZone = outlook.TimeZones.CurrentTimeZone


def create_all_day_event(
    calendar_folder,
    request: TimeOffRequest,
    current_date: datetime,
    day_offset: int,
    config: EventConfig,
    outlook
) -> bool:
    """
    Create an all-day calendar event.
    
    Args:
        calendar_folder: Target calendar folder
        request: TimeOffRequest object
        current_date: Date for the event
        day_offset: Day offset for multi-day events
        config: EventConfig object
        outlook: Outlook application object
        
    Returns:
        True if event was created successfully, False otherwise
    """
    try:
        # Check for duplicate
        subject = create_event_title(request, config)
        end_date = current_date + timedelta(days=1)
        existing, needs_update = check_duplicate_event(calendar_folder, subject, current_date, end_date, is_all_day=True)
        
        if existing and not needs_update:
            print(f"Skipped (duplicate): {request.name} - {current_date.strftime('%m/%d/%Y')} (All Day)")
            return False
        
        appointment = existing if existing else calendar_folder.Items.Add(OUTLOOK_APPOINTMENT_ITEM)
        appointment.Subject = create_event_title(request, config)
        appointment.Start = current_date
        appointment.End = current_date + timedelta(days=1)
        appointment.AllDayEvent = True
        appointment.Categories = EVENT_CATEGORY
        appointment.BusyStatus = OUTLOOK_FREE_STATUS
        
        if config.include_descriptions:
            appointment.Body = create_event_body(request, day_offset, request.get_num_days())
        
        appointment.Save()
        print(f"Created: {request.name} - {current_date.strftime('%m/%d/%Y')} (All Day)")
        return True
    
    except Exception as e:
        print(f"Error creating event for {request.name}: {str(e)}")
        return False


def create_partial_day_event(
    calendar_folder,
    request: TimeOffRequest,
    current_date: datetime,
    config: EventConfig,
    outlook
) -> bool:
    """
    Create a partial day (hourly) calendar event.
    
    Args:
        calendar_folder: Target calendar folder
        request: TimeOffRequest object
        current_date: Date for the event
        config: EventConfig object
        outlook: Outlook application object
        
    Returns:
        True if event was created successfully, False otherwise
    """
    if request.start_time is None:
        return False
    
    try:
        # Calculate event times for duplicate check
        start_datetime = datetime.combine(current_date.date(), request.start_time)
        
        # Determine if duration is in hours or days
        if request.days_hours_type == HOURS_INDICATOR:
            # Duration is already in hours
            duration_hours = request.duration
        else:
            # Duration is in days, convert to hours
            duration_hours = request.duration * 24
        
        end_datetime = start_datetime + timedelta(hours=duration_hours)
        
        # Check for duplicate using LOCAL times
        subject = create_event_title(request, config)
        existing, needs_update = check_duplicate_event(calendar_folder, subject, start_datetime, end_datetime, is_all_day=False)
        
        # Adjust for UTC for Outlook
        start_datetime_utc = adjust_time_for_utc(start_datetime)
        end_datetime_utc = adjust_time_for_utc(end_datetime)
        
        if existing and not needs_update:
            print(f"Skipped (duplicate): {request.name} - {start_datetime.strftime('%m/%d/%Y %I:%M %p')} ({duration_hours:.1f} hrs)")
            return False
        elif existing and needs_update:
            print(f"Updating: {request.name} - {start_datetime.strftime('%m/%d/%Y %I:%M %p')} ({duration_hours:.1f} hrs)")
            appointment = existing
        else:
            appointment = calendar_folder.Items.Add(OUTLOOK_APPOINTMENT_ITEM)
        appointment.AllDayEvent = False
        configure_appointment_base(appointment, request, config, outlook)
        appointment.Start = start_datetime_utc
        appointment.End = end_datetime_utc
        
        if config.include_descriptions:
            appointment.Body = create_event_body(request, duration_hours=duration_hours)
        
        appointment.Save()
        
        print(f"Created: {request.name} - {start_datetime.strftime('%m/%d/%Y %I:%M %p')} "
              f"to {end_datetime.strftime('%I:%M %p')} ({duration_hours:.1f} hrs)")
        return True
    
    except Exception as e:
        print(f"Error creating event for {request.name}: {str(e)}")
        return False


def create_full_day_event(
    calendar_folder,
    request: TimeOffRequest,
    current_date: datetime,
    day_offset: int,
    config: EventConfig,
    outlook
) -> bool:
    """
    Create a full work day calendar event with specific hours.
    
    Args:
        calendar_folder: Target calendar folder
        request: TimeOffRequest object
        current_date: Date for the event
        day_offset: Day offset for multi-day events
        config: EventConfig object
        outlook: Outlook application object
        
    Returns:
        True if event was created successfully, False otherwise
    """
    if request.start_time is None:
        return False
    
    try:
        # Calculate event times for duplicate check
        work_start = datetime.combine(current_date.date(), request.start_time)
        work_end = work_start + timedelta(hours=STANDARD_WORK_HOURS)
        
        # Check for duplicate using LOCAL times
        subject = create_event_title(request, config)
        existing, needs_update = check_duplicate_event(calendar_folder, subject, work_start, work_end, is_all_day=False)
        
        # Adjust for UTC for Outlook
        work_start_utc = adjust_time_for_utc(work_start)
        work_end_utc = adjust_time_for_utc(work_end)
        
        if existing and not needs_update:
            print(f"Skipped (duplicate): {request.name} - {work_start.strftime('%m/%d/%Y %I:%M %p')}")
            return False
        elif existing and needs_update:
            print(f"Updating: {request.name} - {work_start.strftime('%m/%d/%Y %I:%M %p')}")
            appointment = existing
        else:
            appointment = calendar_folder.Items.Add(OUTLOOK_APPOINTMENT_ITEM)
        appointment.AllDayEvent = False
        configure_appointment_base(appointment, request, config, outlook)
        appointment.Start = work_start_utc
        appointment.End = work_end_utc
        
        if config.include_descriptions:
            appointment.Body = create_event_body(request, day_offset, request.get_num_days())
        
        appointment.Save()
        
        print(f"Created: {request.name} - {work_start.strftime('%m/%d/%Y %I:%M %p')} "
              f"to {work_end.strftime('%I:%M %p')}")
        return True
    
    except Exception as e:
        print(f"Error creating event for {request.name}: {str(e)}")
        return False


def create_events_for_request(
    calendar_folder,
    request: TimeOffRequest,
    config: EventConfig,
    outlook
) -> int:
    """
    Create calendar event(s) for a single time-off request.
    
    Args:
        calendar_folder: Target calendar folder
        request: TimeOffRequest object
        config: EventConfig object
        outlook: Outlook application object
        
    Returns:
        Number of events successfully created
    """
    if request.start_date is None:
        return 0
    
    events_created = 0
    num_days = request.get_num_days()
    
    # All-day events (no start time specified)
    if request.start_time is None:
        for day_offset in range(num_days):
            current_date = request.start_date + timedelta(days=day_offset)
            if create_all_day_event(calendar_folder, request, current_date, day_offset, config, outlook):
                events_created += 1
        return events_created
    
    # Partial day event (single event only)
    if request.is_partial_day():
        if create_partial_day_event(calendar_folder, request, request.start_date, config, outlook):
            events_created += 1
        return events_created
    
    # Full work day events (may span multiple days)
    for day_offset in range(num_days):
        current_date = request.start_date + timedelta(days=day_offset)
        if create_full_day_event(calendar_folder, request, current_date, day_offset, config, outlook):
            events_created += 1
    
    return events_created


# ============================================================================
# MAIN WORKFLOW
# ============================================================================

def import_time_off_to_outlook(
    excel_file: str,
    calendar_base_name: str = "Employee Time Off",
    verbose_titles: bool = False,
    include_descriptions: bool = False,
    date_range: Optional[Tuple[datetime, datetime]] = None,
    clear_existing: bool = False
) -> None:
    """
    Main function to import time-off requests from Excel to Outlook.
    
    Args:
        excel_file: Path to the Excel file
        calendar_base_name: Base name for the calendar
        verbose_titles: Include reason code in event titles
        include_descriptions: Include detailed descriptions in event body
        date_range: Optional tuple of (start_date, end_date) to filter events
        clear_existing: Whether to clear existing calendar before importing
    """
    # Load and filter requests
    all_requests = load_time_off_requests(excel_file)
    approved_requests = filter_requests(all_requests, date_range)
    
    # Calculate calendar name
    earliest, latest = calculate_date_range_from_requests(approved_requests)
    calendar_name = generate_calendar_name(calendar_base_name, earliest, latest)
    print(f"Calendar will be named: {calendar_name}")
    
    # Clear calendar if requested
    if clear_existing:
        clear_calendar(calendar_name)
    
    # Connect to Outlook and get calendar
    outlook_conn = OutlookConnection()
    calendar_folder = outlook_conn.get_calendar_folder()
    target_calendar = outlook_conn.get_or_create_subfolder(calendar_folder, calendar_name)
    
    # Create event configuration
    config = EventConfig(verbose_titles=verbose_titles, include_descriptions=include_descriptions)
    
    # Create events
    events_created = 0
    events_skipped = len(all_requests) - len(approved_requests)
    duplicates_found = 0
    
    for request in approved_requests:
        created = create_events_for_request(target_calendar, request, config, outlook_conn.outlook)
        events_created += created
        # Count duplicates (events that should have been created but weren't)
        if request.is_valid():
            expected_events = request.get_num_days()
            duplicates_found += (expected_events - created)
    
    # Summary
    print(f"\n✓ Complete!")
    print(f"  Events created: {events_created}")
    print(f"  Events skipped (not approved/invalid): {events_skipped}")
    if duplicates_found > 0:
        print(f"  Duplicates skipped: {duplicates_found}")
    print(f"\nCalendar '{calendar_name}' has been created/updated in Outlook.")
    print("You can now view all employee time off in this calendar.")


def handle_clear_operation(excel_file: str, calendar_base_name: str) -> None:
    """
    Handle the calendar clearing operation before import.
    
    Args:
        excel_file: Path to the Excel file
        calendar_base_name: Base name for the calendar
    """
    print("Clear flag detected. Calendar will be cleared before importing.\n")
    
    try:
        print(f"Reading {excel_file} to determine calendar name...")
        all_requests = load_time_off_requests(excel_file)
        approved_requests = filter_requests(all_requests)
        
        earliest, latest = calculate_date_range_from_requests(approved_requests)
        calendar_name = generate_calendar_name(calendar_base_name, earliest, latest)
        
        clear_calendar(calendar_name)
    except Exception as e:
        print(f"Error during clear operation: {e}")
        print("Continuing with import...\n")


# ============================================================================
# COMMAND LINE INTERFACE
# ============================================================================

def parse_arguments():
    """Parse and validate command-line arguments."""
    parser = argparse.ArgumentParser(
        description='Import employee time-off requests from Excel to Outlook calendar',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  %(prog)s example.xlsx
  %(prog)s example.xlsx --clear
  %(prog)s example.xlsx --verbose
  %(prog)s example.xlsx --name "Staff Vacation Calendar"
  %(prog)s example.xlsx --clear --verbose
  %(prog)s example.xlsx --range 02-01-2026 02-14-2026
  %(prog)s example.xlsx --clear --verbose --range 02-01-2026 02-28-2026 --name "Q1 Time Off"
        ''')
    
    parser.add_argument(
        'excel_file',
        nargs='?',
        default='example.xlsx',
        help='Path to the Excel file (default: example.xlsx)'
    )
    
    parser.add_argument(
        '--clear',
        action='store_true',
        help='Clear existing calendar before importing'
    )
    
    parser.add_argument(
        '--verbose',
        action='store_true',
        help='Include reason code in event titles (e.g., "John Doe - PTO")'
    )
    
    parser.add_argument(
        '--range',
        nargs=2,
        metavar=('START', 'END'),
        help='Only import events within date range (format: MM-DD-YYYY MM-DD-YYYY)'
    )
    
    parser.add_argument(
        '--name',
        type=str,
        default='Employee Time Off',
        help='Base name for the calendar (default: "Employee Time Off")'
    )
    
    return parser.parse_args()


def main():
    """Main entry point for the application."""
    args = parse_arguments()
    
    # Parse date range if provided
    date_range = None
    if args.range:
        try:
            date_range = parse_date_range_args(args.range[0], args.range[1])
            print(f"Filtering events from {date_range[0].strftime('%m/%d/%Y')} "
                  f"to {date_range[1].strftime('%m/%d/%Y')}\n")
        except ValueError as e:
            print(f"Error: {e}")
            print("Please use format: MM-DD-YYYY MM-DD-YYYY")
            sys.exit(1)
    
    # Import time-off requests
    try:
        import_time_off_to_outlook(
            excel_file=args.excel_file,
            calendar_base_name=args.name,
            verbose_titles=args.verbose,
            date_range=date_range,
            clear_existing=args.clear
        )
    except FileNotFoundError:
        print(f"Error: File '{args.excel_file}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
