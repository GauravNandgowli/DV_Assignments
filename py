import pandas as pd
import win32com.client
import datetime
import re
from datetime import timedelta
import time

def clean_id(meeting_body):
    """Extract and clean Meeting ID from meeting body."""
    pattern = r"Meeting ID:\s*([\d\s]+)"
    regex = re.compile(pattern, re.IGNORECASE)
    match = regex.search(meeting_body)
    if match:
        meeting_id = match.group(1).strip().replace('\r\n', '').replace('\n', '')
        return meeting_id
    return ""

def update_meetings(meeting, subject, full_start_time, full_end_time, required_attendees, optional_attendees):
    """Update an existing Outlook meeting."""
    try:
        meeting.Subject = subject
        meeting.Start = full_start_time
        meeting.End = full_end_time
        meeting.MeetingStatus = 1  # olMeeting
        if required_attendees:
            meeting.RequiredAttendees = required_attendees
        if optional_attendees:
            meeting.OptionalAttendees = optional_attendees
        meeting.Save()
        print(f"Updated meeting: {subject}")
        return "done"
    except Exception as e:
        print(f"Error updating meeting: {e}")
        return "error"

def process_meeting_items(meeting_id, meeting_date, subject, start_time, end_time, required_attendees, optional_attendees):
    """Process Outlook calendar items to find and update a meeting."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(6)  # olFolderCalendar
        items = calendar.Items

        # Restrict items to a date range
        today = datetime.date.today()
        end_date = meeting_date + timedelta(days=4)
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        found = False
        for item in items:
            try:
                item_start = item.Start
                if isinstance(item_start, str):
                    item_start = datetime.datetime.strptime(item_start, "%m/%d/%Y %I:%M %p")
                if today <= item_start.date() <= end_date:
                    if item.Body and len(item.Body) > 0:
                        cleaned_id = clean_id(item.Body)
                        if cleaned_id == meeting_id and item.Subject == "Master":
                            result = update_meetings(item, subject, start_time, end_time, required_attendees, optional_attendees)
                            found = True
                            break
            except Exception as e:
                print(f"Error processing item: {e}")
                continue

        print(f"Meeting ID {meeting_id} found: {found}")
        return found
    except Exception as e:
        print(f"Error in process_meeting_items: {e}")
        return False

def main():
    # Path to the Excel workbook
    workbook_path = "C:/path/to/your/workbook.xlsx"  # Update with your actual path
    sheet_name = "Create Meetings"

    try:
        # Read Excel data
        df = pd.read_excel(workbook_path, sheet_name=sheet_name, header=None, skiprows=1)

        # Process rows (starting from row 2 in Excel, index 0 in DataFrame)
        for index, row in df.iterrows():
            # Skip empty rows
            if row.isna().all():
                continue

            try:
                # Extract data
                meeting_date = pd.to_datetime(row[0]).date() if pd.notna(row[0]) else None
                start_time_str = row[1] if pd.notna(row[1]) else None
                end_time_str = row[2] if pd.notna(row[2]) else None
                subject = str(row[3]) if pd.notna(row[3]) else "Meeting"
                required_attendees = str(row[4]) if pd.notna(row[4]) else ""
                optional_attendees = str(row[5]) if pd.notna(row[5]) else ""
                meeting_id = str(row[6]) if pd.notna(row[6]) else ""

                # Validate date
                if not meeting_date:
                    print(f"Invalid date in row {index + 2}")
                    continue

                # Parse times
                try:
                    start_time = pd.to_datetime(str(start_time_str), format='%H:%M:%S').time()
                    end_time = pd.to_datetime(str(end_time_str), format='%H:%M:%S').time()
                except ValueError:
                    print(f"Invalid time format in row {index + 2}")
                    continue

                # Combine date and time
                full_start_time = datetime.datetime.combine(meeting_date, start_time)
                full_end_time = datetime.datetime.combine(meeting_date, end_time)

                # Clean meeting ID
                trimmed_id = meeting_id.strip().replace('\r\n', '').replace('\n', '')

                # Process the meeting
                process_meeting_items(
                    trimmed_id, meeting_date, subject,
                    full_start_time, full_end_time,
                    required_attendees, optional_attendees
                )

                # Small delay to mimic VBA's Application.Wait
                time.sleep(2)

            except Exception as e:
                print(f"Error processing row {index + 2}: {e}")
                continue

        print("Meeting scheduling process completed.")

    except Exception as e:
        print(f"Error reading workbook: {e}")

if __name__ == "__main__":
    main()