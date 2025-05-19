import win32com.client
from datetime import datetime, timedelta
import re

# --- Hardcoded Meeting Metadata: ID -> details ---
meeting_data = {
    "315767930059": {
        "subject": "Project Kickoff",
        "attendees": "alice@example.com; bob@example.com",
        "body": "Agenda: Kickoff discussion"
    },
    "472993002142": {
        "subject": "Design Review",
        "attendees": "carol@example.com; dave@example.com",
        "body": "Agenda: Design walkthrough"
    },
    "569103420788": {
        "subject": "Sprint Planning",
        "attendees": "erin@example.com; frank@example.com",
        "body": "Agenda: Sprint objectives"
    },
    "628390114765": {
        "subject": "Tech Sync",
        "attendees": "grace@example.com; hank@example.com",
        "body": "Agenda: Technical updates"
    },
    "748293760112": {
        "subject": "Retrospective",
        "attendees": "ivy@example.com; jack@example.com",
        "body": "Agenda: Team feedback and improvements"
    }
}

# --- Date range ---
today = datetime.now()
end_date = today + timedelta(days=3)

# --- Shared Room ---
meeting_room = "MeetingRoom@example.com"

# --- Access Outlook Calendar ---
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
calendar = namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar

items = calendar.Items
items.Sort("[Start]")
items.IncludeRecurrences = False

# --- Apply date filter ---
restriction = "[Start] >= '{}' AND [Start] <= '{}'".format(
    today.strftime("%m/%d/%Y %I:%M %p"),
    end_date.strftime("%m/%d/%Y %I:%M %p")
)
filtered_items = items.Restrict(restriction)

# --- Process Matching Items ---
updated_ids = set()

for item in filtered_items:
    try:
        if item.Subject.strip() == "Master":
            body = item.Body
            match = re.search(r"Meeting ID[:\s]+([\d\s]{11,})", body)
            if match:
                raw_id = match.group(1)
                meeting_id = re.sub(r"\D", "", raw_id)  # Remove spaces and non-digits

                if meeting_id in meeting_data and meeting_id not in updated_ids:
                    data = meeting_data[meeting_id]

                    print(f"Updating item with Meeting ID {meeting_id}...")

                    item.Subject = data["subject"]
                    item.MeetingStatus = 1  # olMeeting
                    item.RequiredAttendees = f"{data['attendees']}; {meeting_room}"
                    item.Body = data["body"]
                    item.ReminderMinutesBeforeStart = 15
                    item.BusyStatus = 2  # olBusy

                    item.Save()
                    item.Send()

                    updated_ids.add(meeting_id)
    except Exception as e:
        print(f"Error processing item: {e}")

print(f"Updated {len(updated_ids)} Master meetings.")