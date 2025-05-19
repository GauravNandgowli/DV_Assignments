import win32com.client
from datetime import datetime, timedelta

# --- Step 1: Set the meeting date manually ---
meeting_date = datetime(2025, 5, 20)  # Replace with your desired date

# --- Step 2: Set date range ---
today = datetime.today()
end_date = meeting_date + timedelta(days=3)

# --- Step 3: Access Outlook Calendar ---
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
calendar = outlook.GetDefaultFolder(9)  # 9 = olFolderCalendar

items = calendar.Items
items.Sort("[Start]")
items.IncludeRecurrences = False

# --- Step 4: Filter items ---
restriction = "[Start] >= '{}' AND [Start] <= '{}'".format(
    today.strftime("%m/%d/%Y"),
    end_date.strftime("%m/%d/%Y")
)

restricted_items = items.Restrict(restriction)

# --- Step 5: Print matching calendar items ---
print(f"Calendar items from {today.date()} to {end_date.date()}:\n")

if restricted_items.Count == 0:
    print("No meetings found in this date range.")
else:
    for item in restricted_items:
        try:
            print(f"{item.Start.strftime('%Y-%m-%d %H:%M')} - {item.Subject}")
        except Exception as e:
            print("Error reading item:", e)