Sub CreateMeeting()
    Dim outlookApp As Object
    Dim meetingItem As Object
    
    ' Create Outlook Application
    Set outlookApp = CreateObject("Outlook.Application")
    
    ' Create a Meeting Item
    Set meetingItem = outlookApp.CreateItem(olMeetingRequest)
    
    ' Set Meeting Details
    meetingItem.Subject = "Team Meeting"
    meetingItem.Location = "Conference Room"
    meetingItem.Start = Now + TimeValue("10:00 AM")
    meetingItem.Duration = 60
    
    ' Add Attendees
    meetingItem.Recipients.Add("john.doe@company.com")
    meetingItem.Recipients.Add("jane.smith@company.com")
    
    ' Display the meeting invite
    meetingItem.Display
End Sub
