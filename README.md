Sub CreateMeeting()
    Dim outlookApp As Object
    Dim meetingItem As Object
    Dim requiredAttendee As Outlook.Recipient
    Dim optionalAttendee As Outlook.Recipient
    Dim resourceAttendee As Outlook.Recipient
    
    ' Create Outlook Application
    Set outlookApp = CreateObject("Outlook.Application")
    
    ' Create a Meeting Item (not an appointment)
    Set meetingItem = outlookApp.CreateItem(olMeetingRequest)
    
    ' Set Meeting Details
    With meetingItem
        ' Meeting Subject
        .Subject = "Strategy Meeting"
        
        ' Meeting Location
        .Location = "Conference Room B"
        
        ' Set Start Time and Duration
        .Start = DateValue("9/24/2024") + TimeValue("1:30 PM")
        .Duration = 90 ' Minutes
        
        ' Add Required Attendees
        Set requiredAttendee = .Recipients.Add("nate.sun@company.com")
        requiredAttendee.Type = olRequired
        
        ' Add Optional Attendees
        Set optionalAttendee = .Recipients.Add("kevin.kennedy@company.com")
        optionalAttendee.Type = olOptional
        
        ' Add Resource (Room)
        Set resourceAttendee = .Recipients.Add("conference.room.b@company.com")
        resourceAttendee.Type = olResource
        
        ' Meeting Body/Description
        .Body = "Detailed meeting agenda or additional information goes here."
        
        ' Optional: Set Reminder
        .ReminderSet = True
        .ReminderMinutesBeforeStart = 15
    End With
    
    ' Display or Send the meeting
    meetingItem.Display ' Use .Send() to send directly
End Sub

' Alternative with more flexibility
Sub CreateMeetingWithParameters(meetingSubject As String, startDateTime As Date, durationMinutes As Integer, requiredAttendees As String, optionalAttendees As String)
    Dim outlookApp As Object
    Dim meetingItem As Object
    Dim attendeeList As Variant
    Dim i As Integer
    
    ' Create Outlook Application
    Set outlookApp = CreateObject("Outlook.Application")
    
    ' Create a Meeting Item
    Set meetingItem = outlookApp.CreateItem(olMeetingRequest)
    
    ' Set Meeting Details
    With meetingItem
        .Subject = meetingSubject
        .Start = startDateTime
        .Duration = durationMinutes
        
        ' Add Required Attendees
        If Len(Trim(requiredAttendees)) > 0 Then
            attendeeList = Split(requiredAttendees, ";")
            For i = LBound(attendeeList) To UBound(attendeeList)
                Dim reqAttendee As Outlook.Recipient
                Set reqAttendee = .Recipients.Add(Trim(attendeeList(i)))
                reqAttendee.Type = olRequired
            Next i
        End If
        
        ' Add Optional Attendees
        If Len(Trim(optionalAttendees)) > 0 Then
            attendeeList = Split(optionalAttendees, ";")
            For i = LBound(attendeeList) To UBound(attendeeList)
                Dim optAttendee As Outlook.Recipient
                Set optAttendee = .Recipients.Add(Trim(attendeeList(i)))
                optAttendee.Type = olOptional
            Next i
        End If
    End With
    
    ' Display the meeting
    meetingItem.Display
End Sub

' Example of using the parameterized method
Sub ExampleMeetingCreation()
    CreateMeetingWithParameters _
        "Quarterly Review Meeting", _
        DateValue("10/15/2024") + TimeValue("10:00 AM"), _
        60, _
        "john.doe@company.com;jane.smith@company.com", _
        "mike.wilson@company.com"
End Sub
