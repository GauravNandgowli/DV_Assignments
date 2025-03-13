


Sub CreateMeetingRequest()
    Dim outlookApp As Object
    Dim meetingItem As Object
    Dim recipient As Object
    
    ' Create a new instance of Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    
    ' Create a new meeting request
    Set meetingItem = outlookApp.CreateItem(1) ' 1 corresponds to olAppointmentItem
    
    ' Set the properties of the meeting
    With meetingItem
        .Subject = "Team Meeting"
        .Location = "Conference Room A"
        .Start = #10/25/2023 10:00:00 AM# ' Set the start date and time
        .End = #10/25/2023 11:00:00 AM# ' Set the end date and time
        .Body = "Please attend the team meeting to discuss project updates."
        
        ' Add a recipient
        Set recipient = .Recipients.Add("recipient@example.com")
        recipient.Type = 1 ' 1 corresponds to olRequired
        
        ' Send the meeting request
        .Send
    End With



Sub CreateMeeting()
    Dim outlookApp As Object
    Dim meetingItem As Object
    
    ' Create Outlook Application
    On Error Resume Next
    Set outlookApp = CreateObject("Outlook.Application")
    If outlookApp Is Nothing Then
        MsgBox "Outlook is not installed or not running.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    MsgBox "Outlook application created."
    
    ' Create a Meeting Item
    On Error Resume Next
    Set meetingItem = outlookApp.CreateItem(1) ' 1 is the constant value for olMeetingRequest
    If meetingItem Is Nothing Then
        MsgBox "Failed to create meeting item.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    MsgBox "Meeting item created."

    ' Set Meeting Details
    meetingItem.Subject = "Team Meeting"
    meetingItem.Location = "Conference Room"
    meetingItem.Start = Now + TimeValue("10:00 AM")
    meetingItem.Duration = 60

    ' Add Attendees
    meetingItem.Recipients.Add "john.doe@company.com"
    meetingItem.Recipients.Add "jane.smith@company.com"

    ' Display the meeting invite
    meetingItem.Display
    MsgBox "Meeting invite displayed."
End Sub





    
    ' Clean up
    Set recipient = Nothing
    Set meetingItem = Nothing
    Set outlookApp = Nothing
End Sub 

End sub
