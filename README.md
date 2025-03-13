


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
    
    ' Clean up
    Set recipient = Nothing
    Set meetingItem = Nothing
    Set outlookApp = Nothing
End Sub 

End sub
