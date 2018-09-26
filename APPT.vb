Sub Appt(StartDate As Date, subj As String, Bod As String)
'Function that Populates the details of the outlook appointment
Dim olApp As Outlook.Application
Dim olAppItem As Outlook.AppointmentItem


Set olApp = GetObject("", "Outlook.Application")
Set olApp = CreateObject("Outlook.Application")
Set olAppItem = olApp.CreateItem(olAppointmentItem)

Worksheets("Sheet1").Activate

With olAppItem
    
    .Subject = subj
    .Body = Bod
    .ReminderSet = True
    .ReminderMinutesBeforeStart = 1440
    .Start = StartDate
    .AllDayEvent = True
    .Save
    
    

End With

Set olApp = Nothing
Set olAppItem = Nothing



End Sub
