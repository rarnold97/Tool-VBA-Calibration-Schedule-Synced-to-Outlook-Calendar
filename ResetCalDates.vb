Sub ResetCalDates(StartDate As Date) 
'Function that Clears the calibration dates 

'Initialize Outlook Application Objects
 Dim oApp As Outlook.Application 
 Dim oNameSpace As Outlook.Namespace
 Dim oApptItem As Outlook.AppointmentItem
 Dim oFolder As Outlook.MAPIFolder
 Dim oMeetingoApptItem As Outlook.MeetingItem
 Dim oObject As Object
 Dim sErrorMessage As String
 'Check for error upon opening outlook
 On Error Resume Next
 'Check if outlook is open and running
 Set oApp = GetObject("Outlook.Application")
 If Err <> 0 Then
    'if not running start
    Set oApp = CreateObject("Outlook.Application")
    
 End If
 
 On Error GoTo Err_Handler
 Set oNameSpace = oApp.GetNamespace("MAPI")
 Set oFolder = oNameSpace.GetDefaultFolder(olFolderCalendar)
 'Scan Through outlook folders to find appointment items 
 For Each oObject In oFolder.Items
 
    If oObject.Class = olAppointment Then
        
        Set oApptItem = oObject
        'Clear all existing outlook appointments that match dates in spreadsheet
        If oApptItem.Start = StartDate Then
        
            oApptItem.Delete
        
        
        End If
    
    
    End If
 
 
 Next oObject

'Clear Outlook Objects
Set oApp = Nothing
Set oNameSpace = Nothing
Set oApptItem = Nothing
Set oFolder = Nothing
Set oObject = Nothing

Exit Sub

Err_Handler:
    ' Error message
    sErrorMessage = Err.Number & " " & Err.Description
    MsgBox sErrorMessage

End Sub
