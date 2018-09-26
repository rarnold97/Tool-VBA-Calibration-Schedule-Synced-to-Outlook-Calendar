Private Sub Workbook_AfterSave(ByVal Success As Boolean)
Dim DateChck As Date
    Dim Today As Date
    Dim i As Long
    Dim lastrow As Long
    Dim Body As String, Subject As String
    Dim FirstDatePos As Integer
    
    Dim pos1 As String
    
Worksheets("Sheet1").Activate

FirstDatePos = FindEmpty() + 1

pos1 = Cells(1, FirstDatePos).Address

'Checks whether the calibration date is in the future and then populates an outlook appointment
With Sheets("Sheet1")
    Today = Date
    lastrow = .Range(pos1, Range(pos1).End(xlDown)).Rows.Count
    
    For i = FirstDatePos To lastrow
    
        DateChck = .Cells(i, 1).Value
        
              
        If Today <= DateChck Then
            
            Subject = "Recalibration of: " & .Cells(i, 2).Value & " " & .Cells(i, 3).Value
            Body = "Recalibrate the following tool: " & .Cells(i, 2).Value & " " & .Cells(i, 3).Value
          
            Call Appt(DateChck, Subject, Body)
            
        
        End If
        
    
    
    Next
    
End With
    
End Sub
'Function that Populates the dates upon opening the spreadsheet
Private Sub Workbook_Open()
 Dim DateChck As Date
 Dim Today As Date
 Dim i As Long
 Dim lastrow As Long
 Dim FirstDatePos As Integer

 Dim pos1 As String

Worksheets("Sheet1").Activate

MsgBox ("Please don't enter dates in columns other than Column A for Calibration Dates")

FirstDatePos = FindEmpty() + 1

pos1 = Cells(1, FirstDatePos).Address


With Sheets("Sheet1")
    Today = Date
    lastrow = .Range(pos1, Range(pos1).End(xlDown)).Rows.Count
    
    For i = FirstDatePos To lastrow
    
        DateChck = .Cells(i, 1).Value
        
              
        If Today <= DateChck Then
            
            Call ResetCalDates(DateChck)
            
        
        End If
        
    
    
    Next
    
End With
    
End Sub


'Seearches for the end of the date column
Public Function FindEmpty()
    Dim Index As Integer

    Index = 1

Do While IsEmpty(Cells(Index, 1).Value)
    
    Index = Index + 1
    
    Loop
    
FindEmpty = Index

End Function

