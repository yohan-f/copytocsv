Public Sub CopyToCSV()
    Dim copy_sheet As String
    Dim csv_file_name As String
    
    On Error GoTo ErrorHandler
    
    copy_sheet = InputBox("Please enter the sheet to be converted to CSV", "Sheet to be copied")
    
    If copy_sheet = vbNullString Then
        MsgBox ("No value entered. Exiting")
        Exit Sub
    End If
    
    csv_file_name = InputBox("Please enter the CSV file name", "CSV File Name")
    
    If csv_file_name = vbNullString Then
        MsgBox ("No value entered. Exiting")
        Exit Sub
    End If
    
    Sheets(copy_sheet).Copy
    
    With ActiveWorkbook
        .SaveAs Filename:=csv_file_name, FileFormat:=xlCSV, CreateBackup:=False
        .Close False
        Exit Sub
    End With
    
'    With ActiveWorkbook
'        .Close False
'        Exit Sub
'    End With
    
ErrorHandler:
    MsgBox ("Error encountered. Exiting.")
    Exit Sub

End Sub
