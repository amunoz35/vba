Attribute VB_Name = "Module1"
Sub ExportWorkbookJ()
    Dim FileExStr As String
    Dim FileFormatNum As Long
    Dim xWs As Worksheet
    Dim xWb As Workbook
    Dim xNWb As Workbook
    Dim FolderName As String
    Application.ScreenUpdating = False
    Set xWb = Application.ThisWorkbook
    
    DateString = Format(Now, "yyyy-mm-dd")
    FolderName = xWb.Path & "\" & "CSHEET Sent to JOSE" & " " & DateString
    
    If Val(Application.Version) < 12 Then
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        Select Case xWb.FileFormat
            Case 51:
                FileExt.Str = ".xlsx": FileFormatNum = 51
            Case 52:
                If Application.ActiveWorkbook.HasVBProject Then
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56:
                FileExtStr = ".xls": FileFormatNum = 56
            Case Else:
                FileExtStr = ".xlsb": FileFormatNum = 50
            End Select
        End If
    
    MkDir FolderName
    For Each xWs In xWb.Worksheets
    On Error GoTo NErro
        If xWs.Visible = xlSheetVisible Then
            xWs.Select
            xWs.Copy
            xFile = FolderName & "\" & xWs.Name & FileExtStr
            Set xNWb = Application.Workbooks.Item(Application.Workbooks.Count)
            xNWb.SaveAs xFile, FileFormat:=FileFormatNum
            xNWb.Close False, xFile
            End If
NErro:
        xWb.Activate
Next
    MsgBox "You can find the files in " & FolderName
    Application.ScreenUpdating = True
End Sub

Sub AutoFilterJose()
    Dim WS As Worksheet
    Dim rng As Range
    Dim LastRow As Long
    
    Set WS = ActiveWorkbook.Sheets("Sheet1")
    LastRow = WS.Range("I" & WS.Rows.Count).End(xlUp).row
    Set rng = WS.Range("I1:I" & LastRow)
    With rng
        .AutoFilter Field:=1, Criteria1:="<>*JOSE*"
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End With
    'Turn off all the filters
    WS.AutoFilterMode = False
End Sub

Sub AutoFilterJoseV2()
    Dim WS As Worksheet
    Dim rng As Range
    Dim LastRow As Long
    Dim i As Long
    Dim pos As Integer
    
    Set WS = ActiveWorkbook.Sheets("Sheet1")
    LastRow = WS.Range("I" & WS.Rows.Count).End(xlUp).row
    Set rng = WS.Range("I1:I" & LastRow)
    
    For i = LastRow To 1 Step -1
        pos = InStr(LCase(rng.Item(i).Value), LCase("JOSE"))
        If pos > 0 And Rows(i + 1) = "" Then
            rng.Item(i).EntireRow.Delete
        End If
    Next
End Sub

Sub RenameJ()
    Dim File As Worksheet
    Worksheets("Sheet1").Activate
    Sheets(2).Name = "JOSE"
End Sub

Sub ChangeRange()
    A = ActiveSheet.UsedRange.Rows.Count
End Sub

Sub compareLines()
    Dim currentRow, nextRow As Integer
    
    currentRow = 11
    nextRow = 1
    ActiveSheet.Cells(currentRow, 1).Select
    
    While ActiveCell.Value <> ""
        If ActiveCell.Offset(nextRow, 0).Value = ActiveCell.Value Then
            If ActiveCell.Offset(0, 1).Value - ActiveCell.Offset(nextRow, 1).Value < 0 Then
                ActiveCell.EntireRow.Delete
            Else
                ActiveCell.Offset(nextRow, 0).EntireRow.Delete
            End If
        Else
            ActiveCell.Offset(nextRow, 0).Select
        End If
    Wend
End Sub

Sub DeleteDuplicates()
    Dim LastRow As Long
    Dim i As Long
    
    LastRow = Cells(Rows.Count, "A").End(xlUp).row
    
    Application.ScreenUpdating = False
    
    For i = LastRow To 11 Step -1
        If Cells(i, "B").Value <> Application.Evaluate("MAX(IF(" & Range(Cells(1, "A")).Address _
           & "=" & Cells(i, "A").Address & "," & Range(Cells(1, "B"), Cells(LastRow, "B")).Address & "))") Then
                Rows(i).Delete
        End If
    Next i
    
    Application.ScreenUpdating = True
        
End Sub

Sub Delete_Rows_Based_On_Value()
'Apply a filter to a range and delete visible row
'Source:https://www.excelcampus.com/vba/delete-rows-cell-values/

Dim WS As Worksheet

    'Set reference to the in the workbook
    Set WS = ThisWorkbook.Worksheets("Regular Range")
    WS.Activate ' not required but allows user to view sheet if arning message appears
    
    'Clear any existing filters
    On Error Resume Next
        WS.ShowAllData
    On Error GoTo 0
    
    '1. Apply Filter
    WS.Range("B3:G1000").AutoFilter Field:=4, Criteria1:=""
    '2. Delete Rows
    Application.DisplayAlerts = False
        WS.Range("B4:G1000").SpecialCells(xlCellTypeVisible).Delete
    Application.DisplayAlerts = True
    
    '3. C;ear Filter
    On Error Resume Next
       WS.ShowAllData
    On Error GoTo 0
End Sub

'Applying the macro to tables
Sub Delete_rows_based_on_value_table()
    'Source
    'Apply a filter to a table and delete visible orws
    
    Dim lo As ListObject
        'Set reference to the sheet and table
        Set lo = Sheet3.ListObjects(1)
        WS.Activate
        
        'Clear any existing filters
        lo.AutoFilter.ShowAllData
        
        '1. Apply Filter
        lo.Range.AutoFilter Field:=4, Criteria1:="Product 2"
        
        '2. Delete Rows
        Application.DisplayAlerts = False
            lo.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
        Application.DisplayAlerts = True
        
        '3. Clear Filter
        lo.AutoFilter.ShowAllData
End Sub

Sub Delete_Rows_Based_On_Value_Table_Message()
    'Display Yes/No message prompt before deleting rows
    'Source:
    
    Dim lo As ListObject
    Dim lRows As Long
    Dim vbAnswer As VbMsgBoxResult
    
        'Set reference to the sheet and Table
        Set lo = Sheet.ListObjects(1)
        lo.ParentActivate 'Activate sheet that Table is on
        
        'Clear any existing filters
        lo.AutoFilter.ShowAllData
        
        '1. Apply Filter
        lo.Range.AutoFilter Field:=4, Criteria1:="Product 2"
        
        'Count Rows & display message
        On Error Resume Next
            lRows = WorksheetFunction.Subtotal(103, lo.ListColumns(1).DataBodyRange.SpecialCells(xCellTypeVisible))
        On Error GoTo 0
        
        vbAnswer = MsgBox(lRows & " Rows Will be deleted. Do you want to continue?", vbYesNo, "Delete Rows Macro")
        
        If vbAnswer = vbYes Then
        
            'DeleteRows
            Application.DisplayAlerts = False
                lo.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
            Application.DisplayAlerts = True
            
            'Clear Filter
            lo.AutoFilter.ShowAllData
            
        End If
End Sub


