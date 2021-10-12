Attribute VB_Name = "Module2"
Sub ExportWorkbookQ()
    Dim FileExStr As String
    Dim FileFormatNum As Long
    Dim xWs As Worksheet
    Dim xWb As Workbook
    Dim xNWb As Workbook
    Dim FolderName As String
    Application.ScreenUpdating = False
    Set xWb = Application.ThisWorkbook
    
    DateString = Format(Now, "mm-dd-yyyy")
    FolderName = xWb.Path & "\" & "CSHEETS Sent to QCLab" & " " & DateString
    
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

Sub AutoFilterQC()
    Dim WS As Worksheet
    Dim rng As Range
    Dim LastRow As Long
    
    Set WS = ActiveWorkbook.Sheets("Sheet1")
    LastRow = WS.Range("I" & WS.Rows.Count).End(xlUp).row
    Set rng = WS.Range("I1:I" & LastRow)
    With rng
        .AutoFilter Field:=1, Criteria1:="<>*QCLAB*"
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End With
    'Turn off all the filters
    WS.AutoFilterMode = False
End Sub

Sub AutoFilterQCLABV2()
    Dim WS As Worksheet
    Dim rng As Range
    Dim LastRow As Long
    Dim i As Long
    Dim pos As Integer
    
    Set WS = ActiveWorkbook.Sheets("Sheet1")
    LastRow = WS.Range("I" & WS.Rows.Count).End(xlUp).row
    Set rng = WS.Range("I1:I" & LastRow)
    
    For i = LastRow To 1 Step -1
        pos = InStr(LCase(rng.Item(i).Value), LCase("QCLAB"))
        If pos > 0 And Rows(i) = "" Then
            rng.Item(i).EntireRow.Delete
        End If
    Next
End Sub

Sub RenameQ()
    Dim File As Worksheet
    Worksheets("Sheet1").Activate
    Sheets(2).Name = "QCLAB"
End Sub

Sub ChangeRangeQC()
    A = ActiveSheet.UsedRange.Rows.Count
End Sub

Sub vba_check_empty_cells()
    Dim i As Long
    Dim c As Long
    Dim myRange As Range
    Dim myCell As Range
    
    Set myRange = Range("A1:A")
    
    For Each myCell In myRange
        c = c + 1
        If IsEmpty(myCell) Then
            i = i + 1
        End If
    Next myCell
    
    MsgBox _
    "There are total " & i & " empty cell(s)out of " & c & "."
End Sub

Sub CheckRows()
    Dim cl As Range
    Dim Ws1 As Worksheet
    Dim Ws2 As Worksheet
    Dim Vlu As String
    Dim Lc As Long
    
    Set Ws1 = Sheets("Sheet1")
    Set Ws2 = Sheets("Sheet2")
    Lc = Ws2.Cells(1, Columns.Count).End(xlToLeft).Column
    With CreateObject("scripting.dictionary")
        For Each cl In Ws2.Range("A2", Ws2.Range("A" & Rows.Count).End(xlUp))
            Vlu = Join(Application.Index(cl.Resize(, Lc).Value, 1, 0), "|")
            .Item(Vlu) = Empty
        Next cl
        For Each cl In Ws1.Range("A2", Ws1.Range("A" & Rows.Count).End(xlUp))
            Vlu = Join(Application.Index(cl.Resize(, Lc).Value, 1, 0), "|")
            If .exists(Vlu) Then cl.Resize(, Lc).Interior.Color = vbRed
        Next cl
    End With
End Sub

Sub ExitFor_Loop()
    Dim i As Integer
    For i = 1 To 1000
        Range("A" & i).Select
        MsgBox "Error Found"
        Exit For
        End If
    Next i
End Sub

Sub ForEach_CountTo10_Even()
    Dim n As Integer
    For n = 2 To 10 Step 2
        MsgBox n
    Next n
    MsgBox "List Off"
End Sub

Sub ForEach_DeleteRows_BlankCells()
    Dim n As Integer
    For n = 10 To 1 Step -1
        If Range("a" & n).Value = "" Then
            Range("a" & n).EntireRow.Delete
        End If
    Next n
End Sub
Sub Nested_ForEach_MultiplicationTable()

    Dim row  As Integer, col As Integer
    
    For row = 1 To 9
        For col = 1 To 9
            Cells(row + 1, col + 1).Value = row * cik
        Next col
    Next row
End Sub
Sub ChangeRangeQ()
    A = ActiveSheet.UsedRange.Rows.Count
End Sub
