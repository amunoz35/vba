Attribute VB_Name = "Module3"
Sub ResetRange()
    ActiveSheet.UsedRange
End Sub

'This sub will only highlight the blank cells in the workbook;
'will not get the empty rows, so single cells may be highlighted within the range.
Sub ColorCodeReport()
    Dim xRg As Range
    Dim row As Range
    Set rng = Range("A1:V1400")
    
    For Each row In rng
        If row.Value = "" Then
            row.Interior.Color = RGB(255, 255, 255)
        End If
    Next
End Sub

'VLOOKUP ERROR REFERENCE
'https://strugglingtoexcel.com/2014/08/11/numbers-stored-as-text-error/
Sub testColor()
    Dim NumRowsToInsert As Long
    Dim RowIncrement As Long
    Dim ws As Excel.Worksheet
    Dim lastrow As Long
    
    NumRowsToInsert = 1
    RowIncrement = 6
    Set ws = ActiveSheet
    With ws
        lastrow = .Range("A" & .Rows.Count).End(xlUp).row
        LastEvenlyDivisibleRow = Int(lastrow / RowIncrement) * RowIncrement
        If LastEvenlyDivisibleRow = 0 Then
            Exit Sub
        End If
        Application.ScreenUpdating = False
        For i = LastEvenlyDivisibleRow To 1 Step -RowIncrement
            .Range(i & ":" & i + (NumRowsToInsert - 1)).Insert xlShiftDown
            .Range("A" & i & ":H" & i + (NumRowsToInsert - 1)).Interior.TintAndShade = -0#
        Next i
    End With
    Application.ScreenUpdating = True
End Sub

Sub ExportWorkbook()
    Dim FileExStr As String
    Dim FileFormatNum As Long
    Dim xWs As Worksheet
    Dim xWb As Workbook
    Dim xNWb As Workbook
    Dim FolderName As String
    Application.ScreenUpdating = False
    Set xWb = Application.ThisWorkbook
    
    DateString = Format(Now, "yyyy-mm-dd hh-mm-ss")
    FolderName = xWb.Path & "\" & xWb.Name & " " & DateString
    
    If val(Application.Version) < 12 Then
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        Select Case xWb.FileFormat
            Case 51:
                FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If Application.ActiveWorkbook.HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
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

Sub CopySheet()
    Dim Source As Range, Dest As Range
    Dim r As Long, c As Long
    Set Source = Worksheets("Greige Goods Inventario").Range("A:V")
    Set Dest = Worksheets("Greige Goods").Range("A:V")
    'Get min rows and column
    r = WorksheetFunction.Min(Source.Rows.Count, Dest.Rows.Count)
    c = WorksheetFunction.Min(Source.Columns.Count, Dest.Columns.Count)
    'Resize to same Ranges
    Set Source = Source.Resize(r, c)
    Set Dest = Dest.Resize(r, c)
    Source.Copy
    Dest.PasteSpecial xlPasteValuesAndNumberFormats
End Sub

Sub ColorYarns()
    Dim IRow As Long
    Dim ICntr As Long
    Dim Counter As Long
    Counter = 0
    IRow = 900
    Worksheets("Greige Goods").Activate
    For ICntr = IRow To 1 Step -1
        If Trim(Cells(ICntr, 1)) = "" Then
            Counter = Counter + 1
        End If
        If Trim(Cells(ICntr, 1)) <> "" Then
            If Counter Mod 4 = 2 Then
                Rows(ICntr).Select
                Rows(ICntr).Interior.Color = RGB(170, 190, 220) 'blue
                GoTo NextIteration
            End If
            If Counter Mod 4 = 1 Then
                Rows(ICntr).Select
                Rows(ICntr).Interior.Color = RGB(230, 190, 190) 'rustic red
                GoTo NextIteration
            End If
            If Counter Mod 4 = 3 Then
                Rows(ICntr).Select
                Rows(ICntr).Interior.Color = RGB(180, 200, 150) 'green
                GoTo NextIteration
            End If
            If Counter Mod 4 = 0 Then
                Rows(ICntr).Select
                Rows(ICntr).Interior.Color = RGB(255, 255, 0)  'yellow
                GoTo NextIteration
            End If
        End If
NextIteration:
    Next
End Sub

Sub ClearTitle()
    Dim IRow As Long
    IRow = 1
    Worksheets("Greige Goods").Activate
    Rows(1).Select
    Rows(1).Interior.Color = RGB(255, 255, 255)
End Sub

Sub DeleteDuplicates()
    With Application
        .ScreenUpdating = False
        Dim LastColumn As Integer
        LastColumn = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column + 1
        With Range("A1:A" & Cells(Rows.Count, 1).End(xlUp).row)
            .AdvancedFilter Action:=xlFilterInPlace, Unique:=True
            .SpecialCells(xlCellTypeVisible).Offset(0, LastColumn - 1).Value = 1
            On Error Resume Next
            ActiveSheet.ShowAllData
            Columns(LastColumn).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
            Err.Clear
        End With
        Columns(LastColumn).Clear
        .ScreenUpdating = True
    End With
End Sub

Sub FormatTotals()
    Dim IRow As Long
    Dim ICntr As Long
    Dim Counter As Integer
    Counter = 0
    IRow = 900
    Worksheets("Greige Goods").Activate
    For ICntr = IRow To 1 Step -1
        If Cells(ICntr, 1) = "" Then
            Counter = Counter + 1
            If Counter Mod 4 = 2 Then
                Rows(ICntr).Select
                Rows(ICntr).Font.Bold = True
                Rows(ICntr).Interior.Color = RGB(170, 190, 220)
                GoTo NextIteration
            End If
            If Counter Mod 4 = 1 Then
                Rows(ICntr).Select
                Rows(ICntr).Font.Bold = True
                Rows(ICntr).Interior.Color = RGB(230, 190, 190)
                GoTo NextIteration
            End If
            If Counter Mod 4 = 3 Then
                Rows(ICntr).Select
                Rows(ICntr).Font.Bold = True
                Rows(ICntr).Interior.Color = RGB(180, 200, 150)
                GoTo NextIteration
            End If
            If Counter Mod 4 = 0 Then
                Rows(ICntr).Select
                Rows(ICntr).Font.Bold = True
                Rows(ICntr).Interior.Color = RGB(255, 255, 0)
                GoTo NextIteration
            End If
        End If
NextIteration:
    Next
End Sub

Sub CloseActiveBook()
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
End Sub
