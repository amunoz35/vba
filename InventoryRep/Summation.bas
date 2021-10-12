Attribute VB_Name = "Module1"
'https://www.extendoffice.com/documents/excel/3963-excel-sum-until-blank.html
Sub InsertTotals()
    Dim xRg As Range
    Dim i, j, StartRow, StartCol As Integer
    Dim xTxt As String
    On Error Resume Next
    Worksheets("Greige Goods Inventario").Activate
    Set xRg = Range("H2:H900")
    If xRg Is Nothing Then Exit Sub
    StartRow = xRg.row
    StartCol = xRg.Column
    For i = StartCol To xRg.Columns.Count + StartCol - 1
        For j = xRg.row To xRg.Rows.Count + StartRow - 1
            If Cells(j, i) = "" Then
                Cells(j, i).Formula = "=SUM(" & Cells(StartRow, i).Address & ":" & Cells(j - 1, i).Address & ")"
                StartRow = j + 1
            End If
        Next
        StartRow = xRg.row
    Next
    
    Set xRg = Range("F2:G900")
    If xRg Is Nothing Then Exit Sub
    StartRow = xRg.row
    StartCol = xRg.Column
    For i = StartCol To xRg.Columns.Count + StartCol - 1
        For j = xRg.row To xRg.Rows.Count + StartRow - 1
            If Cells(j, i) = "" Then
                Cells(j, i).Formula = "=SUM(" & Cells(StartRow, i).Address & ":" & Cells(j - 1, i).Address & ")"
                StartRow = j + 1
            End If
        Next
        StartRow = xRg.row
    Next
    
    Set xRg = Range("K2:K900")
     If xRg Is Nothing Then Exit Sub
    StartRow = xRg.row
    StartCol = xRg.Column
    For i = StartCol To xRg.Columns.Count + StartCol - 1
        For j = xRg.row To xRg.Rows.Count + StartRow - 1
            If Cells(j, i) = "" Then
                Cells(j, i).Formula = "=SUM(" & Cells(StartRow, i).Address & ":" & Cells(j - 1, i).Address & ")"
                StartRow = j + 1
            End If
        Next
        StartRow = xRg.row
    Next
    
    Set xRg = Range("N2:N900")
    If xRg Is Nothing Then Exit Sub
    StartRow = xRg.row
    StartCol = xRg.Column
    For i = StartCol To xRg.Columns.Count + StartCol - 1
        For j = xRg.row To xRg.Rows.Count + StartRow - 1
            If Cells(j, i) = "" Then
                Cells(j, i).Formula = "=SUM(" & Cells(StartRow, i).Address & ":" & Cells(j - 1, i).Address & ")"
                StartRow = j + 1
            End If
        Next
        StartRow = xRg.row
    Next
    
    Set xRg = Range("Q2:Q900")
    If xRg Is Nothing Then Exit Sub
    StartRow = xRg.row
    StartCol = xRg.Column
    For i = StartCol To xRg.Columns.Count + StartCol - 1
        For j = xRg.row To xRg.Rows.Count + StartRow - 1
            If Cells(j, i) = "" Then
                Cells(j, i).Formula = "=SUM(" & Cells(StartRow, i).Address & ":" & Cells(j - 1, i).Address & ")"
                StartRow = j + 1
            End If
        Next
        StartRow = xRg.row
    Next
End Sub

Sub HideSheet()
    ActiveSheet.Visible = xlSheetVeryHidden
End Sub

