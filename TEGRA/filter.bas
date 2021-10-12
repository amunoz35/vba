Attribute VB_Name = "Module1"
'Written by: Alexander Munoz

'This will copy the appropriate color, ship to, price, and style

'Sub FillDown()
'Dim xRng As Range
'Dim xRows As Long, xCols As Long
'Dim xRow As Integer, xCol As Integer
'Set xRng = Selection
'Application.ScreenUpdating = False
'xCols = xRng.Columns.CountLarge
'xRows = xRng.Rows.CountLarge
'For xCol = 1 To xCols
'  For xRow = 1 To xRows - 1
'    If xRng.Cells(xRow, xCol) <> "" Then
'      xRng.Cells(xRow, xCol) = xRng.Cells(xRow, xCol).Value
'      If xRng.Cells(xRow + 1, xCol) = "" Then
'        xRng.Cells(xRow + 1, xCol) = xRng.Cells(xRow, xCol).Value
'      End If
'    End If
'  Next xRow
'Next xCol
'Application.ScreenUpdating = True
'End Sub

'Copies the sheet full of references to a clean sheet

Sub Test_Copy()
    Dim Source As Range, Dest As Range
    Dim r As Long, c As Long
    'Setup source and dest
    Set Source = Worksheets("5202ref").Range("A:Y")
    Set Dest = Worksheets("OPEN").Range("A:Y")
    'Get min rows and columns
    r = WorksheetFunction.Min(Source.Rows.Count, Dest.Rows.Count)
    c = WorksheetFunction.Min(Source.Columns.Count, Dest.Columns.Count)
    'Resize to same ranges
    Set Source = Source.Resize(r, c)
    Set Dest = Dest.Resize(r, c)
    Source.Copy
    Dest.PasteSpecial xlPasteValuesAndNumberFormats
    
End Sub

''Deletes rows with missing PO numbers
'Sub deleteBlankRows()
'    'Worksheets("OPEN").Activate
'    Dim IRow As Long
'    Dim ICntr As Long
'    Dim ws As Worksheet
'    Set ws = ActiveSheet
'    IRow = 2600
'    For ICntr = IRow To 2 Step -1
'        If Cells(ICntr, 3) = "" Then
'            Rows(ICntr).Delete
'        End If
'    Next
'End Sub

'This will delete the row if there are errors located within the row
Sub sbDelete_Rows_If_Cell_Contains_Error()
    Worksheets("OPEN").Activate
    Dim IRow  As Long
    Dim iCntr As Long
    Dim xRg As Range
    IRow = 2000
    For iCntr = IRow To 1 Step -1
        If IsError(Cells(iCntr, 11)) Then
                Rows(iCntr).Delete
        End If
    Next
End Sub

'This will remove any swap fabcodes from the list of PO's
Sub remove_Swaps()
    Worksheets("OPEN").Activate
    Dim IRow As Long
    Dim iCntr As Long
    Dim xRg As Range
    IRow = 2000
    For iCntr = IRow To 1 Step -1
        If (Cells(iCntr, 2) = "SWAP") Then
            Rows(iCntr).Delete
        End If
    Next
End Sub
'This will remove any overdye bulk from the orders
Sub remove_Overdye()
    Worksheets("OPEN").Activate
    Dim IRow As Long
    Dim iCntr As Long
    Dim xRg As Range
    IRow = 2000
    For iCntr = IRow To 1 Step -1
        If (Cells(iCntr, 3) = "OVERDYEBLK") Then
            Rows(iCntr).Delete
        End If
    Next
End Sub
'This will move the PO orders that are considered to be completed, with 97% or higher
'shipped and no yardage ready to be shipped.

Sub Move_If_Ready()
    Dim xRg As Range
    Dim yRg As Range
    Dim xCell As Range
    Dim I As Long
    Dim J As Long
    Dim K As Long
    I = Worksheets("OPEN").UsedRange.Rows.Count
    J = Worksheets("CLOSE").UsedRange.Rows.Count
    If J = 1 Then
        If Application.WorksheetFunction.CountA(Worksheets("CLOSE").UsedRange) = 0 Then J = 0
    End If
    Set xRg = Worksheets("OPEN").Range("R2:R" & I)
    Set yRg = Worksheets("OPEN").Range("Q2:Q" & I)
    On Error Resume Next
    Application.ScreenUpdating = False
    
    For K = 1 To xRg.Count
        If CStr(IsEmpty(xRg(K))) Then
            Exit For
        Else
            If CStr(yRg(K).Value) = 0 Then
                If CStr(xRg(K).Value) > 0.96 Then
                    xRg(K).EntireRow.Copy Destination:=Worksheets("CLOSE").Range("A" & J + 1)
                    xRg(K).EntireRow.Delete
                    If CStr(xRg(K).Value) > 0.96 Then
                        K = K - 1
                    End If
                    J = J + 1
                End If
            End If
        End If
    Next
    Application.ScreenUpdating = True
End Sub


'This module will move all of the yards ready to ship to its respective ready list
Sub PopulateReadyList5202()
    Dim xRg As Range
    Dim yRg As Range
    Dim xCell As Range
    Dim I As Long
    Dim J As Long
    Dim K As Long
    I = Worksheets("OPEN").UsedRange.Rows.Count
    J = Worksheets("READY LIST 5202").UsedRange.Rows.Count
    
    For K = 1 To xRg.Count
        If CStr(IsEmpty(xRg(K))) Then
            Exit For
        Else
            If CStr(yRg.Value) > 0 Then
                xRg(K).EntireRow.Copy Destination:=Worksheets("READY LIST 5202")
                K = K - 1
            End If
    Next
End Sub

Sub PopulateReadyList5202D()
    Dim xRg As Range
    Dim yRg As Range
    Dim xCell As Range
    Dim I As Long
    Dim J As Long
    Dim K As Long
    
    I = Worksheets("OPEN").UsedRange.Rows.Count
    J = Worksheets("READY LIST 5202D").UsedRange.Rows.Count
    
    For K = 1 To xRg.Count
        If CStr(IsEmpty(xRg(K))) Then
            Exit For
        Else
            If CStr(yRg.Value) > 0 Then
                xRg(K).EntireRow.Copy Destination:=("READY LIST 5202")
                K = K - 1
            End If
        Next
End Sub


'Inserts the header for the CLOSE sheet
Sub insert_Header()
     Sheets("OPEN").Range("1:1").Copy Sheets("CLOSE").Range("1:1")
End Sub
'Using Vlookup, we will populate the reference table with the csheets that are ready to ship
Sub Csheet_Vlookup()
    Dim cSheet As Variant
    Dim wsCurrent As Worksheet
    Dim wsPrevious As Worksheet
    Dim rngSelection As Range
    
    'Initialization
    Set wsCurrent = ActiveSheet
    Set wsPrevious = wsCurrent.Previous
    Set rngSelection = ActiveCell
    
    'Error checking -- do nothing if not inthe correct column
    If Not rngSelection.Column = 27 Then
    
        MsgBox "Please select a cell in column AA.", vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If
    
    'Search to see if you can find the ID on the previous page
    Set varID = wsPrevious.Columns(1).Find(What:=wsCurrent.Cells(rngSelection.Row, 1).Value)
    
    'If it wasn't found, the results will be nothing.
    If varID Is Nothing Then
    
        rngSelection.Value = " - Not Found - "
    Else
    
        'Return the value in the appropriate row and column from the previous sheet
        rngSelection.Value = wsPrevious.Cells(varID.Row, 27).Value
        
    End If
    'Regardless, move to next cell
    wsCurrent.Cells(rngSelection.Row + 1, rngSelection.Column).Select
    
End Sub
'Makes the currently active sheet very hidden
Sub VeryHiddenActiveSheet()
    ActiveSheet.Visible = xlSheetVeryHidden
End Sub

