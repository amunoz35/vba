Attribute VB_Name = "Module3"
'This will move the worksheets that have been output into a new workbook and
'can prompt the user to save the documents; if not, the new document will have
'the same name
'https://www.extendoffice.com/documents/excel/785-excel-save-export-sheet-as-new-workbook.html#:~:text=1%20Hold%20down%20the%20ALT%20%2B%20F11%20keys%2C,the%20new%20exported%20workbooks%2C%20and%20all%20of%20


Sub SplitWorkbook()

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

If Val(Application.Version) < 12 Then
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

Sub saveVisibleSheetsAsXLSM()
    Const exportPath = "C:\Users\alex\Documents\Maria S OUT\"
    Dim ws As Worksheet, wbNew As Workbook
    For Each ws In ThisWorkbook.Sheets
        If ws.Visible Then
            Debug.Print "Exporting: " & ws.Name
            ws.Copy
            Set wbNew = Application.ActiveWorkbook
            wbNew.SaveAs exportPath & ws.Name & ".xlsx", 51
            wbNew.Close
            Set wbNew = Nothing
        End If
    Next ws
    Set ws = Nothing
End Sub
