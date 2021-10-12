Attribute VB_Name = "Module2"
Sub RunReport()
    'Module 1
    Application.ScreenUpdating = False
    Call InsertTotals
    Call CopySheet
    Call ColorYarns
    Call FormatTotals
    Call ClearTitle
    Worksheets("Greige Goods").Visible = xlSheetVisible
    Worksheets("ANTEX GREIGE GOODS LOCATION").Activate
    Call HideSheet
    Call ResetRange
    Call ExportWorkbook
    Call CloseActiveBook
    Application.ScreenUpdating = True
End Sub
