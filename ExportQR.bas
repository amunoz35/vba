Attribute VB_Name = "Module3"
Sub RunJose()
    Call AutoFilterJose
    Call RenameJ
    Worksheets("REF").Activate
    Call VeryHiddenSheets
    Call ExportWorkbookJ
    Call ChangeRange
End Sub

Sub RunQCLab()
    Call AutoFilterQC
    Call RenameQ
    Worksheets("REF").Activate
    Call VeryHiddenSheets
    Call ExportWorkbookQ
    Call ChangeRangeQ
End Sub

Sub VeryHiddenSheets()
    ActiveSheet.Visible = xlSheetVeryHidden
End Sub
