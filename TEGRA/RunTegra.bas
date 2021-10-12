Attribute VB_Name = "Module2"
'Written by: Alexander Munoz
Sub RunTegraReport()
    Call Test_Copy
    Call sbDelete_Rows_If_Cell_Contains_Error
    Call remove_Swaps
    Call remove_Overdye
    Call Move_If_Ready
    Call insert_Header
    Worksheets("READY LIST 5202D").Activate
    [C2] = Now
    [D1] = Now
    [E2] = Now
    Range("C2").Select
    Worksheets("READY LIST 5202D").Range("E2").NumberFormat = "m/d/yyyy"
    Worksheets("READY LIST 5202D").Range("D1").NumberFormat = "m/d/yyyy"
    Worksheets("READY LIST 5202D").Range("C2").NumberFormat = "h:mm AM/PM"
    
   ' Call PopulateReadyList5202D
    
    Worksheets("READY LIST 5202").Activate
    [C2] = Now
    [D1] = Now
    [E2] = Now
    Range("C2").Select
    Worksheets("READY LIST 5202").Range("E2").NumberFormat = "m/d/yyyy"
    Worksheets("READY LIST 5202").Range("D1").NumberFormat = "m/d/yyyy"
    Worksheets("READY LIST 5202").Range("C2").NumberFormat = "h:mm AM/PM"
    
   ' Call PopulateReadyList5202
    Worksheets("5202_5202D").Activate
    Call VeryHiddenActiveSheet
    Worksheets("5202ref").Activate
    Call VeryHiddenActiveSheet
    Call SplitWorkbook
    
End Sub
