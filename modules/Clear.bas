Attribute VB_Name = "Clear"
Sub ClearMN()
    Dim xWs As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each xWs In Application.ActiveWorkbook.Worksheets
            If xWs.Name = "MN_1" Then
                Sheets("MN_1").Select
                Range("A1:D1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.ClearContents
            ElseIf xWs.Name = "MN_2" Then
                Sheets("MN_2").Select
                Range("A1:D1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.ClearContents
            ElseIf xWs.Name = "MN_3" Then
                Sheets("MN_3").Select
                Range("A1:D1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.ClearContents
            ElseIf xWs.Name = "MN_4" Then
                Sheets("MN_4").Select
                Range("A1:D1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.ClearContents
            ElseIf xWs.Name = "MN_5" Then
                Sheets("MN_5").Select
                Range("A1:D1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.ClearContents
            ElseIf xWs.Name = "MN_6" Then
                Sheets("MN_6").Select
                Range("A1:D1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.ClearContents
            ElseIf xWs.Name = "MN_7" Then
                Sheets("MN_7").Select
                Range("A1:D1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.ClearContents
            ElseIf xWs.Name = "MN_8" Then
                Sheets("MN_8").Select
                Range("A1:D1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.ClearContents
            End If
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

Sub ClearRS2()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    
    Sheets("RS2_Export").Select
    Cells.Select
    Selection.ClearContents
    
    Sheets("Master").Select
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

