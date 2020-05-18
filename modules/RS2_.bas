Attribute VB_Name = "RS2_"
Sub RS2_Export()
Attribute RS2_Export.VB_ProcData.VB_Invoke_Func = " \n14"
        
    databook = ActiveWorkbook.Name

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Sheets("RS2_Export").Select
    
    Columns("A:A").Select
    Selection.Copy
    Range("V1").Select
    ActiveSheet.Paste
    Columns("V:V").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("V1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:= _
        ".", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), _
        Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1)), TrailingMinusNumbers:=True

    Dim cl As Range, i&
    Set cl = Range("A1:A" & Cells(Rows.Count, "A").End(xlUp).Row)
    i = cl.Find("*", , xlValues, , xlByRows, xlPrevious).Row + 1
   
    Dim colrow As Integer
    
    colrow = 1
    
    Do Until colrow = i
                                                 
        If Workbooks(databook).Worksheets("RS2_Export").Cells(colrow, "AC").Value <> 0 Then
            If IsNumeric(Workbooks(databook).Worksheets("RS2_Export").Cells(colrow, "AC").Value) = True Then
                StageName = Workbooks(databook).Worksheets("RS2_Export").Cells(colrow, "AC").Value
            ElseIf IsNumeric(Workbooks(databook).Worksheets("RS2_Export").Cells(colrow, "AB").Value) = True Then
                StageName = Workbooks(databook).Worksheets("RS2_Export").Cells(colrow, "AB").Value
                ' If RS2 Stage number doesn't end up in Column U, modify this if statement
            End If
            
        End If
        
        If IsNumeric(Workbooks(databook).Worksheets("RS2_Export").Cells(colrow, "V").Value) = True Then
            Workbooks(databook).Worksheets("RS2_Export").Cells(colrow, "U").Value = StageName
        End If
        
        colrow = colrow + 1
    Loop

    Sheets("RS2_Export").Select
    Columns("V:V").Select
    Selection.Copy

    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Temp"
    End With
    
    ActiveSheet.Paste
    
    Sheets("RS2_Export").Select
    Columns("G:G").Select
    Selection.Copy
    
    Sheets("Temp").Select
    ActiveSheet.Paste Destination:=Worksheets("Temp").Range("B:B")
    
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:$B$1291").RemoveDuplicates Columns:=1, Header:=xlNo
    Selection.AutoFilter
    Range("A1:B1291").Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlNo
    ActiveSheet.Range("$A$1:$A$1291").AutoFilter Field:=1, Criteria1:=">1", _
        Operator:=xlAnd
    Selection.Copy

    Workbooks(databook).Sheets("Master").Select
    Range("H5").Select
    ActiveSheet.Paste
    
    Sheets("Temp").Select
    ActiveWindow.SelectedSheets.Delete
    
    Workbooks(databook).Sheets("Master").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    UpdateDropDownList
End Sub
Sub MatlabInputFile()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim flder As FileDialog
    Dim test As Integer
    
    Set flder = Application.FileDialog(msoFileDialogFolderPicker)
    With flder
    .Title = "Select the folder to place script file within"
    .AllowMultiSelect = True
    test = .Show
    
    If test <> -1 Then GoTo NextCode
    folderpath = .SelectedItems(1)
    End With
NextCode:

    Sheets("RS2_Staging").Select
    Range("B71:I80").Select
    Selection.Copy

    Workbooks.Add
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Sheets("output").Select
    myFile = folderpath & "\MatlabInput.csv"
    ActiveWorkbook.SaveAs Filename:=myFile, FileFormat:= _
        xlCSV, CreateBackup:=False
    ActiveWorkbook.Close
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
