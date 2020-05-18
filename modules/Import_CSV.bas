Attribute VB_Name = "Import_CSV"
Sub ImportCSVData()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

Sheets("Master").Select
Range("B2").Activate
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :=".", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        
Dim wb As Workbook
Dim wbCSV As Workbook
Dim myPath As String
Dim myFile As Variant
Dim fileType As String
Dim i As Integer



'Get Target Folder Path From User
 With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Source Folder"
        .AllowMultiSelect = False
        .Show
'  changed the following line - added the backslash to the path
        myPath = .SelectedItems(1) & "\"
    End With

'Specify file type
  fileType = "*.csv*"

'Target Path with file type
  myFile = Dir(myPath & fileType)

'Add Target Workbook
' Workbooks.Add
'  changed the following line - removed leading space in file name
' ActiveWorkbook.SaveAs Filename:= _
        myPath & "Total Results.xlsm", FileFormat:= _
        xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
'  changed the following line - removed leading space in file name
' Set wb = Workbooks.Open(myPath & "Total Results.xlsm")

'Loop through each Excel file in folder
'  changed the following line - to test myFile against a null string, rather than a value of 0
  Do While myFile <> ""
    Worksheets.Add(Before:=Worksheets("Master")).Name = "MN_" & i + 1
'changed the following line - from *.csv to myFile ** This was probably what was causing the error
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & myPath & myFile _
            , Destination:=ActiveSheet.Range("$A$1"))
            .Name = myFile
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 850
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
            1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False    'Error occurs here!
        End With
    i = i + 1
    myFile = Dir
    
  Loop

  
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    
Sheets("Master").Select
'Message Box when tasks are completed
  MsgBox "Import Complete"

End Sub

