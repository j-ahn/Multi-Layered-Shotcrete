Attribute VB_Name = "Print_"
Sub Print_PDF()

    Path = ActiveSheet.Range("J1")
        
    Filename = ActiveSheet.Range("J2")
    WorkbookName = ActiveSheet.Range("C12")
    
    Range("A1").Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        Path & "\" & WorkbookName & Filename & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False


End Sub



Sub Print_All()
Sheets("Master").Activate
Set Rng = Range("Interaction_Diagram_Stage_No.")


For Each Row In Rng.Rows
  For Each cell In Row.Cells
    If IsEmpty(cell) = False Then
        Range("ActiveMNDiagramNumber").Value = cell.Value
        UpdateDropDownList
        
        'Plot the stage that has the largest number of yielded elements
        MaxNumberOfYieldedElements = 0
        StageIndex = 1
        For i = 1 To 5
            NumberOfYieldedElements = Cells(2, 8).Offset(0, 6 + 4 * (i - 1)).Value
            If NumberOfYieldedElements > MaxNumberOfYieldedElements Then
                MaxNumberOfYieldedElements = NumberOfYieldedElements
                StageIndex = i
            End If
        Next
        
        Sheets("Master").StageDropDown.ListIndex = StageIndex - 1
        
        UpdateLinerPlot
        Print_PDF
    End If
    i = i + 1
  Next cell
Next Row

End Sub




