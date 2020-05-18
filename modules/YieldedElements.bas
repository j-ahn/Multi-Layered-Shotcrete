Attribute VB_Name = "YieldedElements"
Public Function PtInPoly(Xcoord As Double, Ycoord As Double, Polygon As Variant) As Variant
  Dim x As Long, NumSidesCrossed As Long, m As Double, b As Double, Poly As Variant
  Poly = Polygon
  For x = LBound(Poly) To UBound(Poly) - 1
    If Poly(x, 1) > Xcoord Xor Poly(x + 1, 1) > Xcoord Then
      m = (Poly(x + 1, 2) - Poly(x, 2)) / (Poly(x + 1, 1) - Poly(x, 1))
      b = (Poly(x, 2) * Poly(x + 1, 1) - Poly(x, 1) * Poly(x + 1, 2)) / (Poly(x + 1, 1) - Poly(x, 1))
      If m * Xcoord + b > Ycoord Then NumSidesCrossed = NumSidesCrossed + 1
    End If
  Next
  PtInPoly = CBool(NumSidesCrossed Mod 2)
End Function
Function setChartAxis(sheetName As String, chartName As String, MinOrMax As String, _
    ValueOrCategory As String, PrimaryOrSecondary As String, Value As Variant)

'Create variables
Dim cht As Chart
Dim valueAsText As String

'Set the chart to be controlled by the function
Set cht = Application.Caller.Parent.Parent.Sheets(sheetName) _
    .ChartObjects(chartName).Chart

'Set Value of Primary axis
If (ValueOrCategory = "Value" Or ValueOrCategory = "Y") _
    And PrimaryOrSecondary = "Primary" Then

    With cht.Axes(xlValue, xlPrimary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'Set Category of Primary axis
If (ValueOrCategory = "Category" Or ValueOrCategory = "X") _
    And PrimaryOrSecondary = "Primary" Then

    With cht.Axes(xlCategory, xlPrimary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'Set value of secondary axis
If (ValueOrCategory = "Value" Or ValueOrCategory = "Y") _
    And PrimaryOrSecondary = "Secondary" Then

    With cht.Axes(xlValue, xlSecondary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'Set category of secondary axis
If (ValueOrCategory = "Category" Or ValueOrCategory = "X") _
    And PrimaryOrSecondary = "Secondary" Then
    With cht.Axes(xlCategory, xlSecondary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'If is text always display "Auto"
If IsNumeric(Value) Then valueAsText = Value Else valueAsText = "Auto"

'Output a text string to indicate the value
setChartAxis = ValueOrCategory & " " & PrimaryOrSecondary & " " _
    & MinOrMax & ": " & valueAsText

End Function
Sub UpdateLinerPlot()
Sheets("Master").Activate

RowCounter = 5

Do While IsEmpty(Cells(RowCounter, 8))

RowCounter = RowCounter + 1
Loop

RowStart = RowCounter

Do While Not IsEmpty(Cells(RowCounter, 8))

RowCounter = RowCounter + 1
Loop


SelectedRS2Stage = Sheets("Master").StageDropDown.ListIndex

RowEnd = RowCounter - 1

'Update Liner_MN Plot

ActiveSheet.ChartObjects("Liner_MN").Activate
For Each s In ActiveChart.SeriesCollection
   s.Delete
Next s
scount = 1
ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(scount).Name = "'Liners'"
ActiveChart.SeriesCollection(scount).XValues = "='" & ActiveSheet.Name & "'!$AE" & RowStart & ":$AE" & RowEnd
ActiveChart.SeriesCollection(scount).Values = "='" & ActiveSheet.Name & "'!$AF" & RowStart & ":$AF" & RowEnd
ActiveChart.SeriesCollection(scount).Border.Color = RGB(150, 150, 150)

scount = 2

RowCounter = RowStart

NodeNumber = Cells(RowCounter, 8).Value
Do While Not IsEmpty(Cells(RowCounter, 8))
NodeNumber = Cells(RowCounter, 8).Value
InEnvelope = Cells(RowCounter, 8).Offset(0, 6 + 4 * SelectedRS2Stage).Value
StartX = Cells(RowCounter, 8).Offset(0, 23).Value 'Column StartX
StartY = Cells(RowCounter, 8).Offset(0, 24).Value 'Column StartY
EndX = Cells(RowCounter, 8).Offset(0, 25).Value 'Column EndX
EndY = Cells(RowCounter, 8).Offset(0, 26).Value 'Column EndY

If Not InEnvelope Then
If scout = 255 Then

MsgBox "Number of yielded elements should be less than 255"
Exit Sub
End If

ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(scount).Name = "='" & ActiveSheet.Name & "'!" & Cells(RowCounter, 8).Address(RowAbsolute:=False, ColumnAbsolute:=False)
ActiveChart.SeriesCollection(scount).XValues = Array(StartX, EndX)
ActiveChart.SeriesCollection(scount).Values = Array(StartY, EndY)
ActiveChart.SeriesCollection(scount).Border.Color = RGB(255, 0, 0)
scount = scount + 1
End If


RowCounter = RowCounter + 1
Loop


'Update Liner_NV Plot
ActiveSheet.ChartObjects("Liner_NV").Activate
For Each s In ActiveChart.SeriesCollection
   s.Delete
Next s
scount = 1
ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(scount).Name = "'Liners'"
ActiveChart.SeriesCollection(scount).XValues = "='" & ActiveSheet.Name & "'!$AE" & RowStart & ":$AE" & RowEnd
ActiveChart.SeriesCollection(scount).Values = "='" & ActiveSheet.Name & "'!$AF" & RowStart & ":$AF" & RowEnd
ActiveChart.SeriesCollection(scount).Border.Color = RGB(150, 150, 150)

scount = 2

RowCounter = RowStart

NodeNumber = Cells(RowCounter, 8).Value
Do While Not IsEmpty(Cells(RowCounter, 8))
NodeNumber = Cells(RowCounter, 8).Value
InEnvelope = Cells(RowCounter, 1).Offset(0, 79 + SelectedRS2Stage).Value
NeedMinReo = Cells(RowCounter, 1).Offset(0, 84 + SelectedRS2Stage).Value
StartX = Cells(RowCounter, 8).Offset(0, 23).Value 'Column StartX
StartY = Cells(RowCounter, 8).Offset(0, 24).Value 'Column StartY
EndX = Cells(RowCounter, 8).Offset(0, 25).Value 'Column EndX
EndY = Cells(RowCounter, 8).Offset(0, 26).Value 'Column EndY

If Not InEnvelope Then
If scout = 255 Then

MsgBox "Number of yielded elements should be less than 255"
Exit Sub
End If

ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(scount).Name = "='" & ActiveSheet.Name & "'!" & Cells(RowCounter, 8).Address(RowAbsolute:=False, ColumnAbsolute:=False)
ActiveChart.SeriesCollection(scount).XValues = Array(StartX, EndX)
ActiveChart.SeriesCollection(scount).Values = Array(StartY, EndY)
ActiveChart.SeriesCollection(scount).Border.Color = RGB(255, 0, 0)
scount = scount + 1
End If

If Not NeedMinReo Then
If scout = 255 Then

MsgBox "Number of yielded elements should be less than 255"
Exit Sub
End If

ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(scount).Name = "='" & ActiveSheet.Name & "'!" & Cells(RowCounter, 8).Address(RowAbsolute:=False, ColumnAbsolute:=False)
ActiveChart.SeriesCollection(scount).XValues = Array(StartX, EndX)
ActiveChart.SeriesCollection(scount).Values = Array(StartY, EndY)
ActiveChart.SeriesCollection(scount).Border.Color = RGB(255, 165, 0)
scount = scount + 1
End If


RowCounter = RowCounter + 1
Loop

End Sub


Public Sub UpdateDropDownList()
With Sheets("Master").StageDropDown
.Clear

Dim InteractionDiagramStage As String
Dim Rng As Range
Dim StageList(5) As Integer
Dim Row As Range
Dim cell As Range
Dim i As Integer

InteractionDiagramStage = Range("ActiveMNDiagramNumber").Value

Set Rng = Range("Interaction_Diagram_Stage_No.")

For Each Row In Rng.Rows
  For Each cell In Row.Cells
    If InteractionDiagramStage = cell.Value Then
        For i = 1 To 5
            Stage = Range(cell.Address).Offset(0, i + 1).Value
            If IsNumeric(Stage) Then
                StageList(i) = Range(cell.Address).Offset(0, i + 1).Value
            Else
                GoTo end_of_for
            End If
        Next
        GoTo end_of_for
    End If
  Next cell
Next Row

end_of_for:

For i = 1 To 5

If StageList(i) > 0 Then
.AddItem StageList(i)
End If

Next

SelectedRS2Stage = Sheets("Master").StageDropDown.ListIndex

If SelectedRS2Stage = -1 Then

Sheets("Master").StageDropDown.ListIndex = 0

End If

End With


UpdateLinerPlot

End Sub




