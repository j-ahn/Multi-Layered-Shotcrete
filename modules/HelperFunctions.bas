Attribute VB_Name = "HelperFunctions"
Public Function generateCumulativeArray(dataInput As Variant) As Variant

    Dim i                   As Long
    Dim dataReturn          As Variant
    Dim fromValue As Variant
    Dim toValue As Variant
    Dim arrLength As Variant
    
    fromValue = LBound(dataInput)
    toValue = UBound(dataInput)
    arrLength = toValue - fromValue + 1
    ReDim dataReturn(1 To arrLength, 1)
    
    dataReturn(1, 1) = dataInput(fromValue, 1)
    
    For i = 2 To arrLength
        dataReturn(i, 1) = dataReturn(i - 1, 1) + dataInput(i, 1)
    Next i
    
    generateCumulativeArray = dataReturn
    
End Function

Public Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Public Function arrlen(arr As Variant) As Integer

    arrlen = UBound(arr, 1) - LBound(arr, 1) + 1
    
End Function

Public Function linspace(ByVal a As Double, ByVal b As Double, ByVal num As Double, Optional ByVal endpoint As Boolean = True) As Double()
' Generate linearly spaced numbers
    Dim step As Double
    Dim i As Long
    Dim y() As Double

    ReDim y(1 To num, 1)

    If endpoint = True Then
        step = (b - a) / (num - 1)     ' with endpoint
    ElseIf endpoint = False Then
        step = (b - a) / num           ' without endpoint
    End If

    For i = 1 To num
        y(i, 1) = a + (i - 1) * step
    Next

    'Debug.Print step
    linspace = y

End Function



