Sub vba_array()
    
    Dim iosArray()  As Variant
    iosArray = Array("c8", "c5", "a1", "a122", "a2", "a3", "a4", "a5", "a8", "a9", "a10", "a11")
    Dim lengthIosArray As Integer
    lengthIosArray = UBound(iosArray) - LBound(iosArray) + 1
    
    Dim androidArray() As Variant
    androidArray = Array("c1", "a1", "b2", "a3", "a4", "b2356", "b5", "b6", "b7", "a9", "b10", "a11", "a12", "a14")
    Dim lengthAndroidArray As Integer
    lengthAndroidArray = UBound(androidArray) - LBound(androidArray) + 1
    
    Dim bothArray() As Variant
    bothArray = Array()
    Dim lengtBothArray As Integer
    
    For i = 0 To (lengthIosArray - 1)
        lengtBothArray = UBound(bothArray) - LBound(bothArray) + 1
        ReDim Preserve bothArray(lengtBothArray)
        For j = 0 To (lengthAndroidArray - 1)
            Dim intResult As Integer
            intResult = StrComp(iosArray(i), androidArray(j))
            If intResult = 0 Then
                bothArray(i) = iosArray(i)
                iosArray(i) = ""
                androidArray(j) = ""
            End If
        Next j
    Next i
    
    Dim numberStartColumn As Integer
    
    numberStartColumn = 13
    
    Sheets("Sheet1").range("M" & numberStartColumn & ":M" & numberStartColumn + lengthIosArray - 1).Value = Application.WorksheetFunction.Transpose(iosArray)
    Sheets("Sheet1").range("N" & numberStartColumn & ":N" & numberStartColumn + lengthAndroidArray - 1).Value = Application.WorksheetFunction.Transpose(androidArray)
    Sheets("Sheet1").range("O" & numberStartColumn & ":O" & numberStartColumn + lengtBothArray).Value = Application.WorksheetFunction.Transpose(bothArray)
    
    Debug.Print lengthIosArray
    Debug.Print lengthAndroidArray
    Debug.Print lengtBothArray
    
End Sub
