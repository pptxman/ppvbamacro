Sub SortShapesInSelectionPane()
    Dim sld As Slide
    Dim shp As Shape
    Dim shapeNames() As String
    Dim i As Long, j As Long
    Dim tempName As String
    
    Set sld = ActiveWindow.View.Slide
    
    If sld.Shapes.Count < 2 Then
        MsgBox "There are not enough shapes to sort.", vbInformation
        Exit Sub
    End If
    
    ' Store all shape names
    ReDim shapeNames(1 To sld.Shapes.Count)
    For i = 1 To sld.Shapes.Count
        shapeNames(i) = sld.Shapes(i).Name
    Next i
    
    ' Sort names alphabetically (A → Z)
    For i = 1 To UBound(shapeNames) - 1
        For j = i + 1 To UBound(shapeNames)
            If UCase(shapeNames(i)) > UCase(shapeNames(j)) Then
                tempName = shapeNames(i)
                shapeNames(i) = shapeNames(j)
                shapeNames(j) = tempName
            End If
        Next j
    Next i
    
    ' Reorder shapes so that first alphabetically is on top of Selection Pane
    For i = UBound(shapeNames) To 1 Step -1
        sld.Shapes(shapeNames(i)).ZOrder msoBringToFront
    Next i
    
    MsgBox "Shapes sorted alphabetically!", vbInformation
End Sub