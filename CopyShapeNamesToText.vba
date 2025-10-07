Sub CopyShapeNamesToText()
    Dim sld As Slide
    Dim shp As Shape
    Dim shpName As String

    ' Loop through each slide in the presentation
    For Each sld In ActivePresentation.Slides
        ' Loop through each shape in the slide
        For Each shp In sld.Shapes
            ' Get the name of the shape
            shpName = shp.Name
            ' Insert the name as text into the shape
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    shp.TextFrame.TextRange.Text = shpName
                Else
                    shp.TextFrame.TextRange.Text = shpName
                End If
            End If
        Next shp
    Next sld

    MsgBox "Shape names copied to shapes successfully!", vbInformation
End Sub
