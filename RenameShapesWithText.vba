Sub RenameShapesWithText()
    Dim sld As Slide
    Dim shp As Shape
    Dim shpText As String

    ' Loop through each slide in the presentation
    For Each sld In ActivePresentation.Slides
        ' Loop through each shape in the slide
        For Each shp In sld.Shapes
            ' Check if the shape has a text frame and contains text
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    shpText = shp.TextFrame.TextRange.Text
                    ' Use the text to rename the shape
                    ' Truncate the text to 255 characters if it is too long
                    If Len(shpText) > 255 Then
                        shpText = Left(shpText, 255)
                    End If
                    ' Replace invalid characters with underscores
                    shpText = Replace(shpText, ":", "_")
                    shpText = Replace(shpText, "\", "_")
                    shpText = Replace(shpText, "/", "_")
                    shpText = Replace(shpText, "*", "_")
                    shpText = Replace(shpText, "?", "_")
                    shpText = Replace(shpText, """", "_")
                    shpText = Replace(shpText, "<", "_")
                    shpText = Replace(shpText, ">", "_")
                    shpText = Replace(shpText, "|", "_")
                    
                    ' Rename the shape
                    shp.Name = shpText
                End If
            End If
        Next shp
    Next sld

    MsgBox "Shapes renamed successfully!", vbInformation
End Sub
