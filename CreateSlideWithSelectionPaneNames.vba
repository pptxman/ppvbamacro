Sub CreateSlideWithSelectionPaneNames()
    Dim slide As slide
    Dim shape As shape
    Dim shapeNames As String
    Dim newSlide As slide
    Dim textBox As shape
    
    ' Loop through each slide
    For Each slide In ActivePresentation.Slides
        ' Loop through each shape on the slide
        For Each shape In slide.Shapes
            ' Add shape name to the list
            shapeNames = shapeNames & shape.Name & vbCrLf
        Next shape
    Next slide
    
    ' Add a new slide at the end of the presentation
    Set newSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutText)
    
    ' Add a text box to the new slide and set its text to the list of shape names
    Set textBox = newSlide.Shapes.Placeholders(2)
    textBox.TextFrame.TextRange.Text = shapeNames
    
    ' Notify the user
    MsgBox "New slide created with a list of shape names.", vbInformation
End Sub
