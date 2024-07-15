Sub ListLayoutsAndSlides()
    Dim oPres As Presentation
    Dim oSld As Slide
    Dim oLayout As CustomLayout
    Dim layoutInfo As String
    Dim slideList As String
    Dim slideCount As Long
    Dim newSlide As Slide
    
    Set oPres = ActivePresentation
    layoutInfo = "Layouts and Corresponding Slides"
    
    ' Create a new slide with a blank layout
    Set newSlide = oPres.Slides.AddSlide(oPres.Slides.Count + 1, oPres.SlideMaster.CustomLayouts(1))
    newSlide.Shapes(1).TextFrame.TextRange.Text = layoutInfo
    newSlide.Shapes(1).TextFrame.TextRange.Font.Size = 11
    
    ' Set autofit text to resize shape
    newSlide.Shapes(1).TextFrame2.AutoSize = ppAutoSizeShapeToFitText
    
    ' Set text box width and height
    newSlide.Shapes(1).Width = 33.87 * 28.35 ' Convert cm to points
    newSlide.Shapes(1).Height = 19.05 * 28.35 ' Convert cm to points
    
    ' Center the text box horizontally
    newSlide.Shapes(1).Left = (oPres.PageSetup.SlideWidth - newSlide.Shapes(1).Width) / 2
    
    ' Move the text box to the top edge of the slide
    newSlide.Shapes(1).Top = 0
    
    For Each oLayout In oPres.Designs(1).SlideMaster.CustomLayouts
        layoutInfo = vbCrLf & "Layout: " & oLayout.Name & vbCr
        slideList = ""
        slideCount = 0
        
        For Each oSld In oPres.Slides
            If oSld.CustomLayout.Name = oLayout.Name Then
                slideList = slideList & oSld.SlideIndex & ", "
                slideCount = slideCount + 1
            End If
        Next oSld
        
        ' Remove the trailing comma and space
        If Len(slideList) > 0 Then
            slideList = Left(slideList, Len(slideList) - 2)
            layoutInfo = layoutInfo & "Slides: " & slideList & " (" & slideCount & " slides)"
        End If
        newSlide.Shapes(1).TextFrame.TextRange.Text = newSlide.Shapes(1).TextFrame.TextRange.Text & layoutInfo
    Next oLayout
    
    ' Set the text box fill color to white
    newSlide.Shapes(1).Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    ' Set the text color to black
    newSlide.Shapes(1).TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    
    ' Align the text to the left
    newSlide.Shapes(1).TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    
    ' Set the text box to the front
    newSlide.Shapes(1).ZOrder msoBringToFront
    
    MsgBox "Information added to a new slide at the end of the deck.", vbInformation, "Layouts and Slides"
End Sub
