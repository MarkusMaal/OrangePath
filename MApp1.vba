' Example app

Sub App1(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "1"
    Dim rndCol As Longd
    rndCol = RGB(Int(255 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
    Slide2.Shapes("WindowApp1_").TextFrame2.TextRange.Font.Glow.Color = rndCol
    Slide2.Shapes("WindowApp1_").TextFrame2.TextRange.Characters.Font.Line.ForeColor.RGB = rndCol
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
End Sub