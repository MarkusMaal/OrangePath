' Help app (Generated from devCreateApp)

' This is executed when the application is launched
Sub AppHelp(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Help"
    Slide2.Shapes("AppHelp").Visible = msoTrue
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide2.Shapes("AppHelp").Visible = msoFalse
End Sub

' This gets executed when a user clicks a file, which is associated with this application
Sub AssocHelp(Shp As Shape)
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Help"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
End Sub

' This gets executed when a user clicks icon of a file, which is associated with this application
Sub AssocIHelp(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocHelp Slide1.Shapes(ShapeName)
End Sub
