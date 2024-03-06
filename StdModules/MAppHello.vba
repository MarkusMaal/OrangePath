' Hello app (Generated from devCreateApp)

' This is executed when the application is launched
Sub AppHello(Shp As shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Hello"
    Slide2.Shapes("AppHello").Visible = msoTrue
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide2.Shapes("AppHello").Visible = msoFalse
End Sub

' This gets executed when a user clicks a file, which is associated with this application
Sub AssocHello(Shp As shape)
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Hello"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
End Sub

' This gets executed when a user clicks icon of a file, which is associated with this application
Sub AssocIHello(Shp As shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocHello Slide1.Shapes(ShapeName)
End Sub
