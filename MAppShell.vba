' Shell app


Sub AppShell(Shp As Shape)
    Shp.ParentGroup.Delete
    If Slide1.Shapes("Username").TextFrame.TextRange.Text = "Guest" Then
        AppMessage "Guests can't access the system shell", "Access denied", "Error", False
        'ActivePresentation.SlideShowWindow.View.GotoSlide (4)
        Exit Sub
    End If
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Shell"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
End Sub
