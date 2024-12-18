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
    UpdateTime
End Sub


Sub AppShellSizeChanged(AppID As String, Optional IsFullScreen As Boolean = False)
    ' Makes sure that the size of the textbox stays the same when resizing window
    Dim C1 As Shape
    Dim T1 As Shape
    Dim W As Shape
    
    Set C1 = Slide1.Shapes("OutputAppShell:" & AppID)
    Set T1 = Slide1.Shapes("AXTextBox1AppShell:" & AppID)
    If Not IsFullScreen Then
        Set W = Slide1.Shapes("WindowAppShell:" & AppID)
    Else
        Set W = Slide1.Shapes("AnimationRect")
    End If
    
    T1.Height = 25.5337
    T1.Top = W.Top + W.Height - T1.Height
    C1.Height = W.Height - T1.Height - 1.448746
    
    If AAX Then Slide1.AxTextBox.Visible = False
    ApplyTbAttribs T1
End Sub