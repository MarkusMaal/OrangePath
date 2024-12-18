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

Sub AppHelpClickTopic(Shp As Shape)
    Dim T As String
    Dim F As String
    T = Shp.TextFrame.TextRange.Text
    F = "/Defaults/Help/"
    If T = "Window management" Then
        F = F & "Windows"
    ElseIf T = "Recovery mode" Then
        F = F & "Recovery"
    ElseIf T = "Installing software and updates" Then
        F = F & "Flashing"
    ElseIf T = "Calculator" Then
        F = F & "Calc"
    ElseIf T = "Task Manager" Then
        F = F & "Taskmgr"
    ElseIf T = "Guess the number" Then
        F = F & "Guess"
    ElseIf T = "Pixel Paint" Then
        F = F & "Paint"
    ElseIf T = "Notepad" Then
        F = F & "Notes"
    Else
        F = F & T
    End If
    F = F & ".hlp"
    AppModalHelpView F
End Sub

Sub AppHelpSizeChanged(AppID As String)
    
End Sub