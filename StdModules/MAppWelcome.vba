' Welcome app (Generated from devCreateApp)

' This is executed when the application is launched
Sub AppWelcome(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Welcome"
    Slide2.Shapes("AppWelcome").Visible = msoTrue
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide2.Shapes("AppWelcome").Visible = msoFalse
    ' Remove welcome from startup applications after launch
    If GetSysConfig("Autorun") <> "*" Then
        SaveSysConfig "Autorun", Replace(GetSysConfig("Autorun"), "Welcome", "")
    End If
End Sub

' This gets executed when a user clicks a file, which is associated with this application
Sub AssocWelcome(Shp As Shape)
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Welcome"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
End Sub

' This gets executed when a user clicks icon of a file, which is associated with this application
Sub AssocIWelcome(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocWelcome Slide1.Shapes(ShapeName)
End Sub

Sub AppWelcomeButtonClicked(Shp As Shape)
    Dim App As String
    Dim ButtonText As String
    ButtonText = Shp.TextFrame.TextRange.Text
    If ButtonText = "Open ‘Settings’" Then
        App = "Settings"
    ElseIf ButtonText = "Open ‘Files’" Then
        App = "Files"
    ElseIf ButtonText = "Open ‘Help’" Then
        App = "Help"
    ElseIf ButtonText = "Tutorials" Then
        ActivePresentation.FollowHyperlink "https://github.com/MarkusMaal/OrangePath/"
        Exit Sub
    End If
    On Error GoTo FailedOpen
    With Slide1.Shapes

        .AddShape(msoShapeRectangle, 0, 0, 0, 0).Name = "WelcomeDummy1"
    
        .AddShape(msoShapeRectangle, 0, 0, 0, 0).Name = "WelcomeDummy2"
    
        With .Range(Array("WelcomeDummy1", "WelcomeDummy2")).Group
            .Visible = msoFalse
            .Name = "WelcomeDummy"
        End With
    
    End With
    Application.Run "App" & App, Slide1.Shapes("WelcomeDummy1")
FailedOpen:
    Slide1.Shapes("WelcomeDummy").Delete
End Sub