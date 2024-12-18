Option Explicit
' Input box app
Sub AppInputBox(ByVal Text As String, ByVal Title As String, Optional ByVal ToDesktop As Boolean = False)
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "InputBox"
    Slide2.Shapes("WindowAppInputBox_").TextFrame.TextRange.Text = Text
    Dim Backup As Integer
    Backup = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    If Backup = 13 Then
        CreateNewWindow
        ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition).Shapes("WindowTitleAppInputBox:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = Title
        Exit Sub
    End If
    If Backup <> 13 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide (15)
        CleanPopups
        ActivePresentation.SlideShowWindow.View.GotoSlide (Backup)
    End If
    If ToDesktop Then ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition).Shapes("WindowTitleAppInputBox:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = Title
    UpdateTime
    FocusWindow Slide1.Shapes("AppID").TextFrame.TextRange.Text
    If AAX Then
        Slide1.AxTextBox.Visible = True
    End If
End Sub

Sub ConfirmInput(Shp As Shape)
    'On Error GoTo Crash
    Dim Sld As Slide
    Dim Inp As String
    If AAX Then
        Inp = Slide1.AxTextBox.Text
        If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then
            Inp = Slide13.AxTextBox.Text
        End If
    End If
    SetVar "InputValue", Inp
    If CheckVars("%Macro%") <> "" And CheckVars("%Macro%") <> "%Macro%" Then
        Application.Run CheckVars("%Macro%"), Shp
    End If
    UnsetVar "Macro"
    CloseWindow Shp
Done:
    Exit Sub
Crash:
    OSCrash "INPUT_BOX_ERROR", Err
End Sub
