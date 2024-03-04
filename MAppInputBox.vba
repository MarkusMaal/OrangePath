Option Explicit
' Input box app
Sub AppInputBox(ByVal Text As String, ByVal Title As String, Optional ByVal ToDesktop As Boolean = False)
    Dim backup As Integer
    backup = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    ActivePresentation.SlideShowWindow.View.GotoSlide (15)
    CleanPopups
    ActivePresentation.SlideShowWindow.View.GotoSlide (backup)
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "InputBox"
    Slide2.Shapes("WindowAppInputBox_").TextFrame.TextRange.Text = Text
    Slide2.Shapes("WindowTitleAppInputBox_").TextFrame.TextRange.Text = Title
    If ToDesktop Then ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
End Sub

Sub TestConfirm()
    Dim Shp As Shape
    Set Shp = Slide1.Shapes("Shape6AppInputBox:30")
    ConfirmInput Shp
End Sub

Sub ConfirmInput(Shp As Shape)
    'On Error GoTo Crash
    Dim Sld As Slide
    Dim Inp As String
    Inp = Slide1.AxTextBox.Text
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then
        Inp = Slide13.AxTextBox.Text
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
    Slide22.Shapes("Details").TextFrame.TextRange.Text = "Error details: INPUT_BOX_ERROR"
    ActivePresentation.SlideShowWindow.View.GotoSlide (22)
End Sub