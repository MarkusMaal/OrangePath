' Message box app
Sub AppMessage(ByVal Text As String, ByVal Title As String, ByVal MsgType As String, ByVal ToDesktop As Boolean)
    On Error GoTo Crash
    Dim backup As Integer
    backup = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    ActivePresentation.SlideShowWindow.View.GotoSlide (15)
    CleanPopups
    ActivePresentation.SlideShowWindow.View.GotoSlide (backup)
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Message"
    Slide2.Shapes("InfoAppMessage_").Visible = msoFalse
    Slide2.Shapes("ExclamationAppMessage_").Visible = msoFalse
    Slide2.Shapes("ErrorAppMessage_").Visible = msoFalse
    If MsgType = "Info" Then
        Slide2.Shapes("InfoAppMessage_").Visible = msoTrue
    ElseIf MsgType = "Exclamation" Then
        Slide2.Shapes("ExclamationAppMessage_").Visible = msoTrue
    ElseIf MsgType = "Error" Then
        Slide2.Shapes("ErrorAppMessage_").Visible = msoTrue
    End If
    Slide2.Shapes("WindowAppMessage_").TextFrame.TextRange.Text = Text
    Slide2.Shapes("Shape3AppMessage_").TextFrame.TextRange.Text = Title
    If ToDesktop Then ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
Done: d
    Exit Sub
Crash:
    Slide22.Shapes("Details").TextFrame.TextRange.Text = "Error details: MESSAGE_BOX_ERROR"
    ActivePresentation.SlideShowWindow.View.GotoSlide (22)
End Sub