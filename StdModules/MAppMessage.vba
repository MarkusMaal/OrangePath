' Message box app
Sub AppMessage(ByVal Text As String, ByVal Title As String, ByVal MsgType As String, ByVal ToDesktop As Boolean)
    'On Error GoTo Crash
    Dim backup As Integer
    backup = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    ActivePresentation.SlideShowWindow.View.GotoSlide (15)
    CleanPopups
    ActivePresentation.SlideShowWindow.View.GotoSlide (backup)
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Message"
    Slide2.Shapes("WindowAppMessage_").TextFrame.TextRange.Text = Text
    Slide2.Shapes("WindowTitleAppMessage_").TextFrame.TextRange.Text = Title
    If ToDesktop Then ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Dim Sld As Slide
    Set Sld = ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition)
    Dim AppID As String
    AppID = Slide1.Shapes("AppID").TextFrame.TextRange.Text
    Sld.Shapes("InfoAppMessage:" & AppID).Visible = msoFalse
    Sld.Shapes("ExclamationAppMessage:" & AppID).Visible = msoFalse
    Sld.Shapes("ErrorAppMessage:" & AppID).Visible = msoFalse
    If MsgType = "Info" Then
        Sld.Shapes("InfoAppMessage:" & AppID).Visible = msoTrue
    ElseIf MsgType = "Exclamation" Then
        Sld.Shapes("ExclamationAppMessage:" & AppID).Visible = msoTrue
    ElseIf MsgType = "Error" Then
        Sld.Shapes("ErrorAppMessage:" & AppID).Visible = msoTrue
    End If
Done:
    Exit Sub
Crash:
    Slide22.Shapes("Details").TextFrame.TextRange.Text = "Error details: MESSAGE_BOX_ERROR"
    ActivePresentation.SlideShowWindow.View.GotoSlide (22)
End Sub

