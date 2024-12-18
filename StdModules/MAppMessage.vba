' Message box app
Sub AppMessage(ByVal Text As String, ByVal Title As String, ByVal MsgType As String, ByVal ToDesktop As Boolean)
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 38 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 40
        Slide41.Shapes("FailDetails").TextFrame.TextRange.Text = "Error details: " & vbNewLine & vbNewLine & Title & vbNewLine & Text
        Exit Sub
    End If
    On Error Resume Next
    Backup = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    ActivePresentation.SlideShowWindow.View.GotoSlide (15)
    WaitForTransitions
    CleanPopups
    ActivePresentation.SlideShowWindow.View.GotoSlide (Backup)
    WaitForTransitions
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Message"
    Slide2.Shapes("WindowAppMessage_").TextFrame.TextRange.Text = Text
    If ToDesktop Then ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    WaitForTransitions
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
    Sld.Shapes("RegularApp:" & AppID).Left = 0
    Sld.Shapes("RegularApp:" & AppID).Top = 0
    Sld.Shapes("WindowTitleAppMessage:" & AppID).TextFrame.TextRange.Text = Title
    If AAX Then Slide1.AxTextBox.Visible = False
Done:
    Exit Sub
Crash:
    OSCrash "MESSAGE_BOX_ERROR", Err
End Sub

Sub WaitForTransitions()
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition < 1 Then
        ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition).SlideShowTransition.Duration = 0
    End If
End Sub