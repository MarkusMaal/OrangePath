' Guess the number app


Sub AppGuess(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Guess"
    
    Slide2.Shapes("Shape11AppGuess_").TextFrame.TextRange.Text = CStr(Int(100 * Rnd))
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide1.Shapes("WindowTitleAppGuess:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "Guess the number (Guesses: 0)"
    UpdateTime
End Sub

Sub SubtractGuess1(Shp As Shape)
    AppID = GetAppID(Shp)
    FrameText = Split(Slide1.Shapes("WindowAppGuess:" + AppID).TextFrame.TextRange.Text, ":")
    FrameNumber = Split(FrameText(1), " ")
    CurrentNumber = CInt(FrameNumber(1)) - 1
    Slide1.Shapes("WindowAppGuess:" + AppID).TextFrame.TextRange.Text = FrameText(0) & ": " & CStr(CurrentNumber)
End Sub


Sub SubtractGuess10(Shp As Shape)
    AppID = GetAppID(Shp)
    FrameText = Split(Slide1.Shapes("WindowAppGuess:" + AppID).TextFrame.TextRange.Text, ":")
    FrameNumber = Split(FrameText(1), " ")
    CurrentNumber = CInt(FrameNumber(1)) - 10
    Slide1.Shapes("WindowAppGuess:" + AppID).TextFrame.TextRange.Text = FrameText(0) & ": " & CStr(CurrentNumber)
End Sub

Sub AddGuess1(Shp As Shape)
    AppID = GetAppID(Shp)
    FrameText = Split(Slide1.Shapes("WindowAppGuess:" + AppID).TextFrame.TextRange.Text, ":")
    FrameNumber = Split(FrameText(1), " ")
    CurrentNumber = CInt(FrameNumber(1)) + 1
    Slide1.Shapes("WindowAppGuess:" + AppID).TextFrame.TextRange.Text = FrameText(0) & ": " & CStr(CurrentNumber)
End Sub

Sub AddGuess10(Shp As Shape)
    AppID = GetAppID(Shp)
    FrameText = Split(Slide1.Shapes("WindowAppGuess:" + AppID).TextFrame.TextRange.Text, ":")
    FrameNumber = Split(FrameText(1), " ")
    CurrentNumber = CInt(FrameNumber(1)) + 10
    Slide1.Shapes("WindowAppGuess:" + AppID).TextFrame.TextRange.Text = FrameText(0) & ": " & CStr(CurrentNumber)
End Sub

Sub GuessNo(Shp As Shape)
    AppID = GetAppID(Shp)
    If Slide1.Shapes("ButtonShape10AppGuess:" + AppID).TextFrame.TextRange.Text = "Guess" Then
        FrameText = Split(Slide1.Shapes("WindowAppGuess:" + AppID).TextFrame.TextRange.Text, ":")
        FrameNumber = Split(FrameText(1), " ")
        CurrentNumber = CInt(FrameNumber(1))
        CorrectNumber = CInt(Slide1.Shapes("Shape11AppGuess:" + AppID).TextFrame.TextRange.Text)
        WindowTitle = Slide1.Shapes("WindowTitleAppGuess:" + AppID).TextFrame.TextRange.Text
        WindowTitleSplit = Split(WindowTitle, ": ")
        WindowTitleSplit2 = Split(WindowTitleSplit(1), ")")
        guesses = CInt(WindowTitleSplit2(0)) + 1
        Slide1.Shapes("WindowTitleAppGuess:" + AppID).TextFrame.TextRange.Text = "Guess the number (Guesses: " & guesses & ")"
        If CurrentNumber > CorrectNumber Then
            Slide1.Shapes("WindowAppGuess:" + AppID).TextFrame.TextRange.Text = "Lower!" & vbNewLine & "Your guess: " & CStr(CurrentNumber)
        ElseIf CurrentNumber < CorrectNumber Then
            Slide1.Shapes("WindowAppGuess:" + AppID).TextFrame.TextRange.Text = "Higher!" & vbNewLine & "Your guess: " & CStr(CurrentNumber)
        Else
            Slide1.Shapes("ButtonShape10AppGuess:" + AppID).TextFrame.TextRange.Text = "Again"
            Slide1.Shapes("WindowAppGuess:" + AppID).TextFrame.TextRange.Text = "You win!" & vbNewLine & "Your guess: " & CStr(CurrentNumber)
        End If
    Else
        Slide1.Shapes("WindowTitleAppGuess:" + AppID).TextFrame.TextRange.Text = "Guess the number (Guesses: 0)"
        Slide1.Shapes("WindowAppGuess:" + AppID).TextFrame.TextRange.Text = "I’m thinking of a number between 1 and 100. What number am I thinking about?" & vbNewLine & "Your guess: 50"
        Slide1.Shapes("ButtonShape10AppGuess:" + AppID).TextFrame.TextRange.Text = "Guess"
        Slide1.Shapes("Shape11AppGuess:" + AppID).TextFrame.TextRange.Text = CStr(Int(100 * Rnd))
    End If
End Sub
