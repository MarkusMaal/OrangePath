' OrangePath Session Manager
Sub InitUserspace()
    On Error GoTo Crash
    ResetWindows
    With Slide1.Master.Theme
        .ThemeColorScheme(msoThemeAccent1) = RGB(68, 114, 196)
        .ThemeColorScheme(msoThemeAccent2) = RGB(237, 125, 49)
        .ThemeColorScheme(msoThemeAccent3) = RGB(165, 165, 165)
        .ThemeColorScheme(msoThemeAccent4) = RGB(255, 192, 0)
        .ThemeColorScheme(msoThemeAccent5) = RGB(91, 155, 213)
        .ThemeColorScheme(msoThemeAccent6) = RGB(112, 173, 71)
        .ThemeColorScheme(msoThemeDark1) = RGB(0, 0, 0)
        .ThemeColorScheme(msoThemeDark2) = RGB(68, 84, 106)
        .ThemeColorScheme(msoThemeLight1) = RGB(255, 255, 255)
        .ThemeColorScheme(msoThemeLight2) = RGB(231, 230, 230)
    End With
    Dim i As Integer
    i = 0
    For i = 0 To CInt(GetFileContent("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.txt"))
        ChangeTheme
    Next i
    PreparePic "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Background.png"
    Slide1.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide4.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide14.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide16.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide17.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    ActivePresentation.SlideShowWindow.View.GotoSlide (3)
    Exit Sub
Crash:
    SetFileContent "/System/Settings.cnf", "Nobody", "Autologin"
    Slide22.Shapes("Details").TextFrame.TextRange.Text = "Error details: USER_LOGIN_ERROR"
    ActivePresentation.SlideShowWindow.View.GotoSlide 22
End Sub

Sub Logout()
    ' Clear temporary files
    DeleteDir "/Temp/"
    NewFolder "/Temp"
    ResetWindows
    Slide1.Shapes("Username").TextFrame.TextRange.Text = "Nobody"
    ActivePresentation.SlideShowWindow.View.GotoSlide (12)
End Sub

Sub GuestLogin()
    Slide13.PasswordField.Text = ""
    Slide13.UsernameFIeld.Text = ""
    Slide1.AxTextBox.Visible = False
    Slide1.AxComboBox.Visible = False
    Slide13.AxTextBox.Visible = False
    With Slide1.Master.Theme
        .ThemeColorScheme(msoThemeAccent1) = RGB(68, 114, 196)
        .ThemeColorScheme(msoThemeAccent2) = RGB(237, 125, 49)
        .ThemeColorScheme(msoThemeAccent3) = RGB(165, 165, 165)
        .ThemeColorScheme(msoThemeAccent4) = RGB(255, 192, 0)
        .ThemeColorScheme(msoThemeAccent5) = RGB(91, 155, 213)
        .ThemeColorScheme(msoThemeAccent6) = RGB(112, 173, 71)
        .ThemeColorScheme(msoThemeDark1) = RGB(0, 0, 0)
        .ThemeColorScheme(msoThemeDark2) = RGB(68, 84, 106)
        .ThemeColorScheme(msoThemeLight1) = RGB(255, 255, 255)
        .ThemeColorScheme(msoThemeLight2) = RGB(231, 230, 230)
    End With
    Slide15.Export Environ("TEMP") & "\Userpic.PNG", "PNG"
    Slide1.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide4.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide11.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide14.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide16.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide17.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide1.Shapes("Username").TextFrame.TextRange.Text = "Guest"
    ActivePresentation.SlideShowWindow.View.GotoSlide (3)
End Sub

Sub Login()
    On Error GoTo Crash
    DeleteDir "/Temp/"
    NewFolder "/Temp"
    If Slide13.UsernameFIeld.Text = Slide1.Shapes("Username").TextFrame.TextRange.Text Then
        Slide13.PasswordField.Text = ""
        Slide13.UsernameFIeld.Text = ""
        AppMessage "You are already logged in. Please log out first.", "Login", "Exclamation", False
        Exit Sub
    End If
    Dim Shp As Shape
    Dim Username As String
    Username = ""
    Slide1.AxTextBox.Visible = False
    Slide1.AxComboBox.Visible = False
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then Slide13.AxTextBox.Visible = False
    If FileStreamsExist("/Users/" & Slide13.UsernameFIeld.Text & "/") Then
        Username = Slide13.UsernameFIeld.Text
    End If
    
    If Username = "" Then
        Slide13.PasswordField.Text = ""
        Slide13.UsernameFIeld.Text = ""
        AppMessage "Wrong username", "Login", "Exclamation", False
        Exit Sub
    End If
    Slide1.Shapes("Username").TextFrame.TextRange.Text = Slide13.UsernameFIeld.Text
    CorrectPassword = GetFileContent("/Users/" & Username & "/Password.txt")
    Slide1.Shapes("Username").TextFrame.TextRange.Text = "Nobody"
    EnteredPassword = Slide13.PasswordField.Text
    If CorrectPassword <> EnteredPassword Then
        Slide13.PasswordField.Text = ""
        Slide13.UsernameFIeld.Text = ""
        AppMessage "Wrong password", "Login", "Exclamation", False
        Exit Sub
    End If
    Slide1.Shapes("Username").TextFrame.TextRange.Text = Slide13.UsernameFIeld.Text
    InitUserspace
Done:
    Exit Sub
Crash:
    Slide22.Shapes("Details").TextFrame.TextRange.Text = "Error details: USER_LOGIN_ERROR"
    ActivePresentation.SlideShowWindow.View.GotoSlide 22
End Sub

Sub PicClick(Shp As Shape)
    UserID = GetAppID(Shp)
    User = Slide13.Shapes("UserName:" & UserID).TextFrame.TextRange.Text
    Slide1.Shapes("Username").TextFrame.TextRange.Text = User
    CorrectPass = GetFileContent("/Users/" & User & "/Password.txt")
    Slide1.Shapes("Username").TextFrame.TextRange.Text = "Nobody"
    Slide13.UsernameFIeld.Text = User
    If CorrectPass <> "" And CorrectPass <> "*" Then
        SetVar "UserID", UserID
        SetVar "Macro", "FinishInteractiveLogon"
        SetVar "User", User
        AppInputBox "Enter your password", "Login screen"
    Else
        Login
    End If
End Sub

Sub PicClickParent(Shp As Shape)
    PicClick Shp.ParentGroup
End Sub

Sub FinishInteractiveLogon()
    UserID = CheckVars("%UserID%")
    EnteredPass = CheckVars("%InputValue%")
    Slide13.UsernameFIeld.Value = CheckVars("%User%")
    Slide13.PasswordField.Text = EnteredPass
    UnsetVar "UserID"
    UnsetVar "User"
    UnsetVar "InputValue"
    Login
End Sub


Sub AboutSessMan()
    AppMessage "OrangePath Session Manager" + vbNewLine + "Version 1.0 by mmaal" + vbNewLine + "Do not feed dead stars!", "You found me!", "Info", False
End Sub

Sub InitLogon()
    Slide13.Shapes("UserID").TextFrame.TextRange.Text = "1"
    Users = GetFiles("/Users/")
    UsersList = Split(Users, "/")
    For i = Slide13.Shapes.Count To 1 Step -1
        Dim Shp As Shape
        Set Shp = Slide13.Shapes(i)
        If InStr(Shp.Name, ":") Then
            Shp.Delete
        End If
    Next i
    PicX = 59
    PicY = 92
    X = 1
    For i = UBound(UsersList) To 0 Step -1
        User = Replace(UsersList(i), vbNewLine, "")
        If User <> "" Then
            Slide13.Shapes("DefaultPic").Copy
            With Slide13.Shapes.Paste
                .Name = "UserPic:" & Slide13.Shapes("UserID").TextFrame.TextRange.Text
                .Left = PicX
                .Top = PicY
                .Visible = msoTrue
            End With
            Slide13.Shapes("DefaultUname").Copy
            With Slide13.Shapes.Paste
                .Name = "UserName:" & Slide13.Shapes("UserID").TextFrame.TextRange.Text
                .TextFrame.TextRange.Text = User
                .Left = PicX - 6
                .Top = PicY + 75
                .Visible = msoTrue
            End With
            Slide13.Shapes("UserID").TextFrame.TextRange.Text = Int(Slide13.Shapes("UserID").TextFrame.TextRange.Text) + 1
            PicX = PicX + 107.25
            X = X + 1
            If X = 9 Then
                X = 1
                PicX = 59
                PicY = PicY + 107.25
            End If
        End If
    Next i
End Sub

