' OrangePath Session Manager
Sub InitUserspace()
    'On Error GoTo Crash
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
    SetVar "NegateTheme", "True"
    ChangeTheme
    PreparePic "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Background.png"
    Slide1.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide4.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide14.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide16.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    Slide17.Background.Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
    
    Slide4.Shapes("WelcomeText").TextFrame.TextRange.Text = "Welcome, " & Slide1.Shapes("Username").TextFrame.TextRange.Text & "!"
    Dim Shp2 As Shape
    For Each Shp2 In Slide13.Shapes
        If InStr(1, Shp2.Name, "UserName:") = 1 Then
            If Shp2.TextFrame.TextRange.Text = Slide1.Shapes("Username").TextFrame.TextRange.Text Then
                Slide13.Shapes(Replace(Shp2.Name, "UserName", "UserPic")).Copy
            End If
        End If
    Next Shp2
    Dim x As Integer
    Dim y As Integer
    x = Slide4.Shapes("Pic").Left
    y = Slide4.Shapes("Pic").Top
    Slide4.Shapes("Pic").Delete
    With Slide4.Shapes.Paste
        .Left = x
        .Top = y
        .Name = "Pic"
        Dim GI As Shape
        For Each GI In .GroupItems
            GI.ActionSettings(ppMouseClick).Action = ppActionNone
        Next GI
    End With
    Slide4.Shapes("BackgroundOverlay").ZOrder msoBringToFront
    If ShapeExists(Slide1, "BackgroundImg") Then
        Slide1.Shapes("BackgroundImg").Delete
    End If
    GetFileRef("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Background.png").Copy
    With Slide1.Shapes.Paste
        .Name = "BackgroundImg"
        .Left = 0
        .Top = 0
        .Width = ActivePresentation.PageSetup.SlideWidth
        .Height = ActivePresentation.PageSetup.SlideHeight
        .ZOrder msoSendToBack
        .Visible = msoTrue
    End With
    HideCursor
    ActivePresentation.SlideShowWindow.View.GotoSlide (3)
    If Not FileStreamsExist("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop/") Then
        NewFolder "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop"
    End If
    ShowDesktop
    Slide1.Shapes("WaitPlease").Visible = msoFalse
    If GetSysConfig("Autorun") <> "*" Then
        SetVar "Autoran", "False"
    End If
    If FileStreamsExist("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm") Then
        CopyFillFormat GetFileRef("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm").GroupItems("WindowFrame"), Slide1.Shapes("Taskbar")
    Else
        CopyFillFormat GetFileRef("/Defaults/Themes/Default.thm").GroupItems("WindowFrame"), Slide1.Shapes("Taskbar")
    End If
    Exit Sub
Crash:
    SetFileContent "/System/Settings.cnf", "Nobody", "Autologin"
    OSCrash "USER_LOGIN_ERROR", Err
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
    If AAX Then
        Slide1.AxTextBox.Visible = False
        Slide1.AxComboBox.Visible = False
        Slide13.AxTextBox.Visible = False
    End If
    
    Slide4.Shapes("WelcomeText").TextFrame.TextRange.Text = "Welcome, Guest!"
    Slide13.Shapes("DefaultPic").Copy
    Dim x As Integer
    Dim y As Integer
    x = Slide4.Shapes("Pic").Left
    y = Slide4.Shapes("Pic").Top
    Slide4.Shapes("Pic").Delete
    With Slide4.Shapes.Paste
        .Left = x
        .Top = y
        .Visible = msoTrue
        .Name = "Pic"
    End With
    Slide4.Shapes("BackgroundOverlay").ZOrder msoBringToFront
    
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
    Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 1"
    Slide1.Shapes("WorkspaceCircle1").Fill.Transparency = 0
    Slide1.Shapes("WorkspaceCircle2").Fill.Transparency = 0.5
    Slide1.Shapes("WorkspaceCircle3").Fill.Transparency = 0.5
    Slide1.Shapes("WorkspaceCircle4").Fill.Transparency = 0.5
    ActivePresentation.SlideShowWindow.View.GotoSlide (3)
    If ShapeExists(Slide1, "BackgroundImg") Then
        Slide1.Shapes("BackgroundImg").Delete
        GetFileRef("/Defaults/Images/Background.png").Copy
        With Slide1.Shapes.Paste
            .Name = "BackgroundImg"
            .Left = 0
            .Top = 0
            .Width = ActivePresentation.PageSetup.SlideWidth
            .Height = ActivePresentation.PageSetup.SlideHeight
            .Visible = msoTrue
            .ZOrder msoSendToBack
        End With
    End If
    ShowDesktop
    CopyFillFormat GetFileRef("/Defaults/Themes/Default.thm").GroupItems("WindowFrame"), Slide1.Shapes("Taskbar")
End Sub

Sub Login()
    'On Error GoTo Crash
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
    If AAX Then
        Slide1.AxTextBox.Visible = False
        Slide1.AxComboBox.Visible = False
    End If
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
    OSCrash "USER_LOGIN_ERROR", Err
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

Sub TestPicClick()
    PicClick Slide13.Shapes("UserName:22")
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

Sub ModTb()
    CopyFillFormat GetFileRef("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm").GroupItems("WindowFrame"), Slide1.Shapes("Taskbar")
End Sub

Sub AboutSessMan()
    AppMessage "Sunlight Session Manager" + vbNewLine + "Version 1.0 by mmaal" + vbNewLine + "Do not feed dead stars!", "You found me!", "Info", False
End Sub

Sub InitLogon()
    Dim Shp2 As Shape
    For Each Shp2 In Slide13.Shapes
        If InStr(1, Shp2.Name, "Button") Then
            CopyFillFormat GetFileRef("/Defaults/Themes/Default.thm").GroupItems("Button"), Shp2
        End If
    Next Shp2
    Slide13.Shapes("UserID").TextFrame.TextRange.Text = "1"
    Users = GetFiles("/Users/")
    UsersList = Split(Users, "/")
    For I = Slide13.Shapes.Count To 1 Step -1
        Dim Shp As Shape
        Set Shp = Slide13.Shapes(I)
        If InStr(Shp.Name, ":") Then
            Shp.Delete
        End If
    Next I
    PicX = 59
    PicY = 92
    x = 1
    For I = UBound(UsersList) To 0 Step -1
    
        If I < 24 Then
            User = Replace(UsersList(I), vbNewLine, "")
            
            If User <> "" Then
                Slide13.Shapes("DefaultPic").Copy
                With Slide13.Shapes.Paste
                    .Name = "UserPic:" & Slide13.Shapes("UserID").TextFrame.TextRange.Text
                    .Left = PicX
                    .Top = PicY
                    .Visible = msoTrue
                    
                    If FileStreamsExist("/Users/" & User & "/UserPic.png") Then
                        PreparePic "/Users/" & User & "/UserPic.png"
                        .GroupItems("Backdrop").Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
                        .GroupItems("DummyA").Visible = msoFalse
                        .GroupItems("DummyB").Visible = msoFalse
                    End If
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
                x = x + 1
                If x = 9 Then
                    x = 1
                    PicX = 59
                    PicY = PicY + 107.25
                End If
            End If
        End If
    Next I
    If UBound(UsersList) > 24 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide (13)
        AppMessage "Too many user accounts. Some accounts are not displayed!", "Login screen", "Error", False
    End If
End Sub
