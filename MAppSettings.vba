' Settings app

Sub AppSettings(Shp As Shape)
    Shp.ParentGroup.Delete
    If Slide1.Shapes("Username").TextFrame.TextRange.Text = "Guest" Then
        AppMessage "Guests can't change system settings", "Access denied", "Error", False
        'ActivePresentation.SlideShowWindow.View.GotoSlide (4)
        Exit Sub
    End If
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Settings"
    If Slide1.Shapes("AppCreatingEvent").Visible = msoFalse Then
        Slide2.Shapes("106*20*E*Shape6AppSettings_").TextFrame.TextRange.Text = "Enable"
    Else
        Slide2.Shapes("106*20*E*Shape6AppSettings_").TextFrame.TextRange.Text = "Disable"
    End If
    
    If GetFileContent("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Password.txt") <> "" Then
        Slide2.Shapes("Shape19AppSettings_").TextFrame.TextRange.Text = "You have a password"
        Slide2.Shapes("106*20*E*Shape20AppSettings_").TextFrame.TextRange.Text = "Change"
    Else
        Slide2.Shapes("Shape19AppSettings_").TextFrame.TextRange.Text = "No password set"
        Slide2.Shapes("106*20*E*Shape20AppSettings_").TextFrame.TextRange.Text = "Set"
    End If
    If GetFileContent("/System/Settings.cnf", "Autologin") <> "Nobody" Then
        Slide2.Shapes("Shape17AppSettings_").TextFrame.TextRange.Text = "Autologin active"
        Slide2.Shapes("106*20*E*Shape18AppSettings_").TextFrame.TextRange.Text = "Disable"
    Else
        Slide2.Shapes("Shape17AppSettings_").TextFrame.TextRange.Text = "No autologin"
        Slide2.Shapes("106*20*E*Shape18AppSettings_").TextFrame.TextRange.Text = "Enable"
    End If
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
End Sub
Sub ChangeTheme()
    If Slide1.Master.Theme.ThemeColorScheme(msoThemeAccent1) = RGB(68, 114, 196) Then
        With Slide1.Master.Theme
            .ThemeColorScheme(msoThemeAccent1) = RGB(52, 148, 186)
            .ThemeColorScheme(msoThemeAccent2) = RGB(88, 182, 192)
            .ThemeColorScheme(msoThemeAccent3) = RGB(117, 189, 167)
            .ThemeColorScheme(msoThemeAccent4) = RGB(122, 140, 142)
            .ThemeColorScheme(msoThemeAccent5) = RGB(132, 172, 182)
            .ThemeColorScheme(msoThemeAccent6) = RGB(38, 131, 198)
            .ThemeColorScheme(msoThemeDark1) = RGB(0, 0, 0)
            .ThemeColorScheme(msoThemeDark2) = RGB(55, 53, 69)
            .ThemeColorScheme(msoThemeLight1) = RGB(255, 255, 255)
            .ThemeColorScheme(msoThemeLight2) = RGB(206, 219, 230)
        End With
        If Slide1.Shapes("Username").TextFrame.TextRange.Text <> "Guest" Then
            SetFileContent "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.txt", "0"
        End If
    ElseIf Slide1.Master.Theme.ThemeColorScheme(msoThemeAccent1) = RGB(52, 148, 186) Then
        With Slide1.Master.Theme
            .ThemeColorScheme(msoThemeAccent1) = RGB(146, 39, 143)
            .ThemeColorScheme(msoThemeAccent2) = RGB(155, 87, 211)
            .ThemeColorScheme(msoThemeAccent3) = RGB(117, 93, 217)
            .ThemeColorScheme(msoThemeAccent4) = RGB(102, 94, 184)
            .ThemeColorScheme(msoThemeAccent5) = RGB(102, 94, 184)
            .ThemeColorScheme(msoThemeAccent6) = RGB(117, 93, 217)
            .ThemeColorScheme(msoThemeDark1) = RGB(0, 0, 0)
            .ThemeColorScheme(msoThemeDark2) = RGB(99, 46, 98)
            .ThemeColorScheme(msoThemeLight1) = RGB(255, 255, 255)
            .ThemeColorScheme(msoThemeLight2) = RGB(234, 229, 235)
        End With
        If Slide1.Shapes("Username").TextFrame.TextRange.Text <> "Guest" Then SetFileContent "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.txt", "1"
    ElseIf Slide1.Master.Theme.ThemeColorScheme(msoThemeAccent1) = RGB(146, 39, 143) Then
        With Slide1.Master.Theme
            .ThemeColorScheme(msoThemeAccent1) = RGB(240, 127, 9)
            .ThemeColorScheme(msoThemeAccent2) = RGB(159, 41, 54)
            .ThemeColorScheme(msoThemeAccent3) = RGB(78, 165, 216)
            .ThemeColorScheme(msoThemeAccent4) = RGB(78, 133, 66)
            .ThemeColorScheme(msoThemeAccent5) = RGB(240, 127, 9)
            .ThemeColorScheme(msoThemeAccent6) = RGB(193, 152, 89)
            .ThemeColorScheme(msoThemeDark1) = RGB(0, 0, 0)
            .ThemeColorScheme(msoThemeDark2) = RGB(50, 50, 50)
            .ThemeColorScheme(msoThemeLight1) = RGB(255, 255, 255)
            .ThemeColorScheme(msoThemeLight2) = RGB(227, 222, 209)
        End With
        If Slide1.Shapes("Username").TextFrame.TextRange.Text <> "Guest" Then SetFileContent "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.txt", "2"
    Else
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
        If Slide1.Shapes("Username").TextFrame.TextRange.Text <> "Guest" Then SetFileContent "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.txt", "3"
    End If
    With Slide4.Master.Theme
        .ThemeColorScheme(msoThemeAccent1) = Slide1.Master.Theme.ThemeColorScheme(msoThemeAccent1)
        .ThemeColorScheme(msoThemeAccent2) = Slide1.Master.Theme.ThemeColorScheme(msoThemeAccent2)
        .ThemeColorScheme(msoThemeAccent3) = Slide1.Master.Theme.ThemeColorScheme(msoThemeAccent3)
        .ThemeColorScheme(msoThemeAccent4) = Slide1.Master.Theme.ThemeColorScheme(msoThemeAccent4)
        .ThemeColorScheme(msoThemeAccent5) = Slide1.Master.Theme.ThemeColorScheme(msoThemeAccent5)
        .ThemeColorScheme(msoThemeAccent6) = Slide1.Master.Theme.ThemeColorScheme(msoThemeAccent6)
        .ThemeColorScheme(msoThemeDark1) = Slide1.Master.Theme.ThemeColorScheme(msoThemeDark1)
        .ThemeColorScheme(msoThemeDark2) = Slide1.Master.Theme.ThemeColorScheme(msoThemeDark2)
        .ThemeColorScheme(msoThemeLight1) = Slide1.Master.Theme.ThemeColorScheme(msoThemeLight1)
        .ThemeColorScheme(msoThemeLight2) = Slide1.Master.Theme.ThemeColorScheme(msoThemeLight2)
    End With
End Sub

Sub AddInterval(Shp As Shape)
    AppID = GetAppID(Shp)
    interval = CInt(GetFileContent("/System/Settings.cnf", "AutosaveInterval"))
    interval = interval + 1
    Success = SaveSysConfig("AutosaveInterval", CStr(interval))
    Slide1.Shapes("Shape11AppSettings:" & AppID).TextFrame.TextRange.Text = "Autosave interval: " & interval & " mins"
    Slide2.Shapes("Shape11AppSettings_").TextFrame.TextRange.Text = "Autosave interval: " & interval & " mins"
End Sub

Sub SubInterval(Shp As Shape)
    AppID = GetAppID(Shp)
    interval = CInt(GetFileContent("/System/Settings.cnf", "AutosaveInterval"))
    interval = interval - 1
    If interval < 1 Then
        interval = 1
    End If
    Success = SetFileContent("/System/Settings.cnf", CStr(interval), "AutosaveInterval")
    Slide1.Shapes("Shape11AppSettings:" & AppID).TextFrame.TextRange.Text = "Autosave interval: " & interval & " mins"
    Slide2.Shapes("Shape11AppSettings_").TextFrame.TextRange.Text = "Autosave interval: " & interval & " mins"
End Sub

Sub ActuallyChangeBg()
    If Slide1.Shapes("Username").TextFrame.TextRange.Text <> "Guest" Then
        SetFilePic "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Background.png", CheckVars("%InputValue%")
    End If
    
    With Slide1.Background.Fill
    .UserPicture CheckVars("%InputValue%")
    End With
    
    With Slide4.Background.Fill
    .UserPicture CheckVars("%InputValue%")
    End With
    
    With Slide14.Background.Fill
    .UserPicture CheckVars("%InputValue%")
    End With
    
    With Slide16.Background.Fill
    .UserPicture CheckVars("%InputValue%")
    End With
    With Slide17.Background.Fill
    .UserPicture CheckVars("%InputValue%")
    End With
    UnsetVar "Macro"
    UnsetVar "LaunchDir"
    UnsetVar "InputValue"
End Sub

Sub ChangeBg()
    SetVar "Macro", "ActuallyChangeBg"
    SetVar "LaunchDir", "C:\"
    UnsetVar "Save"
    AppModalFiles
End Sub

'Sub EnableDisableAutologin()
Sub EnableDisableAutologin(Shp As Shape)
    AppID = GetAppID(Shp)
    If GetFileContent("/System/Settings.cnf", "Autologin") = "Nobody" Then
        SetFileContent "/System/Settings.cnf", Slide1.Shapes("Username").TextFrame.TextRange.Text, "Autologin"
        Slide1.Shapes("Shape17AppSettings:" & AppID).TextFrame.TextRange.Text = "Autologin active"
        Slide1.Shapes("106*20*E*Shape18AppSettings:" & AppID).TextFrame.TextRange.Text = "Disable"
    Else
        SetFileContent "/System/Settings.cnf", "Nobody", "Autologin"
        Slide1.Shapes("Shape17AppSettings:" & AppID).TextFrame.TextRange.Text = "No autologin"
        Slide1.Shapes("106*20*E*Shape18AppSettings:" & AppID).TextFrame.TextRange.Text = "Enable"
    End If
End Sub

Sub DelUser()
    Username = Slide1.Shapes("Username").TextFrame.TextRange.Text
    Dim Shp As Shape
    DeleteDir "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/"
    ResetWindows
    ActivePresentation.SlideShowWindow.View.GotoSlide (13)
End Sub


Sub ShowHideDebug(Shp As Shape)
    AppID = GetAppID(Shp)
    If Slide1.Shapes("MoveEvent").Visible = msoTrue Then
        Slide1.Shapes("106*20*E*Shape6AppSettings:" + AppID).TextFrame.TextRange.Text = "Enable"
        Slide1.Shapes("AppID").Visible = msoFalse
        Slide1.Shapes("MoveEvent").Visible = msoFalse
        Slide1.Shapes("ResizeEvent").Visible = msoFalse
        Slide1.Shapes("AppCreatingEvent").Visible = msoFalse
        Slide1.Shapes("AutosaveTime").Visible = msoFalse
        Slide1.Shapes("Username").Visible = msoFalse
    Else
        Slide1.Shapes("106*20*E*Shape6AppSettings:" + AppID).TextFrame.TextRange.Text = "Disable"
        Slide1.Shapes("AppID").Visible = msoTrue
        Slide1.Shapes("MoveEvent").Visible = msoTrue
        Slide1.Shapes("ResizeEvent").Visible = msoTrue
        Slide1.Shapes("AppCreatingEvent").Visible = msoTrue
        Slide1.Shapes("AutosaveTime").Visible = msoTrue
        Slide1.Shapes("Username").Visible = msoTrue
    End If
End Sub
