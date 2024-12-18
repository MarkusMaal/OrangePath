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
        Slide2.Shapes("ShpMisc6ButtonAppSettings_").TextFrame.TextRange.Text = "Enable"
    Else
        Slide2.Shapes("ShpMisc6ButtonAppSettings_").TextFrame.TextRange.Text = "Disable"
    End If
    
    If GetFileContent("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Password.txt") <> "" Then
        Slide2.Shapes("ShpUser6AppSettings_").TextFrame.TextRange.Text = "You have a password"
        Slide2.Shapes("ShpUser5ButtonAppSettings_").TextFrame.TextRange.Text = "Change"
    Else
        Slide2.Shapes("ShpUser6AppSettings_").TextFrame.TextRange.Text = "No password set"
        Slide2.Shapes("ShpUser5ButtonAppSettings_").TextFrame.TextRange.Text = "Set"
    End If
    If GetFileContent("/System/Settings.cnf", "Autologin") <> "Nobody" Then
        Slide2.Shapes("ShpUser8AppSettings_").TextFrame.TextRange.Text = "Autologin active"
        Slide2.Shapes("ShpUser7ButtonAppSettings_").TextFrame.TextRange.Text = "Disable"
    Else
        Slide2.Shapes("ShpUser8AppSettings_").TextFrame.TextRange.Text = "No autologin"
        Slide2.Shapes("ShpUser7ButtonAppSettings_").TextFrame.TextRange.Text = "Enable"
    End If
    Interval = CInt(GetFileContent("/System/Settings.cnf", "AutosaveInterval"))
    If Interval = 0 Then
        Slide2.Shapes("ShpPersonalise6ButtonAppSettings_").TextFrame.TextRange.Text = "Enable"
    Else
        Slide2.Shapes("ShpPersonalise6ButtonAppSettings_").TextFrame.TextRange.Text = "Disable"
    End If
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide1.Shapes("WindowTitleAppSettings:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "Settings"
    AppSettingsSwitchCat Slide1.Shapes("CatPersonaliseAppSettings:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text)
    InitAbout Slide1.Shapes("AppID").TextFrame.TextRange.Text
    UpdateTime
End Sub

Sub AppSettingsGetPicture(AppID As String)
    Dim PicPath As String
    PicPath = "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/UserPic.png"
    If FileStreamsExist(PicPath) Then
        Slide1.Shapes("ShpUser11AppSettings:" & AppID).Visible = msoFalse
        Slide1.Shapes("ShpUser10AppSettings:" & AppID).Visible = msoFalse
        PreparePic "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/UserPic.png"
        Slide1.Shapes("ShpUser12AppSettings:" & AppID).Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
        Slide1.Shapes("ShpUser12AppSettings:" & AppID).LockAspectRatio = msoTrue
        Slide1.Shapes("ShpUser12AppSettings:" & AppID).Fill.Transparency = 0
    End If
    Slide1.Shapes("ShpUser9AppSettings:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("Username").TextFrame.TextRange.Text
End Sub

Sub TestPicture()
    AppSettingsGetPicture "22"
End Sub

Sub TestTheme()
    ChangeTheme
End Sub

Sub InitAbout(ByVal AppID As String)
    Dim fS, F
    Dim mB As Double
    Set fS = CreateObject("Scripting.FileSystemObject")
    Set F = fS.GetFile(Application.ActivePresentation.FullName)
    mB = F.Size / 1024 / 1024
    Dim AppCount As Integer
    Dim AppLaunch As Integer
    Dim Suffix As String
    AppCount = 0
    AppLaunch = 0
    
    Dim BuildNo As Integer
    BuildNo = GetBuildNo
    
    If BuildNo < 1000 Then
        Suffix = "Beta 0." & Left(BuildNo, 1)
    Else
        Suffix = Left(BuildNo, 1) & "." & Right(Left(BuildNo, 2), 1)
    End If
    
    Dim I As Integer
    For I = Slide2.Shapes.Count To 1 Step -1
        If InStr(Slide2.Shapes(I).Name, "App") Then
            AppCount = AppCount + 1
        End If
        If ShapeExists(Slide25, Slide2.Shapes(I).Name & ":Icon") Then
            AppLaunch = AppLaunch + 1
        End If
    Next I
    Slide1.Shapes("ShpAbout2AppSettings:" & AppID).TextFrame.TextRange.Text = "Sunlight OS " & Suffix
    Slide1.Shapes("ShpAbout3AppSettings:" & AppID).TextFrame.TextRange.Text = "Build: " & BuildNo & vbNewLine & "Space usage: " & Round(mB, 2) & "MB" & vbNewLine & "ShapeFS inodes: " & CountInodes
    Slide1.Shapes("ShpAbout4AppSettings:" & AppID).TextFrame.TextRange.Text = "User accounts: " & CountUsers & vbNewLine & "Updater: " & Replace(Slide12.Shapes("Version").TextFrame.TextRange.Text, "DFU ", "") & vbNewLine & "Installed apps: " & AppCount & " (" & AppLaunch & " links)"
End Sub

Sub RotateMe(Shp As Shape)
    Shp.ThreeD.RotationZ = Shp.ThreeD.RotationZ + 22.5
    Shp.ThreeD.RotationX = Shp.ThreeD.RotationX + 22.5
    Shp.ThreeD.RotationY = Shp.ThreeD.RotationY + 22.5
End Sub

Function CountUsers() As Integer
    Dim Users As String
    Users = GetFiles("/Users/")
    Dim Usrs As Variant
    Usrs = Split(Users, vbNewLine)
    Dim UsrCount As Integer
    UsrCount = 0
    For Each Usr In Usrs
        If Right(Usr, 1) = "/" Then
            UsrCount = UsrCount + 1
        End If
    Next Usr
    CountUsers = UsrCount
End Function

Sub ChangeTheme(Optional Shp As Shape)
    Dim CurrentIndex As String
    CurrentIndex = GetFileContent("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.txt")
    Dim Theme() As Variant
    Dim ThemeBlue() As Variant
    Dim ThemeMagenta() As Variant
    Dim ThemeOrange() As Variant
    Dim ThemeAqua() As Variant
    Dim ThemeRed() As Variant
    Dim ThemeBlack() As Variant
    Dim ThemeYellow() As Variant
    Dim ThemeLime() As Variant
    Dim ThemeGreen() As Variant
    Dim ThemeGray() As Variant
    ThemeBlue = Array(RGB(52, 148, 186), RGB(88, 182, 192), RGB(117, 189, 167), RGB(122, 140, 142), RGB(132, 172, 182), RGB(38, 131, 198), RGB(0, 0, 0), RGB(55, 53, 69), RGB(255, 255, 255), RGB(206, 219, 230))
    ThemeMagenta = Array(RGB(146, 39, 143), RGB(155, 87, 211), RGB(117, 93, 217), RGB(102, 94, 184), RGB(102, 94, 184), RGB(117, 93, 217), RGB(0, 0, 0), RGB(99, 46, 98), RGB(255, 255, 255), RGB(234, 229, 235))
    ThemeOrange = Array(RGB(240, 127, 9), RGB(159, 41, 54), RGB(78, 165, 216), RGB(78, 133, 66), RGB(240, 127, 9), RGB(193, 152, 89), RGB(0, 0, 0), RGB(50, 50, 50), RGB(255, 255, 255), RGB(227, 222, 209))
    ThemeAqua = Array(RGB(68, 114, 196), RGB(237, 125, 49), RGB(165, 165, 165), RGB(255, 192, 0), RGB(91, 155, 213), RGB(112, 173, 71), RGB(0, 0, 0), RGB(68, 84, 106), RGB(255, 255, 255), RGB(231, 230, 230))
    ThemeRed = Array(RGB(165, 48, 15), RGB(213, 88, 22), RGB(242, 213, 167), RGB(177, 156, 125), RGB(144, 66, 66), RGB(178, 125, 73), RGB(0, 0, 0), RGB(50, 50, 50), RGB(255, 255, 255), RGB(232, 186, 118))
    ThemeBlack = Array(RGB(40, 40, 40), RGB(55, 55, 55), RGB(138, 138, 138), RGB(63, 63, 63), RGB(38, 38, 38), RGB(12, 12, 12), RGB(55, 55, 55), RGB(22, 22, 22), RGB(255, 255, 255), RGB(248, 248, 248))
    ThemeYellow = Array(RGB(240, 162, 46), RGB(165, 100, 78), RGB(181, 139, 128), RGB(195, 152, 109), RGB(161, 149, 116), RGB(193, 117, 41), RGB(0, 0, 0), RGB(78, 59, 48), RGB(255, 255, 255), RGB(251, 238, 201))
    ThemeLime = Array(RGB(153, 203, 56), RGB(99, 165, 55), RGB(55, 167, 111), RGB(68, 193, 163), RGB(78, 179, 207), RGB(81, 195, 249), RGB(0, 0, 0), RGB(69, 95, 81), RGB(255, 255, 255), RGB(226, 223, 204))
    ThemeGreen = Array(RGB(84, 158, 57), RGB(138, 184, 51), RGB(122, 207, 59), RGB(2, 150, 118), RGB(74, 181, 196), RGB(9, 137, 177), RGB(0, 0, 0), RGB(69, 95, 81), RGB(255, 255, 255), RGB(227, 222, 209))
    ThemeGray = Array(RGB(153, 153, 153), RGB(121, 121, 121), RGB(150, 150, 150), RGB(128, 128, 128), RGB(95, 95, 95), RGB(77, 77, 77), RGB(0, 0, 0), RGB(27, 27, 27), RGB(255, 255, 255), RGB(248, 248, 248))
    Dim NextIndex As Integer
    NextIndex = CInt(CurrentIndex) + 1
    If CheckVars("%NegateTheme%") = "True" Then
        UnsetVar "NegateTheme"
        NextIndex = NextIndex - 1
    End If
    Select Case NextIndex
        Case 0, 10
            Theme = ThemeBlue
            NextIndex = 0
        Case 1
            Theme = ThemeMagenta
        Case 2
            Theme = ThemeOrange
        Case 3
            Theme = ThemeAqua
        Case 4
            Theme = ThemeRed
        Case 5
            Theme = ThemeBlack
        Case 6
            Theme = ThemeYellow
        Case 7
            Theme = ThemeLime
        Case 8
            Theme = ThemeGreen
        Case 9
            Theme = ThemeGray
    End Select
    If Slide1.Shapes("Username").TextFrame.TextRange.Text <> "Guest" Then
        SetFileContent "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.txt", CStr(NextIndex)
    Else
        Exit Sub
    End If
    With Slide1.Master.Theme
        .ThemeColorScheme(msoThemeAccent1) = Theme(0)
        .ThemeColorScheme(msoThemeAccent2) = Theme(1)
        .ThemeColorScheme(msoThemeAccent3) = Theme(2)
        .ThemeColorScheme(msoThemeAccent4) = Theme(3)
        .ThemeColorScheme(msoThemeAccent5) = Theme(4)
        .ThemeColorScheme(msoThemeAccent6) = Theme(5)
        .ThemeColorScheme(msoThemeDark1) = Theme(6)
        .ThemeColorScheme(msoThemeDark2) = Theme(7)
        .ThemeColorScheme(msoThemeLight1) = Theme(8)
        .ThemeColorScheme(msoThemeLight2) = Theme(9)
    End With
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

    On Error GoTo ExitScheme
    If Not IsMissing(Shp) Then
        Dim AppID As String
        AppID = GetAppID(Shp)
        AppSettingsSwitchCat Slide1.Shapes("CatPersonaliseAppSettings:" & AppID)
    End If
ExitScheme:
    Exit Sub
End Sub

Sub AddInterval(Shp As Shape)
    AppID = GetAppID(Shp)
    Interval = CInt(GetFileContent("/System/Settings.cnf", "AutosaveInterval"))
    Interval = Interval + 1
    Success = SaveSysConfig("AutosaveInterval", CStr(Interval))
    Slide1.Shapes("ShpPersonalise4AppSettings:" & AppID).TextFrame.TextRange.Text = "Autosave interval: " & Interval & " mins"
    Slide2.Shapes("ShpPersonalise4AppSettings_").TextFrame.TextRange.Text = "Autosave interval: " & Interval & " mins"
End Sub

Sub SubInterval(Shp As Shape)
    AppID = GetAppID(Shp)
    Interval = CInt(GetFileContent("/System/Settings.cnf", "AutosaveInterval"))
    If Interval > 0 Then
        Interval = 0
    Else
        Interval = 1
    End If
    Success = SaveSysConfig("AutosaveInterval", CStr(Interval))
    If Interval = 0 Then
        Slide1.Shapes("ShpPersonalise6ButtonAppSettings:" & AppID).TextFrame.TextRange.Text = "Enable"
    Else
        Slide1.Shapes("ShpPersonalise6ButtonAppSettings:" & AppID).TextFrame.TextRange.Text = "Disable"
    End If
End Sub

Sub ActuallyChangeBg()
    If Slide1.Shapes("Username").TextFrame.TextRange.Text <> "Guest" Then
        If Left(CheckVars("%InputValue%"), 3) = "C:\" Then
            SetFilePic "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Background.png", CheckVars("%InputValue%")
        Else
            DeleteFile "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Background.png"
            CopyFile CheckVars("%InputValue%"), "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Background.png", True
            PreparePic CheckVars("%InputValue%")
            SetVar "InputValue", Environ("TEMP") & "\UserPic.PNG"
        End If
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
    On Error Resume Next
    Slide1.Shapes("BackgroundImg").Delete
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
End Sub

Sub ActuallyChangeVs()
    If FileStreamsExist("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm") Then
        DeleteFile "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm"
    End If
    CopyFile CheckVars("%InputValue%"), "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm", True
        
    UnsetVar "Macro"
    UnsetVar "LaunchDir"
    UnsetVar "InputValue"
    
    CopyFillFormat GetFileRef("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm").GroupItems("WindowFrame"), Slide1.Shapes("Taskbar")
    
    AppMessage "Theme applied", "Settings", "Info", True
End Sub

Sub ActuallyChangeUserPic()
    If Slide1.Shapes("Username").TextFrame.TextRange.Text <> "Guest" Then
        If Left(CheckVars("%InputValue%"), 3) = "C:\" Then
            SetFilePic "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/UserPic.png", CheckVars("%InputValue%")
        Else
            DeleteFile "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/UserPic.png"
            CopyFile CheckVars("%InputValue%"), "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/UserPic.png", True
            PreparePic CheckVars("%InputValue%")
            SetVar "InputValue", Environ("TEMP") & "\UserPic.PNG"
        End If
    End If
    
    Dim AppID As String
    AppID = CheckVars("%AppID%")
    
    Slide1.Shapes("ShpUser12AppSettings:" & AppID).Fill.UserPicture CheckVars("%InputValue%")
    Slide1.Shapes("ShpUser11AppSettings:" & AppID).Visible = msoFalse
    Slide1.Shapes("ShpUser10AppSettings:" & AppID).Visible = msoFalse
    
    UnsetVar "AppID"
    UnsetVar "Macro"
    UnsetVar "LaunchDir"
    UnsetVar "InputValue"
End Sub


Sub AppSettingsFinalizePrepareUpdate()
    Slide12.Shapes("FirmwareSource").TextFrame.TextRange.Text = CheckVars("%InputValue%")
    UnsetVar "Macro"
    UnsetVar "LaunchDir"
    Dim Presentation2 As Presentation
    Set Presentation2 = Presentations.Open(Filename:=Slide12.Shapes("FirmwareSource").TextFrame.TextRange.Text, ReadOnly:=msoTrue, WithWindow:=msoFalse)
    On Error GoTo FailImport
    Presentation2.Application.Run "CheckParent"
    Presentation2.Close
    Exit Sub
FailImport:
    Slide12.Shapes("FirmwareSource").TextFrame.TextRange.Text = "null"
    AppMessage "Failed to import update package. Please check that it's compatible with this version of Sunlight OS.", "Import update package", "Error", True
    Presentation2.Close
End Sub

Sub AppSettingsPrepareUpdate()
    SetVar "Macro", "AppSettingsFinalizePrepareUpdate"
    SetVar "LaunchDir", "C:\"
    SetVar "NoFs", "True"
    UnsetVar "Save"
    AppModalFiles
End Sub

Sub ChangeBg()
    SetVar "Macro", "ActuallyChangeBg"
    SetVar "LaunchDir", "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text
    UnsetVar "Save"
    AppModalFiles
End Sub

Sub ChangeVs()
    SetVar "Macro", "ActuallyChangeVs"
    SetVar "LaunchDir", "/Defaults/Themes/"
    SetVar "NoFs", "True"
    UnsetVar "Save"
    AppModalFiles
End Sub


Sub ChangeUserPic(Shp As Shape)
    SetVar "Macro", "ActuallyChangeUserPic"
    SetVar "AppID", GetAppID(Shp)
    SetVar "LaunchDir", "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/"
    UnsetVar "NoFs"
    UnsetVar "Save"
    AppModalFiles
End Sub

Sub EnableDisableAutologin(Shp As Shape)
    AppID = GetAppID(Shp)
    If GetFileContent("/System/Settings.cnf", "Autologin") = "Nobody" Then
        SetFileContent "/System/Settings.cnf", Slide1.Shapes("Username").TextFrame.TextRange.Text, "Autologin"
        Slide1.Shapes("ShpUser8AppSettings:" & AppID).TextFrame.TextRange.Text = "Autologin active"
        Slide1.Shapes("ShpUser7ButtonAppSettings:" & AppID).TextFrame.TextRange.Text = "Disable"
    Else
        SetFileContent "/System/Settings.cnf", "Nobody", "Autologin"
        Slide1.Shapes("ShpUser8AppSettings:" & AppID).TextFrame.TextRange.Text = "No autologin"
        Slide1.Shapes("ShpUser7ButtonAppSettings:" & AppID).TextFrame.TextRange.Text = "Enable"
    End If
End Sub

Sub DelUser()
    Username = Slide1.Shapes("Username").TextFrame.TextRange.Text
    Dim Shp As Shape
    DeleteDir "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/"
    ResetWindows
    ActivePresentation.SlideShowWindow.View.GotoSlide (12)
End Sub

Sub EnableDisableWordwrap(Shp As Shape)
    If Not AAX Then Exit Sub
    If Slide1.AxTextBox.WordWrap = True Then
        Slide1.AxTextBox.WordWrap = False
        Slide2.Shapes("ShpKbd4ButtonAppSettings_").TextFrame.TextRange.Text = "Enable"
        Shp.TextFrame.TextRange.Text = "Enable"
    Else
        Slide1.AxTextBox.WordWrap = True
        Slide2.Shapes("ShpKbd4ButtonAppSettings_").TextFrame.TextRange.Text = "Disable"
        Shp.TextFrame.TextRange.Text = "Disable"
    End If
End Sub

Sub ShowHideDebug(Shp As Shape)
    AppID = GetAppID(Shp)
    If Slide1.Shapes("MoveEvent").Visible = msoTrue Then
        Slide1.Shapes("ShpMisc6ButtonAppSettings:" + AppID).TextFrame.TextRange.Text = "Enable"
        Slide1.Shapes("AppID").Visible = msoFalse
        Slide1.Shapes("MoveEvent").Visible = msoFalse
        Slide1.Shapes("ResizeEvent").Visible = msoFalse
        Slide1.Shapes("AppCreatingEvent").Visible = msoFalse
        Slide1.Shapes("AutosaveTime").Visible = msoFalse
        Slide1.Shapes("Username").Visible = msoFalse
        Slide1.Shapes("BuildInfo").Visible = msoFalse
        Slide3.Shapes("BuildInfo").Visible = msoFalse
        Slide5.Shapes("BuildInfo").Visible = msoFalse
        Slide7.Shapes("BuildInfo").Visible = msoFalse
    Else
        Slide1.Shapes("ShpMisc6ButtonAppSettings:" + AppID).TextFrame.TextRange.Text = "Disable"
        Slide1.Shapes("AppID").Visible = msoTrue
        Slide1.Shapes("MoveEvent").Visible = msoTrue
        Slide1.Shapes("ResizeEvent").Visible = msoTrue
        Slide1.Shapes("AppCreatingEvent").Visible = msoTrue
        Slide1.Shapes("AutosaveTime").Visible = msoTrue
        Slide1.Shapes("Username").Visible = msoTrue
        Slide1.Shapes("BuildInfo").Visible = msoTrue
        Slide3.Shapes("BuildInfo").Visible = msoTrue
        Slide5.Shapes("BuildInfo").Visible = msoTrue
        Slide7.Shapes("BuildInfo").Visible = msoTrue
    End If
End Sub

' Run this routine after restoring the window
Sub AppSettingsRestore(AppID As String)
    AppSettingsSwitchCat Slide1.Shapes("CatPersonaliseAppSettings:" & AppID)
End Sub

Sub AppSettingsSwitchCat(Shp As Shape)
    ' Toggle settings categories
    Dim AppID As String
    Dim SwitchTo As String
    AppID = GetAppID(Shp)
    Slide1.Shapes("OverlayAppSettings:" & AppID).Visible = msoFalse
    SwitchTo = Replace(Replace(Shp.Name, "Cat", ""), "AppSettings:" & AppID, "")
    
    Dim Cats As Variant
    Cats = Array("Personalise", "User", "Kbd", "Misc", "About")
    
    For Each Cat In Cats
        If Cat = SwitchTo Then
            Slide1.Shapes("Cat" & Cat & "AppSettings:" & AppID).Fill.ForeColor.RGB = Slide1.Shapes("DummySelectedAppSettings:" & AppID).Fill.ForeColor.RGB
    
            Dim Shp2 As Shape
            For Each Shp2 In Slide1.Shapes("RegularApp:" & AppID).GroupItems
                If InStr(1, Shp2.Name, "Shp" & Cat) = 1 Then
                    Shp2.Visible = msoTrue
                ElseIf InStr(1, Shp2.Name, "Shp") = 1 Then
                    Shp2.Visible = msoFalse
                End If
            Next Shp2
        Else
            Slide1.Shapes("Cat" & Cat & "AppSettings:" & AppID).Fill.ForeColor.RGB = Slide1.Shapes("DummyUnselectedAppSettings:" & AppID).Fill.ForeColor.RGB
        End If
    Next Cat
    AppSettingsGetPicture AppID
End Sub

Sub AppSettingsEnableDisableLoadIndicators(Shp As Shape)
    Dim Val As String
    Val = Slide2.Shapes("ShpMisc8ButtonAppSettings_").TextFrame.TextRange.Text
    If GetSysConfig("Loaders") <> "False" Then
        Slide2.Shapes("ShpMisc8ButtonAppSettings_").TextFrame.TextRange.Text = "Enable load indicators"
        Shp.TextFrame.TextRange.Text = "Enable load indicators"
        SaveSysConfig "Loaders", "False"
    Else
        Slide2.Shapes("ShpMisc8ButtonAppSettings_").TextFrame.TextRange.Text = "Disable load indicators"
        Shp.TextFrame.TextRange.Text = "Disable load indicators"
        SaveSysConfig "Loaders", "True"
    End If
End Sub

Sub AppSettingsToggleActiveX(Shp As Shape)
    Dim Uname As String
    Uname = Slide1.Shapes("Username").TextFrame.TextRange.Text
    Dim NoPass As Boolean
    NoPass = False
    
    If GetFileContent("/Users/" & Uname & "/Password.txt") = "" Then
        NoPass = True
    End If
    If GetSysConfig("Autologin") <> "Nobody" Then
        NoPass = True
    End If
    If Not NoPass Then
        AppMessage "ActiveX settings cannot be toggled, because this user has a password", "ActiveX toggle", "Error", True
        Exit Sub
    End If
    ' Unelevate
    If GetSysConfig("NoActiveX") <> "True" Then
        Shp.TextFrame.TextRange.Text = "Enable"
        Slide2.Shapes("ShpKbd5ButtonAppSettings_").TextFrame.TextRange.Text = "Enable"
        SaveSysConfig "NoActiveX", "True"
    Else
        Shp.TextFrame.TextRange.Text = "Disable"
        Slide2.Shapes("ShpKbd5ButtonAppSettings_").TextFrame.TextRange.Text = "Disable"
        SaveSysConfig "NoActiveX", "False"
    End If
End Sub

Sub AppSettingsDevCat()
    ' Toggle settings categories
    Dim Cat As String
    Cat = "Misc"
    
    Dim Cats As Variant
    Cats = Array("Personalise", "User", "Kbd", "Misc", "About")
    
    Dim Shp2 As Shape
    For Each Shp2 In Slide24.Shapes
        If InStr(1, Shp2.Name, "Shp" & Cat) = 1 Then
            Shp2.Visible = msoTrue
        ElseIf InStr(1, Shp2.Name, "Shp") = 1 Then
            Shp2.Visible = msoFalse
        End If
    Next Shp2
    Slide24.Shapes("OverlayAppSettings_").Visible = msoFalse
End Sub

Sub AppSettingsDevReset()
    Dim Shp2 As Shape
    For Each Shp2 In Slide24.Shapes
        If InStr(1, Shp2.Name, "Shp") = 1 Then
            Shp2.Visible = msoTrue
        End If
    Next Shp2
    Slide24.Shapes("OverlayAppSettings_").Visible = msoTrue
End Sub