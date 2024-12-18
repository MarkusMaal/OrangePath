
Sub HideHourglass()
    Slide26.Shapes("Hourglass").Visible = msoFalse
    Slide26.Shapes("WaitText").Visible = msoFalse
End Sub

Sub ShowHourglass()
    Slide26.Shapes("Hourglass").Visible = msoTrue
    Slide26.Shapes("WaitText").Visible = msoTrue
End Sub

Sub SetupRadioUncheck(ByVal RadioGroup As String, ByVal Parent As Shape)
    Dim Shp As Shape
    For Each Shp In Parent.GroupItems
        Shp.Line.ForeColor.RGB = RGB(0, 0, 0)
    Next Shp
End Sub

Sub SetupRadioCheck(Shp As Shape)
    Dim RadioGroup As String
    Dim SplitName() As String
    SplitName = Split(Shp.Name, "/")
    RadioGroup = SplitName(1)
    SetupRadioUncheck RadioGroup, Shp.ParentGroup
    Shp.Line.ForeColor.RGB = RGB(255, 255, 255)
    SetupWarningCheck Shp.TextFrame.TextRange.Text
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = "36" Then
        PreviewSetupTheme
    End If
End Sub

Function IsRadioChecked(ByVal RadioGroup As String, ByVal Label As String, ByVal Sld As Slide)
    If Sld.Shapes(Label + "/" + RadioGroup).Line.ForeColor.RGB = RGB(255, 255, 255) Then
        IsRadioChecked = msoTrue
    Else
        IsRadioChecked = msoFalse
    End If
End Function

Sub SetupStep2()
    If Slide26.Shapes("SetupScreenWelcome").Visible = msoTrue Then
        If IsRadioChecked("1", "Migrate", Slide31) Then
            AppMessage "Not Implemented", "Setup experience", "Error", False
        ElseIf IsRadioChecked("1", "Upgrade", Slide31) Then
            AppMessage "Not Implemented", "Setup experience", "Error", False
        ElseIf IsRadioChecked("1", "FreshInstall", Slide31) Then
            Slide26.Shapes("SetupScreenWelcome").Visible = msoFalse
            Slide26.Shapes("SetupScreenPartition").Visible = msoTrue
            Slide26.Shapes("Progress1").Visible = msoTrue
        End If
    End If
End Sub


Sub SetupWarningCheck(ByVal Text As String)
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 32 Then
        If InStr(1, Text, "Upgrade") Then
            Slide31.Shapes("WarningA").Visible = msoTrue
            Slide31.Shapes("WarningB").Visible = msoTrue
        Else
            Slide31.Shapes("WarningA").Visible = msoFalse
            Slide31.Shapes("WarningB").Visible = msoFalse
        End If
        Slide31.Shapes("MigNextButton").Visible = msoFalse
        Slide31.Shapes("UpgNextButton").Visible = msoFalse
        Slide31.Shapes("FrNextButton").Visible = msoFalse
        If IsRadioChecked("1", "Migrate", Slide31) Then
            Slide31.Shapes("MigNextButton").Visible = msoTrue
        ElseIf IsRadioChecked("1", "FreshInstall", Slide31) Then
            Slide31.Shapes("FrNextButton").Visible = msoTrue
        ElseIf IsRadioChecked("1", "Upgrade", Slide31) Then
            Slide31.Shapes("UpgNextButton").Visible = msoTrue
        End If
    End If
End Sub

Sub SetupMigrate()
    ActivePresentation.SlideShowWindow.View.GotoSlide 34
End Sub

Sub SetupMigrateLightOS()
    'On Error GoTo Crash
    Dim dlgOpen As FileDialog
    Dim strResult As String
    
    Set dlgOpen = Application.FileDialog(Type:=msoFileDialogFilePicker)
    
    With dlgOpen
       .Filters.Clear
      .Filters.Add "Macro-enabled PowerPoint slideshows", "*.ppsm", 1
      .AllowMultiSelect = False
        If .Show = True Then
            strResult = .SelectedItems(1)
            If strResult = "" Then
                Exit Sub
            Else
                Presentations.Open strResult, msoFalse, msoFalse, msoFalse
            End If
        End If
    End With
    
    
    Dim osld As Slide
    Set osld = Presentations(2).Slides(42)
    Dim oshp As Shape
    Dim strReport As String
    Dim sizeX As Integer
    Dim sizeY As Integer
    sizeX = 490
    sizeY = 304
    Dim Filename As String
    
    ' Migrate Words documents
    For Each oshp In osld.Shapes
        If oshp.Type = msoOLEControlObject Then
            If oshp.OLEFormat.ProgID = "Forms.TextBox.1" Then
             If oshp.OLEFormat.Object.Name = "TextBox1" Then
                Filename = "Document.wdoc"
             ElseIf oshp.OLEFormat.Object.Name = "TextBox2" Then
                Filename = "Document2.wdoc"
             ElseIf oshp.OLEFormat.Object.Name = "TextBox3" Then
                Filename = "Document3.wdoc"
             ElseIf oshp.OLEFormat.Object.Name = "TextBox4" Then
                Filename = "Document4.wdoc"
             ElseIf oshp.OLEFormat.Object.Name = "TextBox5" Then
                Filename = "Document5.wdoc"
             End If
             With Slide9.Shapes
                With .AddTextbox(msoTextOrientationHorizontal, 0, 0, sizeX, sizeY)
                    .Name = "AXTextBox2AppWords_A"
                    .TextFrame.TextRange.Font.Name = oshp.OLEFormat.Object.Font.Name
                    .TextFrame.TextRange.Font.Size = oshp.OLEFormat.Object.Font.Size
                    .TextFrame.TextRange.Font.Bold = oshp.OLEFormat.Object.Font.Bold
                    .TextFrame.TextRange.Font.Italic = oshp.OLEFormat.Object.Font.Italic
                    .TextFrame.TextRange.Font.Underline = oshp.OLEFormat.Object.Font.Underline
                    .TextFrame2.TextRange.Font.Strikethrough = oshp.OLEFormat.Object.Font.Strikethrough
                    .TextFrame.TextRange.Font.Color.RGB = oshp.OLEFormat.Object.ForeColor
                    .Fill.ForeColor.RGB = oshp.OLEFormat.Object.BackColor
                    .TextFrame.TextRange.Text = oshp.OLEFormat.Object.Value
                End With
                
                .AddTextbox(msoTextOrientationHorizontal, 0, 0, sizeX, sizeY).Name = "SizeKeyA"
             End With
             With Slide9.Shapes.Range(Array("AXTextBox2AppWords_A", "SizeKeyA")).Group
                 .Name = "/Temp/MigrateData/Words/" & Filename
             End With
             With GetFileRef("/Temp/MigrateData/Words/" & Filename)
                .GroupItems("SizeKeyA").Name = "SizeKey"
                .GroupItems("AXTextBox2AppWords_A").Name = "AXTextBox2AppWords_"
                .Visible = msoFalse
             End With
            End If
        End If
    Next
    
    ' Migrate autologin status
    Slide34.Shapes("AutologinCheck").GroupItems("Check").Fill.ForeColor.RGB = RGB(34, 0, 96)
    Set osld = Presentations(2).Slides(23)
    For Each oshp In osld.Shapes
        If oshp.Type = msoOLEControlObject Then
            If oshp.OLEFormat.ProgID = "Forms.CheckBox.1" Then
                If oshp.OLEFormat.Object.Value = True Then
                    SetupCheckUncheckAutologin (Slide34.Shapes("AutologinCheck"))
                End If
            End If
        End If
    Next
    
    ' Migrate user account
    Set osld = Presentations(2).Slides(10)
    For Each oshp In osld.Shapes
        If oshp.Type = msoOLEControlObject Then
            If oshp.OLEFormat.ProgID = "Forms.TextBox.1" Then
                If oshp.OLEFormat.Object.Name = "UserBox" Then
                    Slide34.TextBox1.Text = oshp.OLEFormat.Object.Value
                End If
            End If
        End If
    Next
    
    ' Migrate wallpaper
    If FileExists(Environ("TEMP") & "\Wallpaper.PNG") Then 'Check if file already exists
       ' First remove readonly attribute, if set
       SetAttr Environ("TEMP") & "\Wallpaper.PNG", vbNormal
       ' Then delete the file
       Kill Environ("TEMP") & "\Wallpaper.PNG"
    End If
    Presentations(2).Slides(6).Export Environ("TEMP") & "\Wallpaper.PNG", "PNG"
    SetFilePic "/Temp/MigrateData/Background.png", Environ("TEMP") & "\Wallpaper.PNG"
    
    ' Migrate presentator document
    SetupMigrateLightOSPresentator
    
    
    Presentations(2).Close
    If Slide34.TextBox1.Text = "" Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 35
    Else
        ActivePresentation.SlideShowWindow.View.GotoSlide 36
    End If
    Exit Sub
Crash:
    OSCrash "INCOMPATIBLE_MIGRATION_FILE", Err
End Sub

Function SetupHostFileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub SetupMigrateLightOSPresentator()
    Dim osld As Slide
    Set osld = Presentations(2).Slides(51)
    
    Dim Slide1Title As String
    Dim Slide1Subtitle As String
    Dim Slide2Title As String
    Dim Slide2Text As String
    Dim SlideBg As Long
    Dim SlideFg As Long
    Dim Fontname As String
    
    For Each oshp In osld.Shapes
        If oshp.Type = msoOLEControlObject Then
            If oshp.OLEFormat.ProgID = "Forms.TextBox.1" Then
                If oshp.OLEFormat.Object.Name = "TextBox1" Then
                    Slide1Title = oshp.OLEFormat.Object.Value
                    SlideBg = oshp.OLEFormat.Object.BackColor
                    SlideFg = oshp.OLEFormat.Object.ForeColor
                    Fontname = oshp.OLEFormat.Object.Font.Name
                ElseIf oshp.OLEFormat.Object.Name = "TextBox2" Then
                    Slide1Subtitle = oshp.OLEFormat.Object.Value
                ElseIf oshp.OLEFormat.Object.Name = "TextBox3" Then
                    Slide2Title = oshp.OLEFormat.Object.Value
                ElseIf oshp.OLEFormat.Object.Name = "TextBox4" Then
                    Slide2Text = oshp.OLEFormat.Object.Value
                End If
            End If
        End If
    Next
    Slide9.Shapes.AddShape(msoShapeRectangle, 0, 0, 351.5274, 197.4676).Name = "SizeKeyA"
    With Slide9.Shapes.AddShape(msoShapeRectangle, 0, 0, 351.5274, 197.4676)
        .Name = "Background"
        .ActionSettings(ppMouseClick).Action = ppActionRunMacro
        .ActionSettings(ppMouseClick).Run = "SelShape"
    End With
    Slide9.Shapes("Background").Fill.ForeColor.RGB = SlideBg
    With Slide9.Shapes.AddShape(msoShapeRectangle, 16.5, 56.65, 320.2, 36.35)
        .Name = "Slide1Title"
        .Fill.ForeColor.RGB = SlideBg
        .Fill.BackColor.RGB = SlideBg
        .TextFrame.TextRange.Font.Name = Fontname
        .TextFrame.TextRange.Font.Color.RGB = SlideFg
        .TextFrame.TextRange.Text = Slide1Title
        .TextFrame.TextRange.Font.Size = 24
        .Line.Visible = msoFalse
        .ActionSettings(ppMouseClick).Action = ppActionRunMacro
        .ActionSettings(ppMouseClick).Run = "SelShape"
    End With
    With Slide9.Shapes.AddShape(msoShapeRectangle, 16.5, 101, 320.2, 57.6)
        .Name = "Slide1Subtitle"
        .Fill.ForeColor.RGB = SlideBg
        .Fill.BackColor.RGB = SlideBg
        .TextFrame.TextRange.Font.Name = Fontname
        .TextFrame.TextRange.Font.Color.RGB = SlideFg
        .TextFrame.TextRange.Text = Slide1Subtitle
        .TextFrame.TextRange.Font.Size = 12
        .Line.Visible = msoFalse
        .ActionSettings(ppMouseClick).Action = ppActionRunMacro
        .ActionSettings(ppMouseClick).Run = "SelShape"
    End With
    With Slide9.Shapes.AddShape(msoShapeRectangle, 0, 0, 351.5274, 197.4676)
        .Name = "Background2"
        .ActionSettings(ppMouseClick).Action = ppActionRunMacro
        .ActionSettings(ppMouseClick).Run = "SelShape"
    End With
    
    Slide9.Shapes("Background2").Fill.ForeColor.RGB = SlideBg
    Slide9.Shapes("Background2").Line.Visible = msoFalse
    Slide9.Shapes("Background").Line.Visible = msoFalse
    Slide9.Shapes("SizeKeyA").Line.Visible = msoFalse
    With Slide9.Shapes.AddShape(msoShapeRectangle, 8.2, 8.2, 334.2, 40.5)
        .Name = "Slide2Title"
        .Fill.ForeColor.RGB = SlideBg
        .Fill.BackColor.RGB = SlideBg
        .TextFrame.TextRange.Font.Name = Fontname
        .TextFrame.TextRange.Font.Color.RGB = SlideFg
        .TextFrame.TextRange.Text = Slide2Title
        .TextFrame.TextRange.Font.Size = 18
        .Line.Visible = msoFalse
        .ActionSettings(ppMouseClick).Action = ppActionRunMacro
        .ActionSettings(ppMouseClick).Run = "SelShape"
    End With
    With Slide9.Shapes.AddShape(msoShapeRectangle, 8.2, 53.6, 334.2, 133.7)
        .Name = "Slide2Text"
        .Fill.ForeColor.RGB = SlideBg
        .Fill.BackColor.RGB = SlideBg
        .TextFrame.TextRange.Font.Name = Fontname
        .TextFrame.TextRange.Font.Color.RGB = SlideFg
        .TextFrame.TextRange.Text = Slide2Text
        .TextFrame.TextRange.Font.Size = 11
        .Line.Visible = msoFalse
        .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
        .ActionSettings(ppMouseClick).Action = ppActionRunMacro
        .ActionSettings(ppMouseClick).Run = "SelShape"
    End With
    With Slide9.Shapes.Range(Array("SizeKeyA", "Background", "Slide1Title", "Slide1Subtitle", "Background2", "Slide2Title", "Slide2Text")).Group
        .Name = "/Temp/MigrateData/Presentator/Presentation.pres"
        .Visible = msoFalse
    End With
    With Slide9.Shapes("/Temp/MigrateData/Presentator/Presentation.pres")
        .GroupItems("SizeKeyA").Name = "SizeKey"
        .GroupItems("Background").Name = "PresSld1Shape1AppPresentator_"
        .GroupItems("Slide1Title").Name = "PresSld1Shape2AppPresentator_"
        .GroupItems("Slide1Subtitle").Name = "PresSld1Shape3AppPresentator_"
        .GroupItems("Background2").Name = "PresSld2Shape4AppPresentator_"
        .GroupItems("Slide2Title").Name = "PresSld2Shape5AppPresentator_"
        .GroupItems("Slide2Text").Name = "PresSld2Shape6AppPresentator_"
    End With
End Sub

Sub TestSetupPageChange()
    SetupPageChange ActivePresentation.SlideShowWindow
End Sub

Sub SetupPageChange(ByVal oSW As SlideShowWindow)
    Dim Colors() As Variant
    Colors = Array("Aqua", "Purple", "Orange", "Blue", "Red", "Yellow", "Lime", "Green", "Black", "Gray")
    Dim Color As Variant
    If oSW.View.CurrentShowPosition = 31 Then
        Slide26.Shapes("MigNextButton").Visible = msoFalse
        Slide26.Shapes("UpgNextButton").Visible = msoFalse
        Slide26.Shapes("FrNextButton").Visible = msoTrue
        Slide26.Shapes("WarningA").Visible = msoFalse
        Slide26.Shapes("WarningB").Visible = msoFalse
        Slide26.Shapes("Migrate/1").Line.ForeColor.RGB = RGB(0, 0, 0)
        Slide26.Shapes("Upgrade/1").Line.ForeColor.RGB = RGB(0, 0, 0)
        Slide26.Shapes("FreshInstall/1").Line.ForeColor.RGB = RGB(255, 255, 255)
        
        With Slide37.Master.Theme
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
        SetupRadioCheck Slide37.Shapes("Blue/1")
    ElseIf oSW.View.CurrentShowPosition = 35 Then
        Slide34.TextBox1.Text = ""
        Slide34.TextBox2.Text = ""
        Slide34.TextBox3.Text = ""
        Slide34.TextBox4.Text = ""
        Slide34.TextBox5.Text = ""
        Slide34.Shapes("AutologinCheck").GroupItems("Check").Fill.ForeColor.RGB = RGB(34, 0, 96)
    ElseIf oSW.View.CurrentShowPosition = 37 Then
        Dim ReviewText As String
        ReviewText = "Data import:" & vbNewLine
        If IsRadioChecked("1", "FreshInstall", Slide31) Then
            ReviewText = ReviewText & "N/A" & vbNewLine & vbNewLine
        Else
            ReviewText = ReviewText & "Import userdata from another PPTOS" & vbNewLine & vbNewLine
        End If
        ReviewText = ReviewText & "User accounts:" & vbNewLine
        If Slide34.TextBox1.Text <> "" Then ReviewText = ReviewText & Slide34.TextBox1.Text & vbNewLine
        If Slide34.TextBox2.Text <> "" Then ReviewText = ReviewText & Slide34.TextBox2.Text & vbNewLine
        If Slide34.TextBox3.Text <> "" Then ReviewText = ReviewText & Slide34.TextBox3.Text & vbNewLine
        If Slide34.TextBox4.Text <> "" Then ReviewText = ReviewText & Slide34.TextBox4.Text & vbNewLine
        If Slide34.TextBox5.Text <> "" Then ReviewText = ReviewText & Slide34.TextBox5.Text & vbNewLine
        If Slide34.Shapes("Check").Fill.ForeColor.RGB = RGB(64, 0, 255) Then ReviewText = ReviewText & vbNewLine & "Autologin enabled" & vbNewLine
        ReviewText = ReviewText & vbNewLine & "Color scheme:" & vbNewLine
        For Each Color In Colors
            If IsRadioChecked("1", Color, Slide37) Then
                ReviewText = ReviewText & Color
            End If
        Next Color
        Slide38.Shapes("ReviewSetup").TextFrame.TextRange.Text = ReviewText
        Slide39.Shapes("Stage").TextFrame.TextRange.Text = "0"
    ElseIf oSW.View.CurrentShowPosition = 38 Then
        If Slide39.Shapes("Stage").TextFrame.TextRange.Text = "0" Then
            Dim PrefCol As String
            PrefCol = "1"
            Dim I As Integer
            I = 0
            For Each Color In Colors
                If IsRadioChecked("1", Color, Slide37) Then PrefCol = CStr(I)
                I = I + 1
            Next Color
            If FileStreamsExist("/Temp/MigrateData/") Then
                Slide1.Shapes("Username").TextFrame.TextRange.Text = Slide34.TextBox1.Text
                CopyFile "/Temp/MigrateData/Background.png", "/Users/" & Slide34.TextBox1.Text & "/"
            End If
            If Slide34.TextBox1.Text <> "" Then SetupCreateUser Slide34.TextBox1.Text, PrefCol
            If Slide34.TextBox2.Text <> "" Then SetupCreateUser Slide34.TextBox2.Text, PrefCol
            If Slide34.TextBox3.Text <> "" Then SetupCreateUser Slide34.TextBox3.Text, PrefCol
            If Slide34.TextBox4.Text <> "" Then SetupCreateUser Slide34.TextBox4.Text, PrefCol
            If Slide34.TextBox5.Text <> "" Then SetupCreateUser Slide34.TextBox5.Text, PrefCol
            
            If Slide34.Shapes("Check").Fill.ForeColor.RGB = RGB(64, 0, 255) Then
                SetFileContent "/System/Settings.cnf", Slide34.TextBox1.Text, "Autologin"
            Else
                SetFileContent "/System/Settings.cnf", "Nobody", "Autologin"
            End If
            SetFileContent "/System/Settings.cnf", "5", "AutosaveInterval"
            SaveSysConfig "Autorun", "Welcome;"
            If FileStreamsExist("/Temp/MigrateData/") Then
                Slide1.Shapes("Username").TextFrame.TextRange.Text = Slide34.TextBox1.Text
                CopyFile "/Temp/MigrateData/Presentator/Presentation.pres", "/Users/" & Slide34.TextBox1.Text & "/Documents/"
                CopyFile "/Temp/MigrateData/Words/Document.wdoc", "/Users/" & Slide34.TextBox1.Text & "/Documents/"
                CopyFile "/Temp/MigrateData/Words/Document2.wdoc", "/Users/" & Slide34.TextBox1.Text & "/Documents/"
                CopyFile "/Temp/MigrateData/Words/Document3.wdoc", "/Users/" & Slide34.TextBox1.Text & "/Documents/"
                CopyFile "/Temp/MigrateData/Words/Document4.wdoc", "/Users/" & Slide34.TextBox1.Text & "/Documents/"
                CopyFile "/Temp/MigrateData/Words/Document5.wdoc", "/Users/" & Slide34.TextBox1.Text & "/Documents/"
                DeleteDir "/Temp/MigrateData"
                Slide1.Shapes("Username").TextFrame.TextRange.Text = "Nobody"
                ActivePresentation.SlideShowWindow.View.GotoSlide 38
            End If
        End If
    End If
End Sub

Sub TestCopy()
    CopyFile "/Temp/MigrateData/Words/Document2.pres", "/Users/" & Slide34.TextBox1.Text & "/Documents/"
End Sub

Sub SetupCreateUser(ByVal Username As String, ByVal Color As String)
    Slide1.Shapes("Username").TextFrame.TextRange.Text = "Nobody"
    SetFileContent "/Users/" & Username & "/Password.txt", ""
    If Not FileStreamsExist("/Users/" & Username & "/Background.png") Then
        Slide15.Export Environ("TEMP") & "\Userpic.PNG", "PNG"
        Slide1.Shapes("Username").TextFrame.TextRange.Text = Username
        SetFilePic "/Users/" & Username & "/Background.png", Environ("TEMP") & "\Userpic.PNG"
    End If
    SetFileContent "/Users/" & Username & "/Password.txt", ""
    SetFileContent "/Users/" & Username & "/Theme.txt", Color
End Sub

Sub SetupCheckUncheckAutologin(Shp As Shape)
    Dim CheckShp As Shape
    Set CheckShp = Shp.ParentGroup.GroupItems("Check")
    If CheckShp.Fill.ForeColor.RGB = RGB(64, 0, 255) Then
        CheckShp.Fill.ForeColor.RGB = RGB(34, 0, 96)
    Else
        CheckShp.Fill.ForeColor.RGB = RGB(64, 0, 255)
    End If
End Sub

Sub PreviewSetupTheme()
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
    ThemeAqua = Array(RGB(52, 148, 186), RGB(88, 182, 192), RGB(117, 189, 167), RGB(122, 140, 142), RGB(132, 172, 182), RGB(38, 131, 198), RGB(0, 0, 0), RGB(55, 53, 69), RGB(255, 255, 255), RGB(206, 219, 230))
    ThemeMagenta = Array(RGB(146, 39, 143), RGB(155, 87, 211), RGB(117, 93, 217), RGB(102, 94, 184), RGB(102, 94, 184), RGB(117, 93, 217), RGB(0, 0, 0), RGB(99, 46, 98), RGB(255, 255, 255), RGB(234, 229, 235))
    ThemeOrange = Array(RGB(240, 127, 9), RGB(159, 41, 54), RGB(78, 165, 216), RGB(78, 133, 66), RGB(240, 127, 9), RGB(193, 152, 89), RGB(0, 0, 0), RGB(50, 50, 50), RGB(255, 255, 255), RGB(227, 222, 209))
    ThemeBlue = Array(RGB(68, 114, 196), RGB(237, 125, 49), RGB(165, 165, 165), RGB(255, 192, 0), RGB(91, 155, 213), RGB(112, 173, 71), RGB(0, 0, 0), RGB(68, 84, 106), RGB(255, 255, 255), RGB(231, 230, 230))
    ThemeRed = Array(RGB(165, 48, 15), RGB(213, 88, 22), RGB(242, 213, 167), RGB(177, 156, 125), RGB(144, 66, 66), RGB(178, 125, 73), RGB(0, 0, 0), RGB(50, 50, 50), RGB(255, 255, 255), RGB(232, 186, 118))
    ThemeBlack = Array(RGB(40, 40, 40), RGB(55, 55, 55), RGB(138, 138, 138), RGB(63, 63, 63), RGB(38, 38, 38), RGB(12, 12, 12), RGB(55, 55, 55), RGB(22, 22, 22), RGB(255, 255, 255), RGB(248, 248, 248))
    ThemeYellow = Array(RGB(240, 162, 46), RGB(165, 100, 78), RGB(181, 139, 128), RGB(195, 152, 109), RGB(161, 149, 116), RGB(193, 117, 41), RGB(0, 0, 0), RGB(78, 59, 48), RGB(255, 255, 255), RGB(251, 238, 201))
    ThemeLime = Array(RGB(153, 203, 56), RGB(99, 165, 55), RGB(55, 167, 111), RGB(68, 193, 163), RGB(78, 179, 207), RGB(81, 195, 249), RGB(0, 0, 0), RGB(69, 95, 81), RGB(255, 255, 255), RGB(226, 223, 204))
    ThemeGreen = Array(RGB(84, 158, 57), RGB(138, 184, 51), RGB(122, 207, 59), RGB(2, 150, 118), RGB(74, 181, 196), RGB(9, 137, 177), RGB(0, 0, 0), RGB(69, 95, 81), RGB(255, 255, 255), RGB(227, 222, 209))
    ThemeGray = Array(RGB(153, 153, 153), RGB(121, 121, 121), RGB(150, 150, 150), RGB(128, 128, 128), RGB(95, 95, 95), RGB(77, 77, 77), RGB(0, 0, 0), RGB(27, 27, 27), RGB(255, 255, 255), RGB(248, 248, 248))
    If IsRadioChecked("1", "Aqua", Slide37) Then
        Theme = ThemeAqua
    ElseIf IsRadioChecked("1", "Purple", Slide37) Then
        Theme = ThemeMagenta
    ElseIf IsRadioChecked("1", "Orange", Slide37) Then
        Theme = ThemeOrange
    ElseIf IsRadioChecked("1", "Blue", Slide37) Then
        Theme = ThemeBlue
    ElseIf IsRadioChecked("1", "Red", Slide37) Then
        Theme = ThemeRed
    ElseIf IsRadioChecked("1", "Yellow", Slide37) Then
        Theme = ThemeYellow
    ElseIf IsRadioChecked("1", "Lime", Slide37) Then
        Theme = ThemeLime
    ElseIf IsRadioChecked("1", "Green", Slide37) Then
        Theme = ThemeGreen
    ElseIf IsRadioChecked("1", "Black", Slide37) Then
        Theme = ThemeBlack
    ElseIf IsRadioChecked("1", "Gray", Slide37) Then
        Theme = ThemeGray
    End If
    
    With Slide37.Master.Theme
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
End Sub