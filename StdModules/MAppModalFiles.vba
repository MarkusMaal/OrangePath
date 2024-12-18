' Modal files app
Sub AppModalFiles()
    If Slide1.Shapes("Username").TextFrame.TextRange.Text = "Guest" Then
        AppMessage "Access is denied", "File chooser", "Error", True
        Exit Sub
    End If
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "ModalFiles"
    LaunchPath = "/"
    If CheckVars("%LaunchDir%") <> "" And CheckVars("%LaunchDir%") <> "%LaunchDir%" Then
        LaunchPath = CheckVars("%LaunchDir%")
    End If
    Slide2.Shapes("PathAppModalFiles_").TextFrame.TextRange.Text = LaunchPath
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    If AAX Then Slide1.AxTextBox.Visible = False
    Slide1.Shapes("WindowTitleAppModalFiles:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "File chooser"
    If CheckVars("%Save%") = "" Or CheckVars("%Save%") = "%Save%" Then
        Slide1.Shapes("ButtonOkAppModalFiles:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).Delete
        Slide1.Shapes("AxTextBox1AppModalFiles:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).Delete
        If AAX Then Slide1.AxTextBox.Visible = False
    End If
    If CheckVars("%NoFs%") = "True" Then
        Slide1.Shapes("ButtonDriveSwitchAppModalFiles:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).Delete
        UnsetVar "NoFs"
    End If
    MReload Slide1.Shapes("AppID").TextFrame.TextRange.Text
End Sub

Function MFileCount(ByVal AppID As String) As Integer
    MFileCount = CInt(Replace(Slide1.Shapes("RegularApp:" & AppID).GroupItems("BottomPanelAppModalFiles:" & AppID).TextFrame.TextRange.Text, " items", ""))
End Function

Sub ModalFilesSwitchDrive(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    If InStr(1, Slide1.Shapes("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text, "C:\") = 1 Then
        Slide1.Shapes("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text = "/"
        MReload AppID
    Else
        Slide1.Shapes("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text = "C:\"
        MReload AppID
    End If
    
End Sub

Sub AssocModal(Shp As Shape)
    On Error GoTo InvalidFile
    Dim AppID As String
    AppID = GetAppID(Shp)
    If CheckVars("%Save%") = "" Or CheckVars("%Save%") = "%Save%" Then
        SetVar "InputValue", Slide1.Shapes("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text & Shp.TextFrame.TextRange.Text
        If CheckVars("%Macro%") <> "" And CheckVars("%Macro%") <> "%Macro%" Then
            Application.Run CheckVars("%Macro%"), Shp
        End If
        UnsetVar "Macro"
        CloseWindow Shp
    Else
        If AAX Then Slide1.AxTextBox.Text = Shp.TextFrame.TextRange.Text
        Slide1.Shapes("AxTextBox1AppModalFiles:" & AppID).TextFrame.TextRange.Text = Shp.TextFrame.TextRange.Text
    End If
    Exit Sub
InvalidFile:
    AppMessage "Cannot open files of this type", "Error", "Error", True
End Sub

Sub SaveFile(Shp As Shape)
    AppID = GetAppID(Shp)
    SetVar "InputValue", Slide1.Shapes("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes("AxTextBox1AppModalFiles:" & AppID).TextFrame.TextRange.Text
    If CheckVars("%Macro%") <> "" And CheckVars("%Macro%") <> "%Macro%" Then
        Application.Run CheckVars("%Macro%"), Shp
    End If
    UnsetVar "Macro"
    UnsetVar "Save"
    CloseWindow Shp
End Sub

Sub AssocIModal(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocModal Slide1.Shapes(ShapeName)
End Sub

Sub MReload(AppID As String, Optional ByVal Attempt As Integer = 1)
    On Error GoTo Except
    FocusWindow Slide1.Shapes("AppID").TextFrame.TextRange.Text
    Dim Shp As Shape
    Set Shp = Slide1.Shapes("ButtonReloadAppModalFiles:" & AppID)
    Dim Sld As Slide
    Dim Ref As Shape
    Dim Dirname As String
    Set Sld = Slide1
    Set Ref = Sld.Shapes("RegularApp:" & AppID).GroupItems("InnerWindowAppModalFiles:" & AppID)
    Dirname = Sld.Shapes("RegularApp:" & AppID).GroupItems("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text
    If (InStr(1, Dirname, "C:\") <> 1) And (Right(Dirname, 1) <> "/") Then
        Dirname = Dirname & "/"
        Sld.Shapes("RegularApp:" & AppID).GroupItems("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text = Dirname
    End If
    Dim I As Integer
    CleanIcons AppID
    WaitCursor Slide1.Shapes("InnerWindowAppModalFiles:" & AppID), "Navigating..."
    Dim File As Variant
    Dim OffsetX As Integer
    Dim OffsetY As Integer
    Dim IDX As Integer
    OffsetX = Ref.Left + 10
    OffsetY = Ref.Top + 10
    IDX = 1
    Dim HideMe As Boolean
    HideMe = False
    For Each File In Split(GetFiles(Dirname), vbNewLine)
        If File <> "" Then
            Dim IsFolder As Boolean
            IsFolder = False
            If Right(File, 1) = "/" Then
                IsFolder = True
            End If
            With Slide1.Shapes.AddTextbox(msoTextOrientationHorizontal, OffsetX - 10, OffsetY + GetFileRef("/Defaults/Icons/Folder.emf").Height, GetFileRef("/Defaults/Icons/Folder.emf").Width + 20, 20)
                .Name = "FileLabel" & CStr(IDX) & "AppModalFiles_"
                .TextFrame.TextRange.Text = File
                .TextFrame.TextRange.Font.Name = "Candara"
                .TextFrame.TextRange.Font.Size = 11
                .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
                If HideMe Then
                    .Visible = msoFalse
                End If
                If IsFolder Then
                    .ActionSettings(ppMouseClick).Run = "MNavigateFolder"
                End If
            End With
            If IsFolder Then
                PasteToGroup Shp, GetFileRef("/Defaults/Icons/Folder.emf"), "FileIcon" & CStr(IDX) & "AppModalFiles:" & AppID, OffsetX + 10, OffsetY, Slide1, "MNavigateIFolder"
                Slide1.Shapes("FileIcon" & CStr(IDX) & "AppModalFiles:" & AppID).Visible = msoTrue
                PasteToGroup Shp, Slide1.Shapes("FileLabel" & CStr(IDX) & "AppModalFiles_"), "FileLabel" & CStr(IDX) & "AppModalFiles:" & AppID, OffsetX, OffsetY + GetFileRef("/Defaults/Icons/Folder.emf").Height, Slide1, "MNavigateFolder"
            Else
                Dim IconType As String
                IconType = "Any"
                NameSplit = Split(File, ".")
                NameExt = NameSplit(UBound(NameSplit))
                FAssoc = FsAssoc(LCase(NameExt))
                If FAssoc = "Notes" Then
                    IconType = "Txt"
                ElseIf FAssoc = "Paint" Then
                    IconType = "Pxl"
                ElseIf FAssoc = "PictureView" Then
                    IconType = "Pic"
                ElseIf FAssoc = "VideoPlayer" Then
                    IconType = "Movie"
                ElseIf FAssoc = "3D" Then
                    IconType = "3D"
                ElseIf FAssoc = "Settings" Then
                    IconType = "Cnf"
                ElseIf FAssoc = "SoundPlayer" Then
                    IconType = "Snd"
                ElseIf FAssoc = "Presentator" Then
                    IconType = "Pres"
                ElseIf FAssoc = "Special" Then
                    IconType = "Special"
                ElseIf FAssoc = "Components" Then
                    IconType = "Grp"
                ElseIf FAssoc = "ModalHelpView" Then
                    IconType = "Hlp"
                ElseIf LCase(NameExt) = "thm" Then
                    IconType = "Thm"
                ElseIf GetSysConfig("Icon" & FAssoc) <> "*" Then
                    IconType = "Custom"
                End If
                Dim EmfFile As String
                EmfFile = "/Defaults/Icons/" & IconType & ".emf"
                If IconType = "Custom" Then
                    EmfFile = GetSysConfig("Icon" & FAssoc)
                End If
                PasteToGroup Shp, GetFileRef(EmfFile), "FileIcon" & CStr(IDX) & "AppModalFiles:" & AppID, OffsetX + 10, OffsetY, Slide1, "AssocIModal"
                Slide1.Shapes("FileIcon" & CStr(IDX) & "AppModalFiles:" & AppID).Visible = msoTrue
                PasteToGroup Shp, Slide1.Shapes("FileLabel" & CStr(IDX) & "AppModalFiles_"), "FileLabel" & CStr(IDX) & "AppModalFiles:" & AppID, OffsetX, OffsetY + GetFileRef("/Defaults/Icons/Folder.emf").Height, Slide1, "AssocModal"
            End If
            If HideMe Then
                Slide1.Shapes("FileIcon" & CStr(IDX) & "AppModalFiles:" & AppID).Visible = msoFalse
                Slide1.Shapes("FileLabel" & CStr(IDX) & "AppModalFiles:" & AppID).Visible = msoFalse
            End If
            Slide1.Shapes("FileLabel" & CStr(IDX) & "AppModalFiles_").Delete
            If IDX Mod 6 = 0 Then
                OffsetX = Ref.Left + 10
                OffsetY = OffsetY + GetFileRef("/Defaults/Icons/Folder.emf").Height + 30
            Else
                OffsetX = OffsetX + GetFileRef("/Defaults/Icons/Folder.emf").Width + 25
            End If
            If IDX Mod 18 = 0 Then
                OffsetX = Ref.Left + 10
                OffsetY = Ref.Top + 10
                HideMe = True
            End If
            IDX = IDX + 1
        End If
    Next File
    Slide1.Shapes("BottomPanelAppModalFiles:" & AppID).TextFrame.TextRange.Text = CStr(IDX - 1) & " items"
    HideCursor
    If CheckVars("%Save%") = "" Or CheckVars("%Save%") = "%Save%" Then
        If AAX Then Slide1.AxTextBox.Visible = False
    End If
    Exit Sub
Except:
    Regroup AppID, Slide1
    Attempt = Attempt + 1
    If Attempt > 10 Then Exit Sub
    Pause 1
    MReload AppID, Attempt
    HideCursor
    Exit Sub
End Sub

Sub MFilesLoadDir(RShp As Shape)
    Dim AppID As String
    AppID = GetAppID(RShp)
    MReload AppID
End Sub

Function MMaximumVisible(AppID As String) As Integer
    Dim MaxVis As Integer
    Dim Shp As Shape
    MaxVis = 0
    For Each Shp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
        If InStr(1, Shp.Name, "FileLabel") = 1 Then
            If Shp.Visible = msoTrue Then
                Dim SValue As Integer
                SValue = CInt(Replace(Replace(Shp.Name, "FileLabel", ""), "AppModalFiles:" & AppID, ""))
                If SValue > MaxVis Then
                    MaxVis = SValue
                End If
            End If
        End If
    Next Shp
    MMaximumVisible = MaxVis
End Function

Function MMinimumVisible(AppID As String) As Integer
    Dim MinVis As Integer
    Dim Shp As Shape
    MinVis = 32767
    For Each Shp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
        If InStr(1, Shp.Name, "FileLabel") = 1 Then
            If Shp.Visible = msoTrue Then
                Dim SValue As Integer
                SValue = CInt(Replace(Replace(Shp.Name, "FileLabel", ""), "AppModalFiles:" & AppID, ""))
                If SValue < MinVis Then
                    MinVis = SValue
                End If
            End If
        End If
    Next Shp
    MMinimumVisible = MinVis
End Function


Sub MFilesNextPage(Ref As Shape)
    Dim AppID As String
    Dim HideUpUntil As Integer
    Dim ShowFrom As Integer
    Dim ShowTo As Integer
    Dim VisibleCount As Integer
    AppID = GetAppID(Ref)
    VisibleCount = FilesVisibleCount(AppID)
    HideUpUntil = MMaximumVisible(AppID)
    ShowFrom = HideUpUntil + 1
    ShowTo = ShowFrom + 17
    If VisibleCount < 18 Then Exit Sub
    Dim Shp As Shape
    ' Hide visible entries
    Dim I As Integer
    For I = 1 To HideUpUntil
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(I) & "AppModalFiles:" & AppID).Visible = msoFalse
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(I) & "AppModalFiles:" & AppID).Visible = msoFalse
    Next I
    For I = ShowFrom To ShowTo
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(I) & "AppModalFiles:" & AppID).Visible = msoTrue
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(I) & "AppModalFiles:" & AppID).Visible = msoTrue
    Next I
    Exit Sub
End Sub

Sub MFilesLastPage(Ref As Shape)
    On Error Resume Next
    Dim AppID As String
    Dim HideUpUntil As Integer
    Dim ShowFrom As Integer
    Dim ShowTo As Integer
    AppID = GetAppID(Ref)
    HideFrom = MMinimumVisible(AppID)
    HideTo = HideFrom + 17
    ShowFrom = HideFrom - 18
    ShowTo = HideFrom - 1
    If HideFrom = 1 And HideTo = 18 Then Exit Sub
    Dim Shp As Shape
    ' Hide visible entries
    Dim I As Integer
    For I = HideFrom To HideTo
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(I) & "AppModalFiles:" & AppID).Visible = msoFalse
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(I) & "AppModalFiles:" & AppID).Visible = msoFalse
    Next I
    For I = ShowFrom To ShowTo
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(I) & "AppModalFiles:" & AppID).Visible = msoTrue
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(I) & "AppModalFiles:" & AppID).Visible = msoTrue
    Next I
    Exit Sub
End Sub

Sub MNavigateFolder(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    If InStr(1, Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text, "/") = 1 Then
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text & Shp.TextFrame.TextRange.Text
    Else
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text & Left(Shp.TextFrame.TextRange.Text, Len(Shp.TextFrame.TextRange.Text) - 1) & "\"
    End If
    MReload AppID
End Sub

Sub MNavigateIFolder(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim LabelName As String
    LabelName = Replace(Shp.Name, "Icon", "Label")
    MNavigateFolder Slide1.Shapes(LabelName)
End Sub

Sub MFilesUp(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    MGoUp AppID
End Sub

Sub MGoUp(AppID As String)
    Dim Sld As Slide
    Dim Path As String
    Set Sld = Slide1
    Path = Sld.Shapes("RegularApp:" & AppID).GroupItems("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text
    If InStr(1, Path, "/") = 1 Then
        Dim LenPath As Integer
        Dim SplitPath() As String
        SplitPath = Split(Path, "/")
        LenPath = UBound(SplitPath) - 1
        LastDir = SplitPath(LenPath)
        PrePath = Left(Path, Len(Path) - Len(LastDir) - 1)
        Sld.Shapes("RegularApp:" & AppID).GroupItems("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text = PrePath
    Else
        SplitPath = Split(Path, "\")
        LastDir = SplitPath(UBound(SplitPath) - 1)
        PrePath = Left(Path, Len(Path) - Len(LastDir) - 1)
        Sld.Shapes("RegularApp:" & AppID).GroupItems("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text = PrePath
    End If
    MReload AppID
End Sub



Function MFilesVisibleCount(AppID As String) As Integer
    Dim ShapeCount As Integer
    Dim Shp As Shape
    ShapeCount = 0
    For Each Shp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
        If InStr(1, Shp.Name, "FileLabel") = 1 Then
            If Shp.Visible = msoTrue Then
                ShapeCount = ShapeCount + 1
            End If
        End If
    Next Shp
    FilesVisibleCount = ShapeCount
End Function