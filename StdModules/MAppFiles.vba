' Files app
Sub AppFiles(Shp As Shape, Optional Path As String = "/")
    If Slide1.Shapes("Username").TextFrame.TextRange.Text = "Guest" Then
        AppMessage "Guests can't access file explorer", "File Explorer", "Error", True
        Exit Sub
    End If
    Dim ParentID As String
    ParentID = GetAppID(Shp)
    If ParentID <> "" Then
        If CInt(ParentID) <> -1 Then Shp.ParentGroup.Delete
    End If
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Files"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide1.Shapes("HomeAppFiles:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).Line.ForeColor.RGB = Slide1.Shapes("ButtonReloadAppFiles:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Font.Color.RGB
    Slide1.Shapes("WindowTitleAppFiles:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "File Explorer"
    Slide1.Shapes("PathAppFiles:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = Path
    Reload Slide1.Shapes("AppID").TextFrame.TextRange.Text
    UpdateTime
    FocusWindow Slide1.Shapes("AppID").TextFrame.TextRange.Text
End Sub

Sub AppFilesRestore(AppID As String)
    Dim Shp As Shape
    For Each Shp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
        Dim Numeral As Integer
        If InStr(1, Shp.Name, "FileLabel") = 1 Then
            Numeral = CInt(Replace(Replace(Shp.Name, "FileLabel", ""), "AppFiles:" & AppID, ""))
        ElseIf InStr(1, Shp.Name, "FileIcon") = 1 Then
            Numeral = CInt(Replace(Replace(Shp.Name, "FileIcon", ""), "AppFiles:" & AppID, ""))
        End If
        If Numeral > 18 Then
            Shp.Visible = msoFalse
        End If
    Next Shp
End Sub

Sub RenameFileHandler(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    If InStr(1, Slide1.Shapes("BottomPanelAppFiles:" & AppID).TextFrame.TextRange.Text, "Selected file: ") <> 1 Then
        AppMessage "No file selected.", "Files", "Exclamation", True
        Exit Sub
    End If
    Dim CName As String
    CName = Replace(Slide1.Shapes("BottomPanelAppFiles:" & AppID).TextFrame.TextRange.Text, "Selected file: ", "")
    
    SetVar "Macro", "HandleRename"
    SetVar "AppID", AppID
    SetVar "CurrentName", CName
    SetVar "FullPath", Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & CName
    AppInputBox "Enter a new name for '" & CName & "'", "Files"
End Sub

Sub HandleRenameDir(Shp As Shape)
    Dim AppID As String
    AppID = CheckVars("%AppID%")
    Dim Newname As String
    Dim Oldname As String
    Dim Fullpath As String
    Newname = CheckVars("%InputValue%")
    Fullpath = CheckVars("%FullPath%")
    Oldname = CheckVars("%CurrentName%")
    RenameDir Fullpath, Replace(Fullpath, "/" & Oldname, "/" & Newname)
    UnsetVar "InputValue"
    UnsetVar "FullPath"
    UnsetVar "CurrentName"
    UnsetVar "AppID"
    Shp.ParentGroup.Delete
    If Not FileStreamsExist(Fullpath) Then
        Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text = Replace(Fullpath, "/" & Oldname, "/" & Newname)
        Reload AppID
        Shp.ParentGroup.Delete
    End If
End Sub

Sub HandleRename(Shp As Shape)
    Dim AppID As String
    AppID = CheckVars("%AppID%")
    Dim Newname As String
    Dim Oldname As String
    Dim Fullpath As String
    Newname = CheckVars("%InputValue%")
    Fullpath = CheckVars("%FullPath%")
    Oldname = CheckVars("%CurrentName%")
    RenameFile Fullpath, Replace(Fullpath, "/" & Oldname, "/" & Newname)
    UnsetVar "InputValue"
    UnsetVar "FullPath"
    UnsetVar "CurrentName"
    UnsetVar "AppID"
    Shp.ParentGroup.Delete
    If Not FileExists(Fullpath) Then
        Reload AppID
        Shp.ParentGroup.Delete
    End If
End Sub

Sub NewFolderHandler(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    SetVar "Macro", "HandleCreateDir"
    SetVar "AppID", AppID
    AppInputBox "Please enter a name for the directory", "Files"
End Sub

Sub HandleCreateDir(Shp As Shape)
    Dim AppID As String
    AppID = CheckVars("%AppID%")
    Dim Dirname As String
    Dirname = CheckVars("%InputValue%")
    Dirname = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Dirname
    NewFolder Dirname
    If FileStreamsExist(Dirname) Then
        Reload AppID
    End If
    Shp.ParentGroup.Delete
    If AAX Then Slide1.AxTextBox.Visible = False
End Sub

Sub CheckUncheckFileMan(Shp As Shape)
    CheckUncheck Shp
    Reload GetAppID(Shp)
End Sub

Function FileCount(ByVal AppID As String) As Integer
    FileCount = CInt(Replace(Slide1.Shapes("RegularApp:" & AppID).GroupItems("BottomPanelAppFiles:" & AppID).TextFrame.TextRange.Text, " items", ""))
End Function

Function GetAssoc(ByVal Filename As String, ByVal AppID As String)
    Dim FinalName As String
    FinalName = Filename
    If InStr(Filename, "*") Then
        FakenameSplit = Split(Filename, "*")
        FinalName = FakenameSplit(0)
    End If
    ' If we don't do this check, the file explorer would display any file without an extension incorrectly
    If Replace(FinalName, ".", "") = FinalName Then
        If Slide1.Shapes("SelectModeCheckAppFiles:" & AppID).Fill.ForeColor.RGB = Slide1.ColorScheme.Colors(ppBackground).RGB Then
            GetAssoc = FsAssoc("")
        Else
            GetAssoc = "General"
        End If
        Exit Function
    End If
    Dim Extension As String
    FileSplit = Split(FinalName, ".")
    Extension = LCase(FileSplit(1))
    If Slide1.Shapes("SelectModeCheckAppFiles:" & AppID).Fill.ForeColor.RGB = Slide1.ColorScheme.Colors(ppBackground).RGB Then
        GetAssoc = FsAssoc(Extension)
    Else
        GetAssoc = "General"
    End If
End Function

Sub AssocGeneral(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Slide1.Shapes("BottomPanelAppFiles:" & AppID).TextFrame.TextRange.Text = "Selected file: " & Shp.TextFrame.TextRange.Text
End Sub

Sub PasteTest()
    PasteFile Slide1.Shapes("PasteButtonAppFiles:15")
End Sub

Sub PasteFile(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim Source As String
    Dim Destination As String
    Source = Slide1.Shapes("WindowAppFiles:" & AppID).TextFrame.TextRange.Text
    If Source <> "" Then
        Destination = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text
        CopyFile Source, Destination
        Slide1.Shapes("WindowAppFiles:" & AppID).TextFrame.TextRange.Text = ""
        If FileExists(Destination) Then
            Reload AppID
        End If
    Else
        AppMessage "Clipboard is empty", "Files", "Exclamation", True
    End If
End Sub

Sub AssocIGeneral(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocGeneral Slide1.Shapes(ShapeName)
End Sub

Sub FilesAddToClipboard(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    If InStr(1, Slide1.Shapes("BottomPanelAppFiles:" & AppID).TextFrame.TextRange.Text, "Selected file: ") <> 1 Then
        AppMessage "Cannot copy: No file selected", "Files", "Exclamation", True
        Exit Sub
    Else
        Dim FileSelector As String
        FileSelector = Slide1.Shapes("BottomPanelAppFiles:" & AppID).TextFrame.TextRange.Text
        Slide1.Shapes("WindowAppFiles:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Right(FileSelector, Len(FileSelector) - Len("Selected file: "))
        AppMessage "Added '" & Slide1.Shapes("WindowAppFiles:" & AppID).TextFrame.TextRange.Text & "' to clipboard", "Files", "Info", True
        Exit Sub
    End If
End Sub

Sub FilesDelete(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    If InStr(1, Slide1.Shapes("BottomPanelAppFiles:" & AppID).TextFrame.TextRange.Text, "Selected file: ") <> 1 Then
        Dim Dirname As String
        Dirname = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text
        DeleteDir Dirname
        If Not FileStreamsExist(Dirname) Then
            GoUp AppID
            AppMessage "The directory '" & Dirname & "' has been deleted", "Files", "Info", True
        End If
        Exit Sub
    Else
        Dim FileSelector As String
        FileSelector = Slide1.Shapes("BottomPanelAppFiles:" & AppID).TextFrame.TextRange.Text
        Dim Deletable As String
        Dim Stream As String
        Stream = ""
        Deletable = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Right(FileSelector, Len(FileSelector) - Len("Selected file: "))
        If InStr(Deletable, "*") Then
            FileSplit = Split(Deletable, "*")
            Deletable = FileSplit(0)
            Stream = FileSplit(1)
            DeleteFile Deletable, Stream
            If Not FileExists(Deletable, Stream) Then
                Reload AppID
                AppMessage "Deleted stream '" & Stream & "' of file '" & Deletable & "'", "Files", "Info", True
            End If
        Else
            DeleteFile Deletable
            If Not FileExists(Deletable) Then
                Reload AppID
                AppMessage "Deleted file '" & Deletable & "'", "Files", "Info", True
            End If
        End If
        Exit Sub
    End If
End Sub

Sub CleanIcons(AppID As String)
    Dim Lim As Integer
    Lim = Slide1.Shapes.Count
    Dim I As Integer
    For I = Lim To 1 Step -1
        If InStr(1, Slide1.Shapes(I).Name, "FileLabel") = 1 Or InStr(1, Slide1.Shapes(I).Name, "FileIcon") = 1 Then
            Slide1.Shapes(I).Delete
        End If
    Next I
    Lim = Slide1.Shapes("RegularApp:" & AppID).GroupItems.Count
    For I = Lim To 1 Step -1
        If InStr(1, Slide1.Shapes("RegularApp:" & AppID).GroupItems(I).Name, "FileLabel") = 1 Or InStr(1, Slide1.Shapes("RegularApp:" & AppID).GroupItems(I).Name, "FileIcon") = 1 Then
            Slide1.Shapes("RegularApp:" & AppID).GroupItems(I).Delete
        End If
    Next I
End Sub

Sub ShowDesktop()
    If Slide1.Shapes("ResizeEvent").TextFrame.TextRange.Text = "True" Then Exit Sub
    If Slide1.Shapes("MoveEvent").TextFrame.TextRange.Text = "True" Then Exit Sub
    If FileStreamsExist("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop/") Then
        Slide1.Shapes("PathAppFiles:-1").TextFrame.TextRange.Text = "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop/"
        Reload -1
    Else
        Dim Shp As Shape
        Dim x As Integer
        x = Slide1.Shapes("RegularApp:-1").GroupItems.Count
        For x = x To 1 Step -1
            Set Shp = Slide1.Shapes("RegularApp:-1").GroupItems(x)
            If (InStr(1, Shp.Name, "FileLabel") = 1) Or (InStr(1, Shp.Name, "FileIcon") = 1) Then
                Shp.Delete
            End If
        Next x
    End If
    Slide1.Shapes("RegularApp:-1").ZOrder msoSendToBack
    Slide1.Shapes("BackgroundImg").ZOrder msoSendToBack
    Slide1.Shapes("AnimationRect").ZOrder msoSendToBack
End Sub

Sub MeasureTest()
    Dim x As Integer
    Dim y As Integer
    x = (Slide1.Shapes("InnerWindowAppFiles:125").Width - 20) / (GetFileRef("/Defaults/Icons/Folder.emf").Width + 25)
    y = (Slide1.Shapes("InnerWindowAppFiles:125").Height - 20) / (GetFileRef("/Defaults/Icons/Folder.emf").Height + 30)
    MsgBox x & "x" & y
End Sub

Sub TestReload()
    Reload "34"
End Sub

Sub Reload(AppID As String, Optional ByVal Attempt As Integer = 1)
    On Error GoTo Except
    If AppID = "-1" Then
        On Error Resume Next
    End If
    FocusWindow AppID
    Dim Shp As Shape
    Set Shp = Slide1.Shapes("ButtonReloadAppFiles:" & AppID)
    Dim Sld As Slide
    Dim Ref As Shape
    Dim MaxWidth As Integer
    Dim MaxHeight As Integer
    Dim Dirname As String
    Set Sld = Slide1
    Set Ref = Sld.Shapes("RegularApp:" & AppID).GroupItems("InnerWindowAppFiles:" & AppID)
    Dirname = Sld.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text
    
    If Not FileStreamsExist(Dirname) Then
        If (Not FileStreamsExist(Dirname)) And (Dir(Dirname, vbDirectory) = "") Then
            Sld.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text = "/"
            Reload AppID, 1
            AppMessage "Directory does not exist", "Files", "Error", True
            Exit Sub
        End If
    End If
    Dim I As Integer
    CleanIcons AppID
    WaitCursor Slide1.Shapes("InnerWindowAppFiles:" & AppID), "Navigating..."
    Dim File As Variant
    Dim OffsetX As Integer
    Dim OffsetY As Integer
    Dim IDX As Integer
    OffsetX = Ref.Left + 10
    OffsetY = Ref.Top + 10
    If AppID <> "-1" Then
        'MaxWidth = 6
        'MaxHeight = 18
        MaxWidth = FilesGetMaxWidth(AppID)
        MaxHeight = FilesGetMaxHeight(AppID) * MaxWidth
    Else
        MaxWidth = 11
        MaxHeight = 66
    End If
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
                .Name = "FileLabel" & CStr(IDX) & "AppFiles_"
                .TextFrame.TextRange.Text = File
                .TextFrame.TextRange.Font.Name = "Candara"
                .TextFrame.TextRange.Font.Size = 11
                .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
                If HideMe Then
                    .Visible = msoFalse
                Else
                    .Visible = msoTrue
                End If
                If IsFolder Then
                    .ActionSettings(ppMouseClick).Run = "NavigateFolder"
                End If
            End With
            If IsFolder Then
                PasteToGroup Shp, GetFileRef("/Defaults/Icons/Folder.emf"), "FileIcon" & CStr(IDX) & "AppFiles:" & AppID, OffsetX + 10, OffsetY, Slide1, "NavigateIFolder"
                Slide1.Shapes("FileIcon" & CStr(IDX) & "AppFiles:" & AppID).Visible = msoTrue
                PasteToGroup Shp, Slide1.Shapes("FileLabel" & CStr(IDX) & "AppFiles_"), "FileLabel" & CStr(IDX) & "AppFiles:" & AppID, OffsetX, OffsetY + GetFileRef("/Defaults/Icons/Folder.emf").Height, Slide1, "NavigateFolder"
            Else
                Dim Assoc As String
                Assoc = GetAssoc(File, AppID)
                'If Assoc = "" Then
                '    PasteToGroup Shp, GetFileRef("/Defaults/Icons/Any.emf"), "FileIcon" & CStr(IDX) & "AppFiles:" & AppID, OffsetX + 10, OffsetY, Slide1
               '     PasteToGroup Shp, Slide1.Shapes("FileLabel" & CStr(IDX) & "AppFiles_"), "FileLabel" & CStr(IDX) & "AppFiles:" & AppID, OffsetX, OffsetY + GetFileRef("/Defaults/Icons/Folder.emf").Height, Slide1
                'Else
                    Dim IconType As String
                    IconType = "Any"
                    NameSplit = Split(File, ".")
                    NameExt = NameSplit(UBound(NameSplit))
                    FAssoc = FsAssoc(LCase(NameExt))
                    Dim MacroName As String
                    MacroName = ""
                    If Assoc <> "" Then
                        MacroName = "AssocI" & Assoc
                    End If
                    If FAssoc = "Notes" Then
                        IconType = "Txt"
                    ElseIf FAssoc = "PictureView" Then
                        IconType = "Pic"
                    ElseIf FAssoc = "Paint" Then
                        IconType = "Pxl"
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
                    If IconType <> "App" Then
                        Dim EmfFile As String
                        EmfFile = "/Defaults/Icons/" & IconType & ".emf"
                        If IconType = "Custom" Then
                            EmfFile = GetSysConfig("Icon" & FAssoc)
                        End If
                        PasteToGroup Shp, GetFileRef(EmfFile), "FileIcon" & CStr(IDX) & "AppFiles:" & AppID, OffsetX + 10, OffsetY, Slide1, MacroName
                        Slide1.Shapes("FileIcon" & CStr(IDX) & "AppFiles:" & AppID).Visible = msoTrue
                    Else
                        Dim AppName2 As String
                        AppName2 = Replace(File, ".app", "")
                        Dim HasIt As Boolean
                        HasIt = False
                        Dim BShp As Shape
                        For Each BShp In Slide25.Shapes
                            If BShp.Name = "App" & AppName2 & ":Icon" Then
                                Slide25.Shapes("App" & AppName2 & ":Icon").Copy
                                With Slide1.Shapes.Paste
                                    .Name = "DummyIcon"
                                    .Visible = msoTrue
                                End With
                                PasteToGroup Shp, Slide1.Shapes("DummyIcon"), "FileIcon" & CStr(IDX) & "AppFiles:" & AppID, OffsetX + 10, OffsetY, Slide1, "AssocI" & Assoc
                                Slide1.Shapes("DummyIcon").Delete
                                HasIt = True
                            End If
                        Next BShp
                        If HasIt = False Then
                            PasteToGroup Shp, GetFileRef("/Defaults/Icons/Any.emf"), "FileIcon" & CStr(IDX) & "AppFiles:" & AppID, OffsetX + 10, OffsetY, Slide1, "AssocI" & Assoc
                        End If
                    End If
                    PasteToGroup Shp, Slide1.Shapes("FileLabel" & CStr(IDX) & "AppFiles_"), "FileLabel" & CStr(IDX) & "AppFiles:" & AppID, OffsetX, OffsetY + GetFileRef("/Defaults/Icons/Folder.emf").Height, Slide1, Replace(MacroName, "AssocI", "Assoc")
                'End If
                Slide1.Shapes("FileIcon" & CStr(IDX) & "AppFiles:" & AppID).Visible = msoTrue
            End If
            If AppID = "-1" Then
                With Slide1.Shapes("FileLabel" & CStr(IDX) & "AppFiles:-1").TextFrame.TextRange.Font
                    .Color = RGB(255, 255, 255)
                    .Shadow = msoTrue
                End With
            End If
            If HideMe Then
                Slide1.Shapes("FileIcon" & CStr(IDX) & "AppFiles:" & AppID).Visible = msoFalse
                Slide1.Shapes("FileLabel" & CStr(IDX) & "AppFiles:" & AppID).Visible = msoFalse
            End If
            Slide1.Shapes("FileLabel" & CStr(IDX) & "AppFiles_").Delete
            If IDX Mod MaxWidth = 0 Then
                OffsetX = Ref.Left + 10
                OffsetY = OffsetY + GetFileRef("/Defaults/Icons/Folder.emf").Height + 30
            Else
                OffsetX = OffsetX + GetFileRef("/Defaults/Icons/Folder.emf").Width + 25
            End If
            If IDX Mod MaxHeight = 0 Then
                OffsetX = Ref.Left + 10
                OffsetY = Ref.Top + 10
                HideMe = True
            End If
            IDX = IDX + 1
        End If
    Next File
    Slide1.Shapes("BottomPanelAppFiles:" & AppID).TextFrame.TextRange.Text = CStr(IDX - 1) & " items"
    HideCursor
    Exit Sub
Except:
    Regroup AppID, Slide1
    Attempt = Attempt + 1
    If Attempt > 10 Then Exit Sub
    Pause 1
    Reload AppID, Attempt
    HideCursor
    Exit Sub
End Sub


Sub FilesLoadDir(RShp As Shape)
    Dim AppID As String
    AppID = GetAppID(RShp)
    WaitCursor Slide1.Shapes("InnerWindowAppFiles:" & AppID)
    Reload AppID
End Sub

Function FilesVisibleCount(AppID As String) As Integer
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

Function MaximumVisible(AppID As String) As Integer
    Dim MaxVis As Integer
    Dim Shp As Shape
    MaxVis = 0
    MaxFile = 0
    For Each Shp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
        If InStr(1, Shp.Name, "FileLabel") = 1 Then
            If Shp.Visible = msoTrue Then
                Dim SValue As Integer
                SValue = CInt(Replace(Replace(Shp.Name, "FileLabel", ""), "AppFiles:" & AppID, ""))
                If SValue > MaxVis Then
                    MaxVis = SValue
                End If
            ElseIf Shp.Visible = msoFalse Then
                Dim ISValue As Integer
                ISValue = CInt(Replace(Replace(Shp.Name, "FileLabel", ""), "AppFiles:" & AppID, ""))
                If ISValue > MaxFile Then
                    MaxFile = ISValue
                End If
            End If
        End If
    Next Shp
    MaximumVisible = MaxVis
End Function

Function MinimumVisible(AppID As String) As Integer
    Dim MinVis As Integer
    Dim Shp As Shape
    MinVis = 32767
    MinFile = 32767
    For Each Shp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
        If InStr(1, Shp.Name, "FileLabel") = 1 Then
            If Shp.Visible = msoTrue Then
                Dim SValue As Integer
                SValue = CInt(Replace(Replace(Shp.Name, "FileLabel", ""), "AppFiles:" & AppID, ""))
                If SValue < MinVis Then
                    MinVis = SValue
                End If
            ElseIf Shp.Visible = msoFalse Then
                Dim ISValue As Integer
                ISValue = CInt(Replace(Replace(Shp.Name, "FileLabel", ""), "AppFiles:" & AppID, ""))
                If ISValue < MinFile Then
                    MinFile = ISValue
                End If
            End If
        End If
    Next Shp
    MinimumVisible = MinVis
End Function

Function FilesGetMaxWidth(ByVal AppID As String) As Integer
    FilesGetMaxWidth = (Slide1.Shapes("InnerWindowAppFiles:" & AppID).Width - 20) / (GetFileRef("/Defaults/Icons/Folder.emf").Width + 25)
    If FilesGetMaxWidth * (GetFileRef("/Defaults/Icons/Folder.emf").Width + 25) + 10 >= Slide1.Shapes("InnerWindowAppFiles:" & AppID).Width Then
        FilesGetMaxWidth = FilesGetMaxWidth - 1
    End If
End Function

Function FilesGetMaxHeight(ByVal AppID As String) As Integer
    FilesGetMaxHeight = (Slide1.Shapes("InnerWindowAppFiles:" & AppID).Height - 20) / (GetFileRef("/Defaults/Icons/Folder.emf").Height + 30)
    If FilesGetMaxHeight * (GetFileRef("/Defaults/Icons/Folder.emf").Height + 30) + 10 >= Slide1.Shapes("InnerWindowAppFiles:" & AppID).Height Then
        FilesGetMaxHeight = FilesGetMaxHeight - 1
    End If
End Function

Sub FilesNextPage(Ref As Shape)
    On Error Resume Next
    Dim AppID As String
    Dim HideUpUntil As Integer
    Dim ShowFrom As Integer
    Dim ShowTo As Integer
    Dim VisibleCount As Integer
    Dim MaxWidth As Integer
    Dim MaxHeight As Integer
    AppID = GetAppID(Ref)
    MaxWidth = FilesGetMaxWidth(AppID)
    MaxHeight = FilesGetMaxHeight(AppID)
    VisibleCount = FilesVisibleCount(AppID)
    HideUpUntil = MaximumVisible(AppID)
    ShowFrom = HideUpUntil + 1
    ShowTo = ShowFrom + (MaxWidth * MaxHeight - 1)
    If VisibleCount < (MaxWidth * MaxHeight) Then Exit Sub
    Dim Shp As Shape
    ' Hide visible entries
    Dim I As Integer
    For I = 1 To HideUpUntil
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(I) & "AppFiles:" & AppID).Visible = msoFalse
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(I) & "AppFiles:" & AppID).Visible = msoFalse
    Next I
    For I = ShowFrom To ShowTo
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(I) & "AppFiles:" & AppID).Visible = msoTrue
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(I) & "AppFiles:" & AppID).Visible = msoTrue
    Next I
    Exit Sub
End Sub


Sub FilesLastPage(Ref As Shape)
    On Error Resume Next
    Dim AppID As String
    Dim HideUpUntil As Integer
    Dim ShowFrom As Integer
    Dim ShowTo As Integer
    Dim MaxWidth As Integer
    Dim MaxHeight As Integer
    Dim MaxDims As Integer
    AppID = GetAppID(Ref)
    MaxWidth = FilesGetMaxWidth(AppID)
    MaxHeight = FilesGetMaxHeight(AppID)
    MaxDims = MaxWidth * MaxHeight
    HideFrom = MinimumVisible(AppID)
    HideTo = HideFrom + (MaxDims - 1)
    ShowFrom = HideFrom - MaxDims
    ShowTo = HideFrom - 1
    If HideFrom = 1 And HideTo = MaxDims Then Exit Sub
    Dim Shp As Shape
    ' Hide visible entries
    Dim I As Integer
    For I = HideFrom To HideTo
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(I) & "AppFiles:" & AppID).Visible = msoFalse
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(I) & "AppFiles:" & AppID).Visible = msoFalse
    Next I
    For I = ShowFrom To ShowTo
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(I) & "AppFiles:" & AppID).Visible = msoTrue
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(I) & "AppFiles:" & AppID).Visible = msoTrue
    Next I
    Exit Sub
End Sub

Sub NavigateFolder(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    If AppID = "-1" Then
        If InStr(1, Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text, "/") = 1 Then
            AppFiles Shp, Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Shp.TextFrame.TextRange.Text
        Else
            AppFiles Shp, Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Left(Shp.TextFrame.TextRange.Text, Len(Shp.TextFrame.TextRange.Text) - 1) & "\"
        End If
    Else
        If InStr(1, Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text, "/") = 1 Then
            Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Shp.TextFrame.TextRange.Text
        Else
            Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Left(Shp.TextFrame.TextRange.Text, Len(Shp.TextFrame.TextRange.Text) - 1) & "\"
        End If
        Reload AppID
    End If
End Sub

Sub NavigateIFolder(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim LabelName As String
    LabelName = Replace(Shp.Name, "Icon", "Label")
    NavigateFolder Slide1.Shapes(LabelName)
End Sub

Sub FilesUp(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    GoUp AppID
End Sub

Sub GoUp(AppID As String)
    Dim Sld As Slide
    Dim Path As String
    Set Sld = Slide1
    Path = Sld.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text
    If InStr(1, Path, "/") = 1 Then
        Dim LenPath As Integer
        Dim SplitPath() As String
        SplitPath = Split(Path, "/")
        LenPath = UBound(SplitPath) - 1
        LastDir = SplitPath(LenPath)
        PrePath = Left(Path, Len(Path) - Len(LastDir) - 1)
        Sld.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text = PrePath
    Else
        SplitPath = Split(Path, "\")
        LastDir = SplitPath(UBound(SplitPath) - 1)
        PrePath = Left(Path, Len(Path) - Len(LastDir) - 1)
        Sld.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text = PrePath
    End If
    Reload AppID
End Sub

Sub GoHome(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text = "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/"
    Reload AppID
End Sub

Sub VFSRoot(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text = "/"
    Reload AppID
End Sub

Sub HostRoot(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text = "C:\"
    Reload AppID
End Sub

Sub AppFilesSizeChanged(AppID As String)
    Reload AppID
End Sub

Sub AppFilesChangeDir(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    SetVar "Macro", "AppFilesActuallyChangeDir"
    SetVar "AppID", AppID
    AppInputBox "Change directory to...", "Files", True
    Slide1.AxTextBox.Text = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text
End Sub

Sub AppFilesActuallyChangeDir()
    Dim AppID As String
    Dim InputValue As String
    AppID = CheckVars("%AppID%")
    InputValue = CheckVars("%InputValue%")
    If Right(InputValue, 1) <> "/" And Not InStr(1, InputValue, "C:\") Then
        InputValue = InputValue & "/"
    End If
    Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text = InputValue
    UnsetVar "AppID"
    UnsetVar "Macro"
    UnsetVar "InputValue"
    Reload AppID
End Sub