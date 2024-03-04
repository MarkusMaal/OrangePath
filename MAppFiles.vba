' Files app
Sub AppFiles(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Files"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Reload Slide1.Shapes("AppID").TextFrame.TextRange.Text
End Sub

Sub DevTest()
    Reload "17"
End Sub

Sub RenameFileHandler(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    If InStr(1, Slide1.Shapes("BottomPanelAppFiles:" & AppID).TextFrame.TextRange.Text, "Selected file: ") <> 1 Then
        AppMessage "No file selected", "Files", "Exclamation", True
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
    Slide1.AxTextBox.Visible = False
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
    Dim i As Integer
    For i = Lim To 1 Step -1
        If InStr(1, Slide1.Shapes(i).Name, "FileLabel") = 1 Or InStr(1, Slide1.Shapes(i).Name, "FileIcon") = 1 Then
            Slide1.Shapes(i).Delete
        End If
    Next i
    Lim = Slide1.Shapes("RegularApp:" & AppID).GroupItems.Count
    For i = Lim To 1 Step -1
        If InStr(1, Slide1.Shapes("RegularApp:" & AppID).GroupItems(i).Name, "FileLabel") = 1 Or InStr(1, Slide1.Shapes("RegularApp:" & AppID).GroupItems(i).Name, "FileIcon") = 1 Then
            Slide1.Shapes("RegularApp:" & AppID).GroupItems(i).Delete
        End If
    Next i
End Sub

Sub Reload(AppID As String, Optional ByVal Attempt As Integer = 1)
    On Error GoTo Except
    Dim Shp As Shape
    Set Shp = Slide1.Shapes("ReloadAppFiles:" & AppID)
    Dim Sld As Slide
    Dim Ref As Shape
    Dim Dirname As String
    Set Sld = Slide1
    Set Ref = Sld.Shapes("RegularApp:" & AppID).GroupItems("WindowAppFiles:" & AppID)
    Dirname = Sld.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text
    Dim i As Integer
    CleanIcons AppID
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
            With Slide1.Shapes.AddTextbox(msoTextOrientationHorizontal, OffsetX - 10, OffsetY + Slide24.Shapes("FileIcon_/").Height, Slide24.Shapes("FileIcon_/").Width + 20, 20)
                .Name = "FileLabel" & CStr(IDX) & "AppFiles_"
                .TextFrame.TextRange.Text = File
                .TextFrame.TextRange.Font.Name = "Candara"
                .TextFrame.TextRange.Font.Size = 11
                .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
                If HideMe Then
                    .Visible = msoFalse
                End If
                If IsFolder Then
                    .ActionSettings(ppMouseClick).Run = "NavigateFolder"
                End If
            End With
            If IsFolder Then
                PasteToGroup Shp, Slide24.Shapes("FileIcon_/"), "FileIcon" & CStr(IDX) & "AppFiles:" & AppID, OffsetX + 10, OffsetY, Slide1, "NavigateIFolder"
                PasteToGroup Shp, Slide1.Shapes("FileLabel" & CStr(IDX) & "AppFiles_"), "FileLabel" & CStr(IDX) & "AppFiles:" & AppID, OffsetX, OffsetY + Slide24.Shapes("FileIcon_/").Height, Slide1, "NavigateFolder"
            Else
                Dim Assoc As String
                Assoc = GetAssoc(File, AppID)
                If Assoc = "" Then
                    PasteToGroup Shp, Slide24.Shapes("FileIcon_*"), "FileIcon" & CStr(IDX) & "AppFiles:" & AppID, OffsetX + 10, OffsetY, Slide1
                    PasteToGroup Shp, Slide1.Shapes("FileLabel" & CStr(IDX) & "AppFiles_"), "FileLabel" & CStr(IDX) & "AppFiles:" & AppID, OffsetX, OffsetY + Slide24.Shapes("FileIcon_/").Height, Slide1
                Else
                    Dim IconType As String
                    IconType = "*"
                    NameSplit = Split(File, ".")
                    NameExt = NameSplit(UBound(NameSplit))
                    FAssoc = FsAssoc(LCase(NameExt))
                    If FAssoc = "Notes" Then
                        IconType = "Txt"
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
                    End If
                    PasteToGroup Shp, Slide24.Shapes("FileIcon_" & IconType), "FileIcon" & CStr(IDX) & "AppFiles:" & AppID, OffsetX + 10, OffsetY, Slide1, "AssocI" & Assoc
                    PasteToGroup Shp, Slide1.Shapes("FileLabel" & CStr(IDX) & "AppFiles_"), "FileLabel" & CStr(IDX) & "AppFiles:" & AppID, OffsetX, OffsetY + Slide24.Shapes("FileIcon_/").Height, Slide1, "Assoc" & Assoc
                End If
            End If
            If HideMe Then
                Slide1.Shapes("FileIcon" & CStr(IDX) & "AppFiles:" & AppID).Visible = msoFalse
                Slide1.Shapes("FileLabel" & CStr(IDX) & "AppFiles:" & AppID).Visible = msoFalse
            End If
            Slide1.Shapes("FileLabel" & CStr(IDX) & "AppFiles_").Delete
            If IDX Mod 6 = 0 Then
                OffsetX = Ref.Left + 10
                OffsetY = OffsetY + Slide24.Shapes("FileIcon_/").Height + 30
            Else
                OffsetX = OffsetX + Slide24.Shapes("FileIcon_/").Width + 25
            End If
            If IDX Mod 18 = 0 Then
                OffsetX = Ref.Left + 10
                OffsetY = Ref.Top + 10
                HideMe = True
            End If
            IDX = IDX + 1
        End If
    Next File
    Slide1.Shapes("BottomPanelAppFiles:" & AppID).TextFrame.TextRange.Text = CStr(IDX - 1) & " items"
    Exit Sub
Except:
    Regroup AppID, Slide1
    Attempt = Attempt + 1
    If Attempt > 10 Then Exit Sub
    Pause 1
    Reload AppID, Attempt
    Exit Sub
End Sub


Sub FilesLoadDir(RShp As Shape)
    Dim AppID As String
    AppID = GetAppID(RShp)
    Reload AppID
End Sub

Function MaximumVisible(AppID As String) As Integer
    Dim MaxVis As Integer
    Dim Shp As Shape
    MaxVis = 0
    For Each Shp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
        If InStr(1, Shp.Name, "FileLabel") = 1 Then
            If Shp.Visible = msoTrue Then
                Dim SValue As Integer
                SValue = CInt(Replace(Replace(Shp.Name, "FileLabel", ""), "AppFiles:" & AppID, ""))
                If SValue > MaxVis Then
                    MaxVis = SValue
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
    For Each Shp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
        If InStr(1, Shp.Name, "FileLabel") = 1 Then
            If Shp.Visible = msoTrue Then
                Dim SValue As Integer
                SValue = CInt(Replace(Replace(Shp.Name, "FileLabel", ""), "AppFiles:" & AppID, ""))
                If SValue < MinVis Then
                    MinVis = SValue
                End If
            End If
        End If
    Next Shp
    MinimumVisible = MinVis
End Function


Sub FilesNextPage(Ref As Shape)
    On Error Resume Next
    Dim AppID As String
    Dim HideUpUntil As Integer
    Dim ShowFrom As Integer
    Dim ShowTo As Integer
    AppID = GetAppID(Ref)
    HideUpUntil = MaximumVisible(AppID)
    ShowFrom = HideUpUntil + 1
    ShowTo = ShowFrom + 17
    Dim Shp As Shape
    ' Hide visible entries
    Dim i As Integer
    For i = 1 To HideUpUntil
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(i) & "AppFiles:" & AppID).Visible = msoFalse
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(i) & "AppFiles:" & AppID).Visible = msoFalse
    Next i
    For i = ShowFrom To ShowTo
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(i) & "AppFiles:" & AppID).Visible = msoTrue
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(i) & "AppFiles:" & AppID).Visible = msoTrue
    Next i
    Exit Sub
End Sub

Sub FilesLastPage(Ref As Shape)
    On Error Resume Next
    Dim AppID As String
    Dim HideUpUntil As Integer
    Dim ShowFrom As Integer
    Dim ShowTo As Integer
    AppID = GetAppID(Ref)
    HideFrom = MinimumVisible(AppID)
    HideTo = HideFrom + 17
    ShowFrom = HideFrom - 18
    ShowTo = HideFrom - 1
    Dim Shp As Shape
    ' Hide visible entries
    Dim i As Integer
    For i = HideFrom To HideTo
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(i) & "AppFiles:" & AppID).Visible = msoFalse
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(i) & "AppFiles:" & AppID).Visible = msoFalse
    Next i
    For i = ShowFrom To ShowTo
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(i) & "AppFiles:" & AppID).Visible = msoTrue
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(i) & "AppFiles:" & AppID).Visible = msoTrue
    Next i
    Exit Sub
End Sub

Sub NavigateFolder(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    If InStr(1, Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text, "/") = 1 Then
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Shp.TextFrame.TextRange.Text
    Else
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppFiles:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Left(Shp.TextFrame.TextRange.Text, Len(Shp.TextFrame.TextRange.Text) - 1) & "\"
    End If
    Reload AppID
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