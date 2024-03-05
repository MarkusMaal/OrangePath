' Modal files app
Sub AppModalFiles()
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "ModalFiles"
    LaunchPath = "/"
    If CheckVars("%LaunchDir%") <> "" And CheckVars("%LaunchDir%") <> "%LaunchDir%" Then
        LaunchPath = CheckVars("%LaunchDir%")
    End If
    Slide2.Shapes("PathAppModalFiles_").TextFrame.TextRange.Text = LaunchPath
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    If CheckVars("%Save%") = "" Or CheckVars("%Save%") = "%Save%" Then
        Slide1.Shapes("OkAppModalFiles:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).Delete
        Slide1.Shapes("AxTextBox1AppModalFiles:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).Delete
        Slide1.AxTextBox.Visible = False
    End If
    MReload Slide1.Shapes("AppID").TextFrame.TextRange.Text
End Sub

Function MFileCount(ByVal AppID As String) As Integer
    MFileCount = CInt(Replace(Slide1.Shapes("RegularApp:" & AppID).GroupItems("BottomPanelAppModalFiles:" & AppID).TextFrame.TextRange.Text, " items", ""))
End Function

Sub AssocModal(Shp As shape)
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
        Slide1.AxTextBox.Text = Shp.TextFrame.TextRange.Text
        Slide1.Shapes("AxTextBox1AppModalFiles:" & AppID).TextFrame.TextRange.Text = Shp.TextFrame.TextRange.Text
    End If
    Exit Sub
InvalidFile:
    AppMessage "Cannot open files of this type", "Error", "Error", True
End Sub

Sub SaveFile(Shp As shape)
    AppID = GetAppID(Shp)
    SetVar "InputValue", Slide1.Shapes("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes("AxTextBox1AppModalFiles:" & AppID).TextFrame.TextRange.Text
    If CheckVars("%Macro%") <> "" And CheckVars("%Macro%") <> "%Macro%" Then
        Application.Run CheckVars("%Macro%"), Shp
    End If
    UnsetVar "Macro"
    UnsetVar "Save"
    CloseWindow Shp
End Sub

Sub AssocIModal(Shp As shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocModal Slide1.Shapes(ShapeName)
End Sub

Sub MReload(AppID As String, Optional ByVal Attempt As Integer = 1)
    On Error GoTo Except
    Dim Shp As shape
    Set Shp = Slide1.Shapes("ReloadAppModalFiles:" & AppID)
    Dim Sld As slide
    Dim Ref As shape
    Dim Dirname As String
    Set Sld = Slide1
    Set Ref = Sld.Shapes("RegularApp:" & AppID).GroupItems("WindowAppModalFiles:" & AppID)
    Dirname = Sld.Shapes("RegularApp:" & AppID).GroupItems("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text
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
                PasteToGroup Shp, Slide24.Shapes("FileIcon_/"), "FileIcon" & CStr(IDX) & "AppModalFiles:" & AppID, OffsetX + 10, OffsetY, Slide1, "MNavigateIFolder"
                PasteToGroup Shp, Slide1.Shapes("FileLabel" & CStr(IDX) & "AppModalFiles_"), "FileLabel" & CStr(IDX) & "AppModalFiles:" & AppID, OffsetX, OffsetY + Slide24.Shapes("FileIcon_/").Height, Slide1, "MNavigateFolder"
            Else
                Dim IconType As String
                IconType = "*"
                NameSplit = Split(File, ".")
                NameExt = NameSplit(UBound(NameSplit))
                If FAssoc = "Notes" Then
                    IconType = "Txt"
                ElseIf FAssoc = "Paint" Then
                    IconType = "Pic"
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
                PasteToGroup Shp, Slide24.Shapes("FileIcon_" & IconType), "FileIcon" & CStr(IDX) & "AppModalFiles:" & AppID, OffsetX + 10, OffsetY, Slide1, "AssocIModal"
                PasteToGroup Shp, Slide1.Shapes("FileLabel" & CStr(IDX) & "AppModalFiles_"), "FileLabel" & CStr(IDX) & "AppModalFiles:" & AppID, OffsetX, OffsetY + Slide24.Shapes("FileIcon_/").Height, Slide1, "AssocModal"
            End If
            If HideMe Then
                Slide1.Shapes("FileIcon" & CStr(IDX) & "AppModalFiles:" & AppID).Visible = msoFalse
                Slide1.Shapes("FileLabel" & CStr(IDX) & "AppModalFiles:" & AppID).Visible = msoFalse
            End If
            Slide1.Shapes("FileLabel" & CStr(IDX) & "AppModalFiles_").Delete
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
    Slide1.Shapes("BottomPanelAppModalFiles:" & AppID).TextFrame.TextRange.Text = CStr(IDX - 1) & " items"
    Exit Sub
Except:
    Regroup AppID, Slide1
    Attempt = Attempt + 1
    If Attempt > 10 Then Exit Sub
    Pause 1
    MReload AppID, Attempt
    Exit Sub
End Sub


Sub MFilesLoadDir(RShp As shape)
    Dim AppID As String
    AppID = GetAppID(RShp)
    MReload AppID
End Sub

Function MMaximumVisible(AppID As String) As Integer
    Dim MaxVis As Integer
    Dim Shp As shape
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
    Dim Shp As shape
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


Sub MFilesNextPage(Ref As shape)
    Dim AppID As String
    Dim HideUpUntil As Integer
    Dim ShowFrom As Integer
    Dim ShowTo As Integer
    AppID = GetAppID(Ref)
    HideUpUntil = MMaximumVisible(AppID)
    ShowFrom = HideUpUntil + 1
    ShowTo = ShowFrom + 17
    Dim Shp As shape
    ' Hide visible entries
    Dim i As Integer
    For i = 1 To HideUpUntil
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(i) & "AppModalFiles:" & AppID).Visible = msoFalse
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(i) & "AppModalFiles:" & AppID).Visible = msoFalse
    Next i
    For i = ShowFrom To ShowTo
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(i) & "AppModalFiles:" & AppID).Visible = msoTrue
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(i) & "AppModalFiles:" & AppID).Visible = msoTrue
    Next i
    Exit Sub
End Sub

Sub MFilesLastPage(Ref As shape)
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
    Dim Shp As shape
    ' Hide visible entries
    Dim i As Integer
    For i = HideFrom To HideTo
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(i) & "AppModalFiles:" & AppID).Visible = msoFalse
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(i) & "AppModalFiles:" & AppID).Visible = msoFalse
    Next i
    For i = ShowFrom To ShowTo
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileLabel" & CStr(i) & "AppModalFiles:" & AppID).Visible = msoTrue
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("FileIcon" & CStr(i) & "AppModalFiles:" & AppID).Visible = msoTrue
    Next i
    Exit Sub
End Sub

Sub MNavigateFolder(Shp As shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    If InStr(1, Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text, "/") = 1 Then
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text & Shp.TextFrame.TextRange.Text
    Else
        Slide1.Shapes("RegularApp:" & AppID).GroupItems("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("PathAppModalFiles:" & AppID).TextFrame.TextRange.Text & Left(Shp.TextFrame.TextRange.Text, Len(Shp.TextFrame.TextRange.Text) - 1) & "\"
    End If
    MReload AppID
End Sub

Sub MNavigateIFolder(Shp As shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim LabelName As String
    LabelName = Replace(Shp.Name, "Icon", "Label")
    MNavigateFolder Slide1.Shapes(LabelName)
End Sub

Sub MFilesUp(Shp As shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    MGoUp AppID
End Sub

Sub MGoUp(AppID As String)
    Dim Sld As slide
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