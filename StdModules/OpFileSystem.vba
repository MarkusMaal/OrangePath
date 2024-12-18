' ShapeFS

Function CountInodes() As Integer
    ' Subtract 28 to account for unrelated shapes in Slide2
    LogData "FS: Count inodes"
    CountInodes = Slide10.Shapes.Count + Slide6.Shapes.Count + Slide10.Shapes.Count + Slide9.Shapes.Count + Slide2.Shapes.Count - 28
End Function

Function FileStreamsExist(ByVal Filename As String) As Boolean
    Dim Exists As Boolean
    Exists = False
    LogData "FS: Check streams for " & Filename
    For Each Shp In Slide6.Shapes
        If InStr(1, Shp.Name, Filename) Then
            Exists = True
        End If
    Next Shp
    For Each Shp In Slide9.Shapes
        If InStr(1, Shp.Name, Filename) Then
            Exists = True
        End If
    Next Shp
    For Each Shp In Slide10.Shapes
        If InStr(1, Shp.Name, Filename) Then
            Exists = True
        End If
    Next Shp
    For Each Shp In Slide2.Shapes
        If InStr(1, Shp.Name, "App") = 1 Then
            Dim FakeName As String
            FakeName = "/Apps/" & Mid(Shp.Name, 4)
            If InStr(1, FakeName, Filename) Then
                Exists = True
            End If
        End If
    Next Shp
    FileStreamsExist = Exists
End Function
Function FileExists(ByVal Filename As String, Optional ByVal Stream As String = "") As Boolean
    Dim Exists As Boolean
    Exists = False
    LogData "FS: Check if file '" & Filename & "' exists"
    Dim Suffix As String
    Suffix = ""
    If Stream <> "" Then
        Suffix = "*" & Stream
    End If
    For Each Shp In Slide6.Shapes
        If Filename & Suffix = Shp.Name Then
            Exists = True
        End If
    Next Shp
    For Each Shp In Slide9.Shapes
        If Filename & Suffix = Shp.Name Then
            Exists = True
        End If
    Next Shp
    For Each Shp In Slide10.Shapes
        If Filename & Suffix = Shp.Name Then
            Exists = True
        End If
    Next Shp
    For Each Shp In Slide2.Shapes
        If InStr(1, Shp.Name, "App") = 1 Then
            Dim FakeName As String
            FakeName = "/Apps/" & Mid(Shp.Name, 4)
            If Filename & Suffix = FakeName Then
                Exists = True
            End If
        End If
    Next Shp
    FileExists = Exists
End Function

Sub NewFolder(ByVal Dirname As String)
    LogData "FS: Create directory " & Dirname
    If InStr(1, Dirname, "C:") <> 1 Then
        Dim Shp As Shape
        Dim Sld As Slide
        Depth = UBound(Split(Dirname, "/"))
        If Left(Dirname, 7) = "/Users/" Then
            Set Sld = Slide10
        ElseIf Left(Dirname, 10) = "/Defaults/" Then
            Set Sld = Slide6
        ElseIf Left(Dirname, 6) = "/Apps/" Then
            Set Sld = Slide2
        ElseIf Dirname = "/" Then
            Set Sld = Slide10
        Else
            Set Sld = Slide9
        End If
         If Left(Dirname, 7) = "/Users/" Then
            If Left(Dirname, 7 + Len(Slide1.Shapes("Username").TextFrame.TextRange.Text)) <> "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text Then
                AppMessage "Access is denied", "Cannot create directory", "Error", True
                Exit Sub
            Else
                With Sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 0, 0)
                    .Visible = msoFalse
                    .Name = Dirname & "/."
                End With
            End If
        ElseIf Left(Dirname, 6) = "/Apps/" Then
            AppMessage "Write access to apps directory is only available in single user mode", "Cannot create directory", "Error", True
            Exit Sub
        ElseIf Left(Dirname, 10) = "/Defaults/" Then
            AppMessage "Read only file system", "Cannot create directory", "Error", True
            Exit Sub
        Else
            With Sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 0, 0)
                .Visible = msoFalse
                .Name = Dirname & "/."
            End With
        End If
    Else
        AppMessage "For security reasons, local disk file operations are read only.", "Unable to write to C:", "Exclamation", True
    End If
    If InStr(1, Dirname, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop/") = 1 Then
        ShowDesktop
    End If
End Sub

Function LinkMovie(ByVal FilePath As String)
    LogData "FS: Linking movie file from host"
    If FileExists("/Temp/Movie.mov") Then
        DeleteFile "/Temp/Movie.mov"
    End If
    With Slide9.Shapes.AddMediaObject2(FilePath, msoTrue, msoFalse, 0, 0, 0, 0)
        .Name = "/Temp/Movie.mov"
        .Visible = msoFalse
        .LockAspectRatio = msoFalse
    End With
End Function

Function ImportMovie(ByVal FilePath As String, ByVal OP_path As String)
    LogData "FS: Importing movie from host to " & OP_path
    If FileExists(OP_path) Then
        DeleteFile OP_path
    End If
    Dim Sld As Slide
    If Left(OP_path, 7) = "/Users/" Then
        If InStr(1, OP_path, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/") <> 1 Then
            AppMessage "Cannot write to this directory", "Access denied", "Error", True
            Exit Function
        End If
        Set Sld = Slide10
    ElseIf Left(OP_path, 10) = "/Defaults/" Then
        Set Sld = Slide6
    Else
        Set Sld = Slide9
    End If
    If InStr(1, OP_path, "/Defaults/") = 1 Then
        AppMessage "Read only file system", "Access denied", "Error", True
        Exit Function
    ElseIf InStr(1, OP_path, "/Apps/") = 1 Then
        AppMessage "Write access to apps directory is only available in single user mode", "Access denied", "Error", True
        Exit Function
    End If
    With Sld.Shapes.AddMediaObject2(FilePath, msoFalse, msoTrue, 0, 0, 0, 0)
        .Name = OP_path
        .Visible = msoFalse
        .LockAspectRatio = msoFalse
    End With
    If InStr(1, OP_path, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop/") = 1 Then
        ShowDesktop
    End If
End Function

Function GetFiles(ByVal Dirname As String) As String
    On Error Resume Next
    If InStr(1, Dirname, "C:") <> 1 Then
        If Right(Dirname, 1) <> "/" Then
            Dirname = Dirname & "/"
        End If
        LogData "FS: Populating file list for " & Dirname & " (HostFS)"
        GetFiles = ""
        Dim Shp As Shape
        Dim Sld As Slide
        Depth = UBound(Split(Dirname, "/"))
        If Left(Dirname, 7) = "/Users/" Then
            Set Sld = Slide10
        ElseIf Left(Dirname, 10) = "/Defaults/" Then
            Set Sld = Slide6
        ElseIf Left(Dirname, 6) = "/Apps/" Then
            Set Sld = Slide2
            Dim I As Integer
            For I = Slide2.Shapes.Count To 1 Step -1
                Set Shp = Slide2.Shapes(I)
                If InStr(1, Shp.Name, "App") = 1 Then
                    Dim FakeName As String
                    FakeName = Mid(Shp.Name, 4)
                    GetFiles = GetFiles & FakeName & ".app" & vbNewLine
                End If
            Next I
            GoTo Done
        ElseIf Dirname = "/" Then
            Set Sld = Slide10
            GetFiles = "System/" & vbNewLine
            GetFiles = GetFiles & "Defaults/" & vbNewLine
            GetFiles = GetFiles & "Temp/" & vbNewLine
            GetFiles = GetFiles & "Apps/" & vbNewLine
        Else
            Set Sld = Slide9
        End If
        For Each Shp In Sld.Shapes
            If Left(Shp.Name, Len(Dirname)) = Dirname Then
                SplitName = Split(Shp.Name, "/")
                If UBound(SplitName) = Depth Then
                    If InStr(GetFiles, GetFakeName(Shp.Name)) Then GoTo Continue
                    If Right(GetFakeName(Shp.Name), Len(GetFakeName(Shp.Name)) - Len(Dirname)) <> "." Then
                        GetFiles = GetFiles & Right(GetFakeName(Shp.Name), Len(GetFakeName(Shp.Name)) - Len(Dirname)) & vbNewLine
                    End If
                Else
                    Dim out As String
                    out = "/"
                    For I = 1 To Depth Step 1
                        out = out & SplitName(I) & "/"
                    Next I
                    out = Right(out, Len(out) - Len(Dirname))
                    If InStr(GetFiles, out) Then GoTo Continue
                    GetFiles = GetFiles & out & vbNewLine
                End If
            End If
Continue:
        Next Shp
        GoTo Done
    Else
        Dim Files As String
        Files = ""
        LogData "FS: Populating file list for " & Dirname & " (ShapeFS)"

        
        'Variable Declaration
        Dim sFilePath As String
        Dim sFileName As String
        
        ' Specify File Path
        sFilePath = Dirname
        
        ' Add subfolders
        Dim F As String
        Set GetFoldersIn = New Collection
        F = Dir(Dirname & "\*", vbDirectory)
        Do While F <> ""
          If F <> "." And F <> ".." Then
            If GetAttr(Dirname & "\" & F) And vbDirectory Then Files = Files & F & "/" & vbNewLine
          End If
          F = Dir
        Loop
        
        'Check for back slash
        If Right(sFilePath, 1) <> "\" Then
            sFilePath = sFilePath & "\"
        End If
            
        sFileName = Dir(sFilePath)
        
        Do While Len(sFileName) > 0
            
            'Display file name in immediate window
            Files = Files & sFileName & vbNewLine
            
            'Set the fileName to the next available file
            sFileName = Dir
        Loop


        GetFiles = Files
        GoTo Done
    End If
Crash:
    OSCrash "OP_FILE_SYSTEM", Err
Done:
End Function

Function GetSubFolders(RootPath As String)
    Dim fso As Object
    Dim fld As Object
    Dim arr As Variant
    Dim sf As Object
    Dim myArr
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(RootPath)
    For Each sf In fld.SubFolders
        ReDim Preserve arr(counter)
        arr(counter) = sf.Path
        counter = counter + 1
        myArr = GetSubFolders(sf.Path)
    Next
    GetSubFolders = arr
    Set sf = Nothing
    Set fld = Nothing
    Set fso = Nothing
End Function


Sub TestList()
    Folder = "C:\"
    Dim F As String
    Set GetFoldersIn = New Collection
    F = Dir(Folder & "\*", vbDirectory)
    Do While F <> ""
      If GetAttr(Folder & "\" & F) And vbDirectory Then MsgBox F
      F = Dir
    Loop
End Sub
Function GetFakeName(ByVal RealName As String) As String
    GetFakeName = RealName
    Exit Function
    If InStr(RealName, "*") Then
        SplitName = Split(RealName, "*")
        GetFakeName = SplitName(0)
    Else
        GetFakeName = RealName
    End If
End Function

Function GetFileRef(ByVal Filename As String, Optional ByVal Stream As String = "") As Shape
    'If InStr(1, Filename, "/Defaults/") = 1 And FileStreamsExist(Replace(Filename, "/Defaults/", "/System/")) Then
    '    Filename = Replace(Filename, "/Defaults/", "/System/")
    'End If
    On Error GoTo CrashRef
    LogData "FS: Reading file '" & Filename & "' (ShapeFS)"
    Dim Suffix As String
    Suffix = "*" & Stream
    If Suffix = "*" Then
        Suffix = ""
    End If
    If Left(Filename, 7) = "/Users/" Then
        If Left(Filename, 7 + Len(Slide1.Shapes("Username").TextFrame.TextRange.Text)) = "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text Then
            Set GetFileRef = Slide10.Shapes(Filename & Suffix)
        End If
    ElseIf Left(Filename, 10) = "/Defaults/" Then
        Set GetFileRef = Slide6.Shapes(Filename & Suffix)
    ElseIf Left(Filename, 6) = "/Apps/" Then
        Dim ActualName As String
        ActualName = Replace(Filename, "/Apps/", "")
        Set GetFileRef = Slide2.Shapes("App" & Left(ActualName, Len(ActualName) - 4) & Suffix)
    Else
        Set GetFileRef = Slide9.Shapes(Filename & Suffix)
    End If
    Exit Function
CrashRef:
    'AppMessage "File does not exist: " + Filename, "Filesystem error", "Error", False
    Set GetFileRef = Nothing
    Exit Function
End Function

Function GetFileContent(ByVal Filename As String, Optional ByVal Stream As String = "") As String
    If InStr(1, Filename, "/Defaults/") = 1 And FileStreamsExist(Replace(Filename, "/Defaults/", "/System/")) Then
        Filename = Replace(Filename, "/Defaults/", "/System/")
    End If
    If InStr(1, Filename, "C:\") <> 1 Then
        LogData "FS: Reading text content for file '" & Filename & "' (ShapeFS)"
        Dim Shp As Shape
        Set Shp = GetFileRef(Filename, Stream)
        If Not Shp Is Nothing Then
            GetFileContent = Shp.TextFrame.TextRange.Text
        Else
            GetFileContent = "*"
        End If
    Else
        LogData "FS: Reading text content for file '" & Filename & "' (HostFS)"
        Dim fS, F
        Const ForReading = 1, ForWriting = 2, ForAppending = 8
        Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
        Set fS = CreateObject("Scripting.FileSystemObject")
        Set F = fS.OpenTextFile(Filename, ForReading, True, TristateFalse)
        GetFileContent = F.ReadAll
        F.Close
    End If
End Function

Sub SetFileContent(ByVal Filename As String, ByVal Content As String, Optional ByVal Stream As String = "")
    LogData "FS: Attempting to save text file as '" & Filename & "'"
    If InStr(1, Filename, "C:\") <> 1 Then
        ' Set correct slide number
        Dim Sld As Slide
        If Left(Filename, 7) = "/Users/" Then
            Set Sld = Slide10
        ElseIf Left(Filename, 10) = "/Defaults/" Then
            Set Sld = Slide6
        Else
            Set Sld = Slide9
        End If
        If InStr(1, Filename, "/Defaults/") = 1 Then
            AppMessage "Read only file system", "Access denied", "Info", True
            Exit Sub
        End If
        If InStr(1, Filename, "/Apps/") = 1 Then
            AppMessage "Write access is only available in single user mode", "Access denied", "Info", True
            Exit Sub
        End If
        ' Setup file stream pointer if specified
        Dim Suffix As String
        Suffix = "*" & Stream
        If Suffix = "*" Then
            Suffix = ""
        End If
        ' Delete shape if file exists
        If FileExists(Filename, Stream) Then
            Sld.Shapes(Filename & Suffix).Delete
        End If
        ' Create shape with file content
        With Sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 0, 0)
            .TextFrame.TextRange.Text = Content
            .Name = Filename & Suffix
            .Visible = msoFalse
        End With
    Else
        AppMessage "For security reasons, local disk file operations are read only.", "Write error", "Exclamation", True
    End If
    If InStr(1, Filename, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop/") = 1 Then
        ShowDesktop
    End If
End Sub

Sub DeleteFile(ByVal Filename As String, Optional ByVal Stream As String = "")
    LogData "FS: Attempting to delete file at '" & Filename & "'"
    If InStr(1, Filename, "C:\") <> 1 Then
        ' Set correct slide number
        Dim Sld As Slide
        If Left(Filename, 7) = "/Users/" Then
            If InStr(1, Filename, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/") <> 1 And Slide1.Shapes("Username").TextFrame.TextRange.Text <> "Nobody" Then
                AppMessage "You have insufficient permissions to delete this file", "Access denied", "Info", True
                Exit Sub
            End If
            Set Sld = Slide10
        ElseIf Left(Filename, 10) = "/Defaults/" Then
            Set Sld = Slide6
        ElseIf Left(Filename, 10) = "/Apps/" Then
            Set Sld = Slide2
        Else
            Set Sld = Slide9
        End If
        If InStr(1, Filename, "/Defaults/") = 1 Then
            AppMessage "Read only file system", "Access denied", "Info", True
            Exit Sub
        ElseIf InStr(1, Filename, "/Apps/") = 1 Then
            AppMessage "App deletion from ShapeFS has not been implemented.", "Notice", "Info", True
            Exit Sub
        End If
        ' Setup file stream pointer if specified
        Dim Suffix As String
        Suffix = "*" & Stream
        If Suffix = "*" Then
            Suffix = ""
        End If
        ' Delete shape if file exists
        If FileExists(Filename, Stream) Then
            Sld.Shapes(Filename & Suffix).Delete
        End If
    Else
        AppMessage "For security reasons, local disk file operations are read only.", "Cannot delete from C:", "Exclamation", True
    End If
    If InStr(1, Filename, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop/") = 1 Then
        ShowDesktop
    End If
End Sub

Sub DeleteDir(ByVal Dirname As String)
    LogData "FS: Attempting to delete directory at '" & Dirname & "'"
    ' Set correct slide number
    Dim Sld As Slide
    If Left(Dirname, 7) = "/Users/" Then
        Set Sld = Slide10
        If InStr(1, Dirname, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/") <> 1 And Slide1.Shapes("Username").TextFrame.TextRange.Text <> "Nobody" Then
            AppMessage "You have insufficient permissions to delete this directory", "Access denied", "Error", True
            Exit Sub
        End If
    ElseIf Left(Dirname, 10) = "/Defaults/" Then
        Set Sld = Slide6
    ElseIf Left(Dirname, 10) = "/Apps/" Then
        Set Sld = Slide6
    Else
        Set Sld = Slide9
    End If
    If InStr(1, Dirname, "/Defaults/") = 1 Then
        AppMessage "Read only file system", "Access denied", "Error", True
        Exit Sub
    ElseIf InStr(1, Dirname, "/Apps/") = 1 Then
        AppMessage "App deletion from ShapeFS has not been implemented.", "Notice", "Info", True
        Exit Sub
    End If
    Dim Shp As Shape
    ' Delete shapes
    For I = Sld.Shapes.Count To 1 Step -1
        If Left(Sld.Shapes(I).Name, Len(Dirname)) = Dirname Then
            Sld.Shapes(I).Delete
        End If
    Next I
    If InStr(1, Dirname, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop/") = 1 Then
        ShowDesktop
    End If
End Sub


' For copying files (duh)
Sub CopyFile(ByVal Source As String, ByVal Destination As String, Optional ByVal EnforceDestinationName As Boolean = False)
    LogData "FS: Attempting to copy file from '" & Source & "' to '" & Destination & "'"
    ' Set correct slide number
    Dim DstSld As Slide
    If Left(Destination, 7) = "/Users/" Then
        Set DstSld = Slide10
        If InStr(1, Destination, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/") <> 1 Then
            AppMessage "Access is denied", "Copy error", "Error", True
        End If
    ElseIf Left(Destination, 10) = "/Defaults/" Then
        AppMessage "Read-only file system", "Copy error", "Error", True
        Exit Sub
    ElseIf Left(Destination, 6) = "/Apps/" Then
        AppMessage "Write access to apps directory is not available in multi-user mode", "Copy error", "Error", True
        Exit Sub
    Else
        Set DstSld = Slide9
    End If
    ' Local copy
    If InStr(1, Source, "C:\") <> 1 Then
        Dim SrcFile As Shape
        SourceSplit = Split(Source, "/")
        Dim SafeSourceName As String
        SafeSourceName = SourceSplit(UBound(SourceSplit))
        Set SrcFile = GetFileRef(Source)
        SrcFile.Copy
        With DstSld.Shapes.Paste
            If Not EnforceDestinationName Then
                .Name = Destination & SafeSourceName
            Else
                .Name = Destination
            End If
        End With
    ' Import from C:
    Else
        SourceSplit2 = Split(Source, ".")
        Dim Ext As String
        Ext = SourceSplit2(UBound(SourceSplit2))
        SourceSplit2 = Split(Source, "\")
        Dim SafeName As String
        SafeName = SourceSplit2(UBound(SourceSplit2))
        Dim Assoc As String
        Assoc = FsAssoc(LCase(Ext))
        If Assoc = "Notes" Then
            Dim TextData As String
            TextData = GetFileContent(Source)
            With DstSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 0, 0)
                .Name = Destination & SafeName
                .TextFrame.TextRange.Text = TextData
                .Visible = msoFalse
            End With
        ElseIf Assoc = "VideoPlayer" Then
            ImportMovie Source, Destination & SafeName
        ElseIf Assoc = "SoundPlayer" Then
            ImportMovie Source, Destination & SafeName
        ElseIf Assoc = "PictureView" Then
            SetFilePic Destination & SafeName, Source
        Else
            AppMessage "This type of file cannot be imported", "Import file", "Info", True
        End If
    End If
    If InStr(1, Destination, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop/") = 1 Then
        ShowDesktop
    End If
End Sub

Function FsAssoc(ByVal Extension As String) As String
    LogData "FS: Getting file association for extension " & Extension
    If Extension = "pres" Then
        FsAssoc = "Presentator"
    ElseIf Extension = "txt" Then
        FsAssoc = "Notes"
    ElseIf Extension = "3d" Then
        FsAssoc = "3D"
    ElseIf Extension = "mp4" Then
        FsAssoc = "VideoPlayer"
    ElseIf Extension = "mov" Then
        FsAssoc = "VideoPlayer"
    ElseIf Extension = "mkv" Then
        FsAssoc = "VideoPlayer"
    ElseIf Extension = "mpg" Then
        FsAssoc = "VideoPlayer"
    ElseIf Extension = "avi" Then
        FsAssoc = "VideoPlayer"
    ElseIf Extension = "wmv" Then
        FsAssoc = "VideoPlayer"
    ElseIf Extension = "webm" Then
        FsAssoc = "VideoPlayer"
    ElseIf Extension = "jpg" Then
        FsAssoc = "PictureView"
    ElseIf Extension = "jpeg" Then
        FsAssoc = "PictureView"
    ElseIf Extension = "jpe" Then
        FsAssoc = "PictureView"
    ElseIf Extension = "jfif" Then
        FsAssoc = "PictureView"
    ElseIf Extension = "jfi" Then
        FsAssoc = "PictureView"
    ElseIf Extension = "png" Then
        FsAssoc = "PictureView"
    ElseIf Extension = "bmp" Then
        FsAssoc = "PictureView"
    ElseIf Extension = "gif" Then
        FsAssoc = "PictureView"
    ElseIf Extension = "pic" Then
        FsAssoc = "PictureView"
    ElseIf Extension = "emf" Then
        FsAssoc = "PictureView"
    ElseIf Extension = "ini" Then
        FsAssoc = "Notes"
    ElseIf Extension = "inf" Then
        FsAssoc = "Notes"
    ElseIf Extension = "log" Then
        FsAssoc = "Notes"
    ElseIf Extension = "json" Then
        FsAssoc = "Notes"
    ElseIf Extension = "bat" Then
        FsAssoc = "Notes"
    ElseIf Extension = "cmd" Then
        FsAssoc = "Notes"
    ElseIf Extension = "cnf" Then
        FsAssoc = "Settings"
    ElseIf Extension = "wav" Then
        FsAssoc = "SoundPlayer"
    ElseIf Extension = "ogg" Then
        FsAssoc = "SoundPlayer"
    ElseIf Extension = "mp3" Then
        FsAssoc = "SoundPlayer"
    ElseIf Extension = "wma" Then
        FsAssoc = "SoundPlayer"
    ElseIf Extension = "pxl" Then
        FsAssoc = "Paint"
    ElseIf Extension = "app" Then
        FsAssoc = "Special"
    ElseIf Extension = "grp" Then
        FsAssoc = "Components"
    ElseIf Extension = "hlp" Then
        FsAssoc = "ModalHelpView"
    Else
        If GetSysConfig("assoc" & Extension) <> "*" Then
            FsAssoc = GetSysConfig("assoc" & Extension)
            Exit Function
        End If
        FsAssoc = ""
    End If
End Function

Sub RenameFile(ByVal Filename As String, ByVal Newname As String, Optional ByVal Stream As String = "")
    LogData "FS: Renaming file '" & Filename & "' to '" & Newname & "'"
    ' Set correct slide number
    Dim Sld As Slide
    If Left(Filename, 7) = "/Users/" Then
        Set Sld = Slide10
        If InStr(1, Filename, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text) <> 1 Then
            AppMessage "You have insufficient permissions to rename this file", "Access denied", "Error", True
            Exit Sub
        End If
    ElseIf Left(Filename, 10) = "/Defaults/" Then
        AppMessage "Read-only file system", "You know this isn't going to work, right?", "Error", True
        Exit Sub
    ElseIf Left(Filename, 10) = "/Apps/" Then
        AppMessage "Apps cannot be renamed in this implementation of ShapeFS", "Rename error", "Error", True
        Exit Sub
    Else
        Set Sld = Slide9
    End If
    ' Setup file stream pointer if specified
    Dim Suffix As String
    Suffix = "*" & Stream
    If Suffix = "*" Then
        Suffix = ""
    End If
    ' Rename shape
    With Sld.Shapes(Filename & Suffix)
        .Name = Newname & Stream
    End With
    If InStr(1, Filename, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop/") = 1 Then
        ShowDesktop
    End If
End Sub

Sub PreparePic(ByVal Filename As String, Optional ByVal Stream As String = "")
    LogData "FS: Temporarily storing image file on host filesystem ('" & Filename & "')"
    ' Set correct slide number
    Dim Sld As Slide
    If Left(Filename, 7) = "/Users/" Then
        Set Sld = Slide10
    ElseIf Left(Filename, 10) = "/Defaults/" Then
        Set Sld = Slide6
    ElseIf Left(Filename, 6) = "/Apps/" Then
        Set Sld = Slide2
    Else
        Set Sld = Slide9
    End If
    ' Setup file stream pointer if specified
    Dim Suffix As String
    Suffix = "*" & Stream
    If Suffix = "*" Then
        Suffix = ""
    End If
    ' Save image to tempfile
    Dim Shp As Shape
    Set Shp = Sld.Shapes(Filename & Suffix)
    Shp.Visible = msoTrue
    Shp.Left = 0
    Shp.Top = 0
    Shp.Width = ActivePresentation.PageSetup.SlideWidth
    Shp.Height = ActivePresentation.PageSetup.SlideHeight
    'Shp.ZOrder msoBringToFront
    Sld.Export Environ("TEMP") & "\Userpic.PNG", "PNG"
    Shp.Visible = msoFalse
    Shp.Width = 0
    Shp.Height = 0
End Sub

Sub SetFilePic(ByVal Filename As String, ByVal Tempfile As String, Optional ByVal Stream As String = "")
    LogData "FS: Importing picture from host ('" & Tempfile & "') to ShapeFS ('" & Filename & "')"
    ' Delete file if it exists
    If FileExists(Filename) Then DeleteFile (Filename)
    ' Set correct slide number
    Dim Sld As Slide
    If Left(Filename, 7) = "/Users/" Then
        Set Sld = Slide10
    ElseIf Left(Filename, 10) = "/Defaults/" Then
        Set Sld = Slide6
    ElseIf Left(Filename, 6) = "/Apps/" Then
        Set Sld = Slide2
    Else
        Set Sld = Slide9
    End If
    ' Setup file stream pointer if specified
    Dim Suffix As String
    Suffix = "*" & Stream
    If Suffix = "*" Then
        Suffix = ""
    End If
    ' Create default picture file
    With Sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 0, 0)
        .Name = Filename & Suffix
        .Fill.UserPicture Tempfile
        .Visible = msoFalse
    End With
    If InStr(1, Filename, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop/") = 1 Then
        ShowDesktop
    End If
End Sub

Function WriteGroup(ByVal Filename As String, ByVal SShp As ShapeRange, ByVal SizeKey As String)
    LogData "FS: Saving group as '" & Filename & "'"
    ' Delete file if it exists
    If FileExists(Filename) Then DeleteFile (Filename)
    ' Set correct slide number
    Dim Sld As Slide
    If Left(Filename, 7) = "/Users/" Then
        Set Sld = Slide10
    ElseIf Left(Filename, 10) = "/Defaults/" Then
        Set Sld = Slide6
    ElseIf Left(Filename, 6) = "/Apps/" Then
        Set Sld = Slide2
    Else
        Set Sld = Slide9
    End If
    If InStr(1, Filename, "/Defaults/") = 1 Then
        AppMessage "Access denied", "Read only file system", "Error", True
        Exit Function
    ElseIf InStr(1, Filename, "/Apps/") = 1 Then
        AppMessage "Write access to apps directory is not available in multi-user mode", "Protected file system", "Error", True
        Exit Function
    End If
    If InStr(1, Filename, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/") <> 1 Then
        AppMessage "Access denied", "Unable to write group", "Error", True
        Exit Function
    End If
    ' Copy group to filesystem slide
    SShp.Copy
    With Sld.Shapes.Paste.Group
        .Name = Filename
        .Visible = msoFalse
        Dim GI As Shape
        For Each GI In .GroupItems
            If GI.Name <> SizeKey Then
                Dim Oldname() As String
                Dim Newname As String
                Oldname = Split(GI.Name, ":")
                Newname = Oldname(0) & "_"
                GI.Name = Newname
                GI.Visible = msoTrue
            Else
                ' Size key makes sure that shape proportions remain the same, when loading a group file
                GI.Name = "SizeKey"
                GI.Visible = msoFalse
            End If
        Next GI
    End With
    If InStr(1, Filename, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop/") = 1 Then
        ShowDesktop
    End If
End Function

Function ReadGroup(ByVal Filename As String, ByVal TSld As Slide, ByVal OffsetX As Integer, ByVal OffsetY As Integer, ByVal Ref As Shape, ByVal sizeX As Integer, ByVal sizeY As Integer)
    If InStr(1, Filename, "/Defaults/") = 1 And FileStreamsExist(Replace(Filename, "/Defaults/", "/System/")) Then
        Filename = Replace(Filename, "/Defaults/", "/System/")
    End If
    If FileExists(Filename) Then
        ' Set correct slide number
        Dim Sld As Slide
        Dim AppID As String
        AppID = GetAppID(Ref)
        LogData "FS: Reading group at '" & Filename & "' and pasting to window with ID " & AppID
        If Left(Filename, 7) = "/Users/" Then
            Set Sld = Slide10
        ElseIf Left(Filename, 10) = "/Defaults/" Then
            Set Sld = Slide6
        ElseIf Left(Filename, 6) = "/Apps/" Then
            Set Sld = Slide2
            Dim ActualName As String
            ActualName = Replace(Filename, "/Apps/", "")
            Filename = "App" & Left(ActualName, Len(ActualName) - 4)
        Else
            Set Sld = Slide9
        End If
        If InStr(1, Filename, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/") <> 1 And InStr(1, Filename, "/Users/") = 1 Then
            AppMessage "Access denied", "Unable to open group", "Error", True
            Exit Function
        End If
        
        Sld.Shapes(Filename).Left = OffsetX
        Sld.Shapes(Filename).Top = OffsetY
        Sld.Shapes(Filename).Width = sizeX
        Sld.Shapes(Filename).Height = sizeY
        Dim GI As Shape
        For Each GI In Sld.Shapes(Filename).GroupItems
            ' This prevents the size key from ever being pasted
            If GI.Name <> "SizeKey" Then
                PasteToGroup Ref, GI, Replace(GI.Name, "_", ":" & AppID), GI.Left, GI.Top, TSld
            End If
        Next GI
    End If
End Function

Sub AssocSpecial(Shp As Shape)
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    Dim ActualName As String
    ActualName = Replace(Filename, "/Apps/", "")
    ActualName = Left(ActualName, Len(ActualName) - 4)
    ' Attempt 1
    On Error GoTo Att2
    Slide1.Shapes("RegularApp:" & AppID).Copy
    With Slide1.Shapes.Paste
        .Name = "DummyApp:" & AppID
        .Visible = msoFalse
    End With
    Application.Run "App" & ActualName, Slide1.Shapes("DummyApp:" & AppID).GroupItems(1)
    On Error Resume Next
    Slide1.Shapes("DummyApp:" & AppID).Delete
    Exit Sub
Att2:
    Slide1.Shapes("DummyApp:" & AppID).Delete
    Application.Run "App" & ActualName
    Exit Sub
End Sub

Sub AssocISpecial(Shp As Shape)
    On Error GoTo ReportIssue2
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocSpecial Slide1.Shapes(ShapeName)
    Exit Sub
ReportIssue2:
    AppMessage Err.Description, Err.Source, "Error", True
End Sub