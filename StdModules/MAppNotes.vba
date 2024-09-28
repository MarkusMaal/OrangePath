' Notes app

Sub AppNotes(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Notes"
    
    Slide2.Shapes("Shape11AppGuess_").TextFrame.TextRange.Text = CStr(Int(100 * Rnd))
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
End Sub


Sub AssocNotes(Shp As Shape)
    ' Get full file path from shape
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    ' Load text content
    TextData = GetFileContent(Filename)
    If TextData = "*" Then
        AppMessage "Access to the file is denied", "File access", "Error", True
        Exit Sub
    End If
    ' Launch notes app
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Notes"
    Slide2.Shapes("Shape11AppGuess_").TextFrame.TextRange.Text = CStr(Int(100 * Rnd))
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    ' Get AppID of newly created window
    AppID = Slide1.Shapes("AppID").TextFrame.TextRange.Text
    ' Set taskbar and window titlebar titles
    If Len(Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text) > 13 Then
        Slide1.Shapes("TaskIcon:" & AppID).TextFrame.TextRange.Text = "..." & Right(Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text, 13)
    Else
        Slide1.Shapes("TaskIcon:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    End If
    Slide1.Shapes("WindowTitleAppNotes:" & AppID).TextFrame.TextRange.Text = Filename
    ' Display file content
    Slide1.Shapes("WindowAppNotes:" & AppID).TextFrame.TextRange.Text = TextData
    Slide1.AxTextBox.Text = TextData
End Sub

Sub AssocINotes(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocNotes Slide1.Shapes(ShapeName)
End Sub

Sub SetNotesBg(Shp As Shape)
    AppID = GetAppID(Shp)
    CurrentColor = Slide1.Shapes("WindowAppNotes:" & AppID).TextFrame2.TextRange.Font.Line.ForeColor.RGB
    Dim NextColor As Long
    
    If CurrentColor = RGB(255, 255, 255) Then
        NextColor = RGB(0, 0, 0)
    ElseIf CurrentColor = RGB(0, 0, 0) Then
        NextColor = RGB(255, 0, 0)
    ElseIf CurrentColor = RGB(255, 0, 0) Then
        NextColor = RGB(0, 255, 0)
    ElseIf CurrentColor = RGB(0, 255, 0) Then
        NextColor = RGB(0, 0, 255)
    ElseIf CurrentColor = RGB(0, 0, 255) Then
        NextColor = RGB(255, 255, 0)
    ElseIf CurrentColor = RGB(255, 255, 0) Then
        NextColor = RGB(0, 255, 255)
    ElseIf CurrentColor = RGB(0, 255, 255) Then
        NextColor = RGB(255, 0, 255)
    Else
        NextColor = RGB(255, 255, 255)
    End If
    
    Slide1.Shapes("WindowAppNotes:" & AppID).TextFrame2.TextRange.Font.Line.ForeColor.RGB = NextColor
    Slide1.AxTextBox.BackColor = NextColor
End Sub

Sub SetNotesFg(Shp As Shape)
    AppID = GetAppID(Shp)
    CurrentColor = Slide1.Shapes("WindowAppNotes:" & AppID).TextFrame.TextRange.Font.Color.RGB
    Dim NextColor As Long
    
    If CurrentColor = RGB(255, 255, 255) Then
        NextColor = RGB(0, 0, 0)
    ElseIf CurrentColor = RGB(0, 0, 0) Then
        NextColor = RGB(255, 0, 0)
    ElseIf CurrentColor = RGB(255, 0, 0) Then
        NextColor = RGB(0, 255, 0)
    ElseIf CurrentColor = RGB(0, 255, 0) Then
        NextColor = RGB(0, 0, 255)
    ElseIf CurrentColor = RGB(0, 0, 255) Then
        NextColor = RGB(255, 255, 0)
    ElseIf CurrentColor = RGB(255, 255, 0) Then
        NextColor = RGB(0, 255, 255)
    ElseIf CurrentColor = RGB(0, 255, 255) Then
        NextColor = RGB(255, 0, 255)
    Else
        NextColor = RGB(255, 255, 255)
    End If
    
    Slide1.Shapes("WindowAppNotes:" & AppID).TextFrame.TextRange.Font.Color.RGB = NextColor
    Slide1.AxTextBox.ForeColor = NextColor
End Sub

Sub SetNotesBlank(Shp As Shape)
    AppID = GetAppID(Shp)
    Slide1.Shapes("WindowAppNotes:" & AppID).TextFrame.TextRange.Text = ""
    Slide1.AxTextBox.Text = ""
End Sub

Sub SetNotesOpen(Shp As Shape)
    AppID = GetAppID(Shp)
    Dim dlgOpen As FileDialog
    Dim strResult As String
    
    Set dlgOpen = Application.FileDialog(Type:=msoFileDialogFilePicker)
    
    With dlgOpen
       .Filters.Clear
      .Filters.Add "All plain text files", "*.*", 1
      .AllowMultiSelect = False
        If .Show = True Then
            strResult = .SelectedItems(1)
            If strResult = "" Then
                MsgBox "No file selected", vbCritical, "Error"
                Exit Sub
            End If
        End If
    End With
    Dim fs, F
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set F = fs.OpenTextFile(strResult, ForReading, True, TristateFalse)
    fContent = F.ReadAll
    F.Close
    Slide1.Shapes("WindowAppNotes:" & AppID).TextFrame.TextRange.Text = fContent
    Slide1.AxTextBox.Text = fContent
End Sub

Sub SetNotesSave(Shp As Shape)
    AppID = GetAppID(Shp)
    If Slide1.Shapes("Shape3AppNotes:" & AppID).TextFrame.TextRange.Text = "Notes" Then
        Dim dlgOpen As FileDialog
        Dim strResult As String
        
        Set dlgOpen = Application.FileDialog(Type:=msoFileDialogFolderPicker)
        
        With dlgOpen
            .Title = "Select folder to save the file to"
            .AllowMultiSelect = False
            If .Show = True Then
                strResult = .SelectedItems(1)
                If strResult = "" Then
                    MsgBox "No folder selected", vbCritical, "Error"
                    Exit Sub
                End If
            End If
        End With
        Filename = InputBox("Choose filename", "Save text document as..")
        Dim fs, F
        Const ForReading = 1, ForWriting = 2, ForAppending = 8
        Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set F = fs.OpenTextFile(strResult + "\\" + Filename, ForWriting, True, TristateFalse)
        F.Write (Slide1.AxTextBox.Text)
        F.Close
    Else
        SetFileContent Slide1.Shapes("Shape3AppNotes:" & AppID).TextFrame.TextRange.Text, Slide1.AxTextBox.Text
    End If
End Sub


