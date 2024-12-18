' Notes app

Sub AppNotes(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Notes"
    
    Slide2.Shapes("Shape11AppGuess_").TextFrame.TextRange.Text = CStr(Int(100 * Rnd))
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide1.Shapes("WindowTitleAppNotes:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "Notepad"
    UpdateTime
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
    If AAX Then Slide1.AxTextBox.Text = TextData
    UpdateTime
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
    If AAX Then Slide1.AxTextBox.BackColor = NextColor
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
    If AAX Then Slide1.AxTextBox.ForeColor = NextColor
End Sub

Sub SetNotesBlank(Shp As Shape)
    AppID = GetAppID(Shp)
    Slide1.Shapes("WindowAppNotes:" & AppID).TextFrame.TextRange.Text = ""
    If AAX Then Slide1.AxTextBox.Text = ""
End Sub

Sub AppNotesFinalizeOpen()
    Dim AppID As String
    AppID = CheckVars("%AppID%")
    UnsetVar "AppID"
    UnsetVar "Macro"
    Dim Content As String
    Content = GetFileContent(CheckVars("%InputValue%"))
    If Content <> "*" Then
        Slide1.Shapes("WindowAppNotes:" & AppID).TextFrame.TextRange.Text = Content
        If AAX Then Slide1.AxTextBox.Text = Content
    Else
        AppMessage "Access is denied", "Cannot open", "Error", True
    End If
    UnsetVar "InputValue"
    FocusWindow AppID
End Sub

Sub AppNotesFinalizeSave()
    Dim AppID As String
    AppID = CheckVars("%AppID%")
    UnsetVar "Macro"
    UnsetVar "AppID"
    UnsetVar "Save"
    SetFileContent CheckVars("%InputValue%"), Slide1.Shapes("AXTextBox2AppNotes:" & AppID).TextFrame.TextRange.Text
    Slide1.Shapes("WindowTitleAppNotes:" & AppID).TextFrame.TextRange.Text = CheckVars("%InputValue%")
    UnsetVar "InputValue"
End Sub

Sub SetNotesOpen(Shp As Shape)
    AppID = GetAppID(Shp)
    SetVar "Macro", "AppNotesFinalizeOpen"
    SetVar "AppID", AppID
    SetVar "LaunchDir", "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/"
    UnsetVar "Save"
    AppModalFiles
End Sub

Sub SetNotesSave(Shp As Shape)
    AppID = GetAppID(Shp)
    If Slide1.Shapes("WindowTitleAppNotes:" & AppID).TextFrame.TextRange.Text = "Notepad" Then
        SetVar "Save", "True"
        SetVar "AppID", AppID
        SetVar "LaunchDir", "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/"
        SetVar "Macro", "AppNotesFinalizeSave"
        AppModalFiles
    Else
        If AAX Then SetFileContent Slide1.Shapes("WindowTitleAppNotes:" & AppID).TextFrame.TextRange.Text, Slide1.AxTextBox.Text
    End If
End Sub

Sub TestNotesLoad()
    AssocModal Slide1.Shapes("FileLabel4AppModalFiles:17")
End Sub
