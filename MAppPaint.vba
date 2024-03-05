' Pixel Paint app

Sub AppPaint(Shp As shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Paint"
    Slide2.Shapes("AppPaint").Visible = msoTrue
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide2.Shapes("AppPaint").Visible = msoFalse
End Sub

Sub Recolor(Shp As shape)
    SplitZ = Split(Shp.ParentGroup.Name, ":")
    AppID = SplitZ(1)
    If Shp.Fill.ForeColor.RGB = Slide1.Shapes("16*16*NE*Shape7AppPaint:" & AppID).Fill.ForeColor.RGB Then
        Shp.Fill.ForeColor.RGB = Slide1.Shapes("16*16*NE*Shape9AppPaint:" & AppID).Fill.ForeColor.RGB
    Else
        Shp.Fill.ForeColor.RGB = Slide1.Shapes("16*16*NE*Shape7AppPaint:" & AppID).Fill.ForeColor.RGB
    End If
End Sub


Sub AssocPaint(Shp As shape)
    ' Get full file path from shape
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    ' Launch paint app
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Paint"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    ' Get AppID of newly created window
    AppID = Slide1.Shapes("AppID").TextFrame.TextRange.Text
    SetVar "AppID", AppID
    SetVar "InputValue", Filename
    LoadDrawing2
End Sub

Sub AssocIPaint(Shp As shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocPaint Slide1.Shapes(ShapeName)
End Sub

Sub Changecolor(Shp As shape)
    If Shp.Fill.ForeColor.RGB = RGB(255, 0, 255) Then
        Shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ElseIf Shp.Fill.ForeColor.RGB = RGB(255, 255, 255) Then
        Shp.Fill.ForeColor.RGB = RGB(0, 0, 0)
    ElseIf Shp.Fill.ForeColor.RGB = RGB(0, 0, 0) Then
        Shp.Fill.ForeColor.RGB = RGB(255, 0, 0)
    ElseIf Shp.Fill.ForeColor.RGB = RGB(255, 0, 0) Then
        Shp.Fill.ForeColor.RGB = RGB(0, 255, 0)
    ElseIf Shp.Fill.ForeColor.RGB = RGB(0, 255, 0) Then
        Shp.Fill.ForeColor.RGB = RGB(255, 255, 0)
    ElseIf Shp.Fill.ForeColor.RGB = RGB(255, 255, 0) Then
        Shp.Fill.ForeColor.RGB = RGB(0, 0, 255)
    ElseIf Shp.Fill.ForeColor.RGB = RGB(0, 0, 255) Then
        Shp.Fill.ForeColor.RGB = RGB(0, 255, 255)
    Else
        Shp.Fill.ForeColor.RGB = RGB(255, 0, 255)
    End If
End Sub

Sub ClearAll(Shp As shape)
    SplitZ = Split(Shp.Name, ":")
    AppID = SplitZ(1)
    Dim X As Long
    Dim oShp As shape
    For Each oShp In Slide1.Shapes
        With Slide1.Shapes("RegularApp:" & AppID)
            For X = 1 To .GroupItems.Count
                With .GroupItems(X)
                    If InStr(.Name, "Rectangle") Then
                        .Fill.ForeColor.RGB = Slide1.Shapes("16*16*NE*Shape9AppPaint:" + AppID).Fill.ForeColor.RGB
                    End If
                End With
            Next
        End With
        If InStr(oShp.Name, "Shape") Then MsgBox (oShp.Name)
    Next
End Sub

Sub SaveDrawing(Shp As shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim Filename As String
    SetVar "Macro", "SaveDrawing2"
    SetVar "AppID", AppID
    SetVar "Save", "Yes"
    AppModalFiles
End Sub

Sub SaveDrawing2()
    AppID = CheckVars("%AppID%")
    Data = ""
    Username = Slide1.Shapes("Username").TextFrame.TextRange.Text
    Dim X As Long
    Dim oShp As shape
    For Each oShp In Slide1.Shapes
        With Slide1.Shapes("RegularApp:" & AppID)
            For X = 1 To .GroupItems.Count
                With .GroupItems(X)
                    If InStr(.Name, "Rectangle") Then
                        Data = Data & CStr(.Fill.ForeColor.RGB) & ";"
                    End If
                End With
            Next
        End With
    Next
    SetFileContent CheckVars("%InputValue%"), Data
End Sub

Sub LoadDrawing(Shp As shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim Filename As String
    SetVar "Macro", "LoadDrawing2"
    SetVar "AppID", AppID
    UnsetVar "Save"
    AppModalFiles
End Sub

Sub LoadDrawing2()
    On Error GoTo Crash
    AppID = CheckVars("%AppID%")
    Data = ""
    Username = Slide1.Shapes("Username").TextFrame.TextRange.Text
    If FileExists(CheckVars("%InputValue%")) Then
        Data = GetFileContent(CheckVars("%InputValue%"))
    Else
        AppMessage "Saved drawing not found", "Load drawing", "Error", True
        Exit Sub
    End If
    DataSplit = Split(Data, ";")
    Dim X As Long
    Dim IDX As Long
    Dim oShp As shape
    For Each oShp In Slide1.Shapes
        With Slide1.Shapes("RegularApp:" & AppID)
            For X = 1 To .GroupItems.Count
                With .GroupItems(X)
                    If InStr(.Name, "Rectangle") Then
                        Fill = DataSplit(IDX)
                        .Fill.ForeColor.RGB = CLng(Fill)
                        IDX = IDX + 1
                    End If
                End With
            Next
        End With
    Next
Done:
    Exit Sub
Crash:
    Slide22.Shapes("Details").TextFrame.TextRange.Text = "Error details: PAINT_LOAD_ERROR"
    ActivePresentation.SlideShowWindow.View.GotoSlide 22
End Sub

Sub ExportDrawing(Shp As shape)
'Sub ExportDrawing()

    AppID = GetAppID(Shp)
    'AppID = "2"
    Dim dlgOpen As FileDialog
    Dim strResult As String
    
    Set dlgOpen = Application.FileDialog(Type:=msoFileDialogFolderPicker)
    
    With dlgOpen
        .Title = "Select folder to save the file to"
        .AllowMultiSelect = False
        If .Show = True Then
            strResult = .SelectedItems(1)
            If strResult = "" Then
                AppMessage "No folder selected", "Error", "Error", True
                Exit Sub
            End If
        End If
    End With
    Filename = InputBox("Choose filename", "Export drawing as..")
    Slide1.Shapes("RegularApp:" & AppID).Ungroup
    Slide1.Shapes("Shape6AppPaint:" & AppID).Copy
    Application.CommandBars.ExecuteMso "Undo"
    With Slide6.Shapes.Paste
        .Left = 0
        .Top = 0
        .Width = ActivePresentation.PageSetup.SlideWidth
        .Height = ActivePresentation.PageSetup.SlideHeight
        .Name = "Export"
    End With
    If Right(LCase(Filename), 4) <> ".bmp" Then
        Filename = Filename & ".bmp"
    End If
    Slide6.Export strResult & "\\" & Filename, "BMP", 672, 512
End Sub
