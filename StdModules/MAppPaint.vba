' Pixel Paint app

Sub AppPaint(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Paint"
    Slide2.Shapes("AppPaint").Visible = msoTrue
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide2.Shapes("AppPaint").Visible = msoFalse
    UpdateTime
End Sub

Sub Recolor(Shp As Shape)
    SplitZ = Split(Shp.ParentGroup.Name, ":")
    AppID = SplitZ(1)
    If Shp.Fill.ForeColor.RGB = Slide1.Shapes("16*16*NE*Shape7AppPaint:" & AppID).Fill.ForeColor.RGB Then
        Shp.Fill.ForeColor.RGB = Slide1.Shapes("16*16*NE*Shape9AppPaint:" & AppID).Fill.ForeColor.RGB
    Else
        Shp.Fill.ForeColor.RGB = Slide1.Shapes("16*16*NE*Shape7AppPaint:" & AppID).Fill.ForeColor.RGB
    End If
End Sub


Sub AssocPaint(Shp As Shape)
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
    UpdateTime
End Sub

Sub AssocIPaint(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocPaint Slide1.Shapes(ShapeName)
End Sub

Sub Changecolor(Shp As Shape)
    SetVar "Macro", "AppPaintChangeColor"
    SetVar "Shape", Shp.Name
    AppModalColorPicker
    Exit Sub
    
    'If Shp.Fill.ForeColor.RGB = RGB(255, 0, 255) Then
    '    Shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    'ElseIf Shp.Fill.ForeColor.RGB = RGB(255, 255, 255) Then
    '    Shp.Fill.ForeColor.RGB = RGB(0, 0, 0)
    'ElseIf Shp.Fill.ForeColor.RGB = RGB(0, 0, 0) Then
    '    Shp.Fill.ForeColor.RGB = RGB(255, 0, 0)
    'ElseIf Shp.Fill.ForeColor.RGB = RGB(255, 0, 0) Then
    '    Shp.Fill.ForeColor.RGB = RGB(0, 255, 0)
    'ElseIf Shp.Fill.ForeColor.RGB = RGB(0, 255, 0) Then
    '    Shp.Fill.ForeColor.RGB = RGB(255, 255, 0)
    'ElseIf Shp.Fill.ForeColor.RGB = RGB(255, 255, 0) Then
    '    Shp.Fill.ForeColor.RGB = RGB(0, 0, 255)
    'ElseIf Shp.Fill.ForeColor.RGB = RGB(0, 0, 255) Then
    '    Shp.Fill.ForeColor.RGB = RGB(0, 255, 255)
    'Else
    '    Shp.Fill.ForeColor.RGB = RGB(255, 0, 255)
    'End If
End Sub

Sub AppPaintChangeColor()
    Dim Shp As Shape
    Dim Clr As Long
    
    Set Shp = Slide1.Shapes(CheckVars("%Shape%"))
    Clr = CLng(CheckVars("%InputValue%"))
    
    Shp.Fill.ForeColor.RGB = Clr
    
    UnsetVar "Shape"
    UnsetVar "InputValue"
End Sub

Sub ClearAll(Shp As Shape)
    SplitZ = Split(Shp.Name, ":")
    AppID = SplitZ(1)
    WaitCursor Slide1.Shapes("WindowAppPaint:" & AppID), "Clearing..."
    Dim x As Long
    Dim oshp As Shape
    For Each oshp In Slide1.Shapes
        With Slide1.Shapes("RegularApp:" & AppID)
            For x = 1 To .GroupItems.Count
                With .GroupItems(x)
                    If InStr(.Name, "Rectangle") Or InStr(.Name, "Ristkülik") Then
                        .Fill.ForeColor.RGB = Slide1.Shapes("16*16*NE*Shape9AppPaint:" + AppID).Fill.ForeColor.RGB
                    End If
                End With
            Next
        End With
        If InStr(oshp.Name, "Shape") Then MsgBox (oshp.Name)
    Next
    HideCursor
End Sub

Sub SaveDrawing(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim Filename As String
    SetVar "Macro", "SaveDrawing2"
    SetVar "AppID", AppID
    SetVar "Save", "Yes"
    SetVar "LaunchDir", "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/"
    AppModalFiles
End Sub

Sub SaveDrawing2()
    Dim FilePath As String
    FilePath = CheckVars("%InputValue%")
    If Right(FilePath, 4) <> ".pxl" Then
        FilePath = FilePath & ".pxl"
    End If
    AppID = CheckVars("%AppID%")
    Data = ""
    Username = Slide1.Shapes("Username").TextFrame.TextRange.Text
    Dim x As Long
    Dim oshp As Shape
    For Each oshp In Slide1.Shapes
        With Slide1.Shapes("RegularApp:" & AppID)
            For x = 1 To .GroupItems.Count
                With .GroupItems(x)
                    If InStr(.Name, "Rectangle") Or InStr(.Name, "Ristkülik") Then
                        Data = Data & CStr(.Fill.ForeColor.RGB) & ";"
                    End If
                End With
            Next
        End With
    Next
    SetFileContent FilePath, Data
End Sub

Sub LoadDrawing(Shp As Shape)
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
    WaitCursor Slide1.Shapes("WindowAppPaint:" & AppID), "Loading..."
    Data = ""
    Username = Slide1.Shapes("Username").TextFrame.TextRange.Text
    If FileExists(CheckVars("%InputValue%")) Then
        Data = GetFileContent(CheckVars("%InputValue%"))
    Else
        AppMessage "Saved drawing not found", "Load drawing", "Error", True
        Exit Sub
    End If
    DataSplit = Split(Data, ";")
    Dim x As Long
    Dim IDX As Long
    Dim oshp As Shape
    For Each oshp In Slide1.Shapes
        With Slide1.Shapes("RegularApp:" & AppID)
            For x = 1 To .GroupItems.Count
                With .GroupItems(x)
                    If InStr(.Name, "Rectangle") Or InStr(.Name, "Ristkülik") Then
                        If UBound(DataSplit) > IDX Then
                            Fill = DataSplit(IDX)
                            If Fill <> "" Then
                            .Fill.ForeColor.RGB = CLng(Fill)
                            End If
                        End If
                        IDX = IDX + 1
                    End If
                End With
            Next
        End With
    Next
Done:
    HideCursor
    Exit Sub
Crash:
    OSCrash "PAINT_LOAD_ERROR", Err
End Sub

Sub ExportDrawing(Shp As Shape)
    SetVar "AppID", GetAppID(Shp)
    SetVar "Macro", "AppPaintExportDrawing2"
    SetVar "Save", "True"
    SetVar "LaunchDir", "/"
    AppModalFiles
End Sub

Sub AppPaintExportDrawing2()
    AppID = CheckVars("%AppID%")
    Filename = CheckVars("%InputValue%")
    Slide1.Shapes("RegularApp:" & AppID).Ungroup
    Slide1.Shapes("Shape6AppPaint:" & AppID).Export Environ("TEMP") & "\\Userpic.png", ppShapeFormatPNG
    ' This undo conflicts with ModalFiles, must regroup manually instead
    'Application.CommandBars.ExecuteMso "Undo"
    SetFilePic Filename, Environ("TEMP") & "\\Userpic.png"
    UnsetVar "AppID"
    UnsetVar "InputValue"
    
    Dim Sld As Slide
    Set Sld = Slide1
    
    
    ' Regroup
    Dim Shapes As String
    Shapes = ""
    Dim Shp2 As Shape
    For Each Shp2 In Sld.Shapes()
        If Shp2.Name = "TaskIcon:" & AppID Then
            GoTo Continue
        ElseIf InStr(Shp2.Name, ":" & AppID) Then
            'If InStr(Shp2.Name, "AXTextBox") Then ApplyTbAttribs Shp2
            Shapes = Shapes & Shp2.Name & ","
        End If
Continue:
    Next Shp2
    
    SplitShapes = Split(Shapes, ",")
    UJ = CInt(UBound(SplitShapes))
    Dim ShapesX() As String
    
    ReDim ShapesX(UJ)
    For I = 0 To CInt(UBound(SplitShapes) - 1)
        CShape = SplitShapes(I)
        If Not IsInArray(CStr(CShape), ShapesX) Then
            ShapesX(I) = SplitShapes(I)
        End If
    Next
    With Sld.Shapes.Range(ShapesX).Group
        .Name = "RegularApp:" & AppID
    End With
End Sub

' SizeChanged event hook for Paint
Sub AppPaintSizeChanged(AppID As String)
    Dim B1 As Shape
    Dim B2 As Shape
    Dim B3 As Shape
    Dim B4 As Shape
    Dim L1 As Shape
    Dim L2 As Shape
    Dim L3 As Shape
    Dim L4 As Shape
    
    Set B1 = Slide1.Shapes("63*16*SE*ButtonShape11AppPaint:" & AppID) ' Clear
    Set B2 = Slide1.Shapes("63*16*SE*ButtonShape12AppPaint:" & AppID) ' Save
    Set B3 = Slide1.Shapes("63*16*SE*ButtonShape13AppPaint:" & AppID) ' Load
    Set B4 = Slide1.Shapes("63*16*SE*ButtonShape14AppPaint:" & AppID) ' Export
    Set L1 = Slide1.Shapes("41*17*NW*Shape8AppPaint:" & AppID)  ' Color 1
    Set L2 = Slide1.Shapes("41*17*NW*Shape10AppPaint:" & AppID) ' Color 2
    Set L3 = Slide1.Shapes("16*16*NE*Shape7AppPaint:" & AppID)  ' Color 1 (Box)
    Set L4 = Slide1.Shapes("16*16*NE*Shape9AppPaint:" & AppID)  ' Color 2 (Box)
    
    ' Move buttons closer to each other
    B2.Top = B1.Top - B1.Height - 5
    B3.Top = B2.Top - B1.Height - 5
    B4.Top = B3.Top - B1.Height - 5
    
    ' Move color labels closer to color boxes
    L1.Left = L3.Left - L1.Width
    L2.Left = L4.Left - L2.Width
End Sub

