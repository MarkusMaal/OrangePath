' Presentator app

Sub AppPresentator(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Presentator"
    
    Slide2.Shapes("Shape11AppGuess_").TextFrame.TextRange.Text = CStr(Int(100 * Rnd))
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide1.Shapes("WindowTitleAppPresentator:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "Presentator – Untitled presentation"
    UpdateTime
End Sub

' Display the correct slide when restoring a minimized window
Sub AppPresentatorRestore(AppID As String)
    Dim CurrentSlide As String
    Dim Shp As Shape
    CurrentSlide = Replace(Slide1.Shapes("Shape15AppPresentator:" & AppID).TextFrame.TextRange.Text, "Slide ", "")
    For Each Shp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
        If InStr(1, Shp.Name, "PresSld" & CurrentSlide) = 1 Then
            Shp.Visible = msoTrue
        ElseIf InStr(1, Shp.Name, "PresSld") = 1 Then
            Shp.Visible = msoFalse
        End If
    Next Shp
End Sub

Sub PresChangecolor(Shp As Shape)
    Dim AppID As String
    Dim SelShape As String
    AppID = GetAppID(Shp)
    SelShape = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    SetVar "Macro", "PresConfirmChangeColor"
    SetVar "SelShape", SelShape
    SetVar "Shape", Shp.Name
    SetVar "AppID", AppID
    AppModalColorPicker
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
    'ElseIf Shp.Fill.ForeColor.RGB = RGB(0, 255, 255) Then
    '    Shp.Fill.ForeColor.RGB = RGB(255, 128, 0)
    'ElseIf Shp.Fill.ForeColor.RGB = RGB(255, 128, 0) Then
    '    Shp.Fill.ForeColor.RGB = RGB(0, 128, 255)
    'ElseIf Shp.Fill.ForeColor.RGB = RGB(0, 128, 255) Then
    '    Shp.Fill.ForeColor.RGB = RGB(128, 0, 255)
    'ElseIf Shp.Fill.ForeColor.RGB = RGB(128, 0, 255) Then
    '    Shp.Fill.ForeColor.RGB = RGB(128, 128, 128)
    'Else
    '    Shp.Fill.ForeColor.RGB = RGB(255, 0, 255)
    'End If
End Sub

Sub PresConfirmChangeColor()
    Dim SelShape As String
    Dim Shp As Shape
    Dim SelCol As Long
    Dim AppID As String
    
    SelShape = CheckVars("%SelShape%")
    Set Shp = Slide1.Shapes(CheckVars("%Shape%"))
    SelCol = CLng(CheckVars("%InputValue%"))
    AppID = CheckVars("%AppID%")
    
    Shp.Fill.ForeColor.RGB = SelCol
    If SelShape <> "void" Then
        If InStr(1, Shp.Name, "Color2") = 1 Then
            Slide1.Shapes(SelShape).Fill.ForeColor.RGB = Shp.Fill.ForeColor.RGB
        Else
            Slide1.Shapes(SelShape).TextFrame.TextRange.Font.Color.RGB = Shp.Fill.ForeColor.RGB
        End If
    End If
    
    UnsetVar "SelShape"
    UnsetVar "Shape"
    UnsetVar "AppID"
End Sub

Sub AssocPresentator(Shp As Shape)
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    'Filename = Replace(Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text, ".pres", "")
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Presentator"
    Slide2.Shapes("Shape11AppGuess_").TextFrame.TextRange.Text = CStr(Int(100 * Rnd))
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    SetVar "AppID", Slide1.Shapes("AppID").TextFrame.TextRange.Text
    SetVar "InputValue", Filename
    PresLoad
    UpdateTime
End Sub

Sub AssocIPresentator(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocPresentator Slide1.Shapes(ShapeName)
End Sub

Sub AddShape(Shp As Shape)
    Dim AppID As String
    Dim ShapeID As Integer
    Dim ShapeName As String
    Dim ReferenceShape As Shape
    Shp.Copy
    AppID = GetAppID(Shp)
    ShapeID = Slide1.Shapes("RegularApp:" & AppID).GroupItems.Count + 1
    CSld = GetCSld(AppID)
    HasSuchShape = True
    Do While HasSuchShape
        HasSuchShape = False
        ShapeName = "PresSld" & CSld & "Shape" & CStr(ShapeID) & "AppPresentator:" & AppID
        For Each GrpItem In Slide1.Shapes("RegularApp:" & AppID).GroupItems
            If GrpItem.Name = ShapeName Then
                HasSuchShape = True
            End If
        Next GrpItem
        If HasSuchShape Then
            ShapeID = ShapeID + 1
        End If
    Loop
    Set ReferenceShape = Slide1.Shapes("SlideAppPresentator:" & AppID)
    PasteToGroup Shp, Shp, ShapeName, ReferenceShape.Left, ReferenceShape.Top, Slide1, "SelShape"
    With Slide1.Shapes(ShapeName)
        .Fill.ForeColor.RGB = Slide1.Shapes("Color2AppPresentator:" & AppID).Fill.ForeColor.RGB
        .TextFrame.TextRange.Font.Color.RGB = Slide1.Shapes("ColorAppPresentator:" & AppID).Fill.ForeColor.RGB
        .TextFrame.TextRange.Font.Name = "Candara"
        .Line.Visible = msoFalse
        .Line.Transparency = 0
    End With
    SelShape Slide1.Shapes(ShapeName)
End Sub


Sub PresFontPlus(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Size = Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Size + 1
    End If
End Sub

Sub PresFontMinus(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Size = Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Size - 1
    End If
End Sub

Sub PresToggleFont(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        Dim CFont As String
        CFont = Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Name
        If CFont = "Candara" Then
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Name = "Candara Light"
        ElseIf CFont = "Candara Light" Then
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Name = "Arial"
        ElseIf CFont = "Arial" Then
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Name = "Calibri"
        ElseIf CFont = "Calibri" Then
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Name = "Times New Roman"
        ElseIf CFont = "Times New Roman" Then
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Name = "Consolas"
        ElseIf CFont = "Consolas" Then
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Name = "DejaVu Sans"
        Else
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Name = "Candara"
        End If
    End If
End Sub

Sub PresAlignLeft(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        Slide1.Shapes(ShapeName).TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
    End If
End Sub

Sub PresAlignCenter(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        Slide1.Shapes(ShapeName).TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    End If
End Sub

Sub PresAlignRight(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        Slide1.Shapes(ShapeName).TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignRight
    End If
End Sub

Sub PresToggleBold(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        If Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Bold = msoFalse Then
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Bold = msoTrue
        Else
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Bold = msoFalse
        End If
    End If
End Sub

Sub PresToggleItalic(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        If Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Italic = msoFalse Then
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Italic = msoTrue
        Else
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Italic = msoFalse
        End If
    End If
End Sub

Sub PresToggleUnderline(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        If Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Underline = msoFalse Then
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Underline = msoTrue
        Else
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Underline = msoFalse
        End If
    End If
End Sub

Sub PresToggleShadow(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        If Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Shadow = msoFalse Then
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Shadow = msoTrue
        Else
            Slide1.Shapes(ShapeName).TextFrame.TextRange.Font.Shadow = msoFalse
        End If
    End If
End Sub

Sub PresToggleStrikethrough(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        If Slide1.Shapes(ShapeName).TextFrame2.TextRange.Font.Strikethrough = msoFalse Then
            Slide1.Shapes(ShapeName).TextFrame2.TextRange.Font.Strikethrough = msoTrue
        Else
            Slide1.Shapes(ShapeName).TextFrame2.TextRange.Font.Strikethrough = msoFalse
        End If
    End If
End Sub

Function GetCSld(AppID As String) As String
    GetCSld = Replace(Slide1.Shapes("Shape15AppPresentator:" & AppID).TextFrame.TextRange.Text, "Slide ", "")
End Function

Sub SelShape(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text = Shp.Name
    Slide1.Shapes("ColorAppPresentator:" & AppID).Fill.ForeColor.RGB = Shp.TextFrame.TextRange.Font.Color.RGB
    Slide1.Shapes("Color2AppPresentator:" & AppID).Fill.ForeColor.RGB = Shp.Fill.ForeColor.RGB
End Sub

Sub ClearShape(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text = "void"
End Sub

' Move shape in specific direction
Sub ShpMoveLeft(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        ModeSplit = Split(Slide1.Shapes("ButtonModeSelectAppPresentator:" & AppID).TextFrame.TextRange.Text, " ")
        Mode = ModeSplit(1)
        If Mode = "Move" Then
            Slide1.Shapes(ShapeName).Left = Slide1.Shapes(ShapeName).Left - 5
        ElseIf Mode = "Size" Then
            Slide1.Shapes(ShapeName).Width = Slide1.Shapes(ShapeName).Width - 5
        ElseIf Mode = "PMove" Then
            Slide1.Shapes(ShapeName).Left = Slide1.Shapes(ShapeName).Left - 1
        ElseIf Mode = "PSize" Then
            Slide1.Shapes(ShapeName).Width = Slide1.Shapes(ShapeName).Width - 1
        ElseIf Mode = "Rotate" Then
            Slide1.Shapes(ShapeName).Rotation = Slide1.Shapes(ShapeName).Rotation - 1
        End If
    End If
End Sub

Sub ShpMoveRight(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        ModeSplit = Split(Slide1.Shapes("ButtonModeSelectAppPresentator:" & AppID).TextFrame.TextRange.Text, " ")
        Mode = ModeSplit(1)
        If Mode = "Move" Then
            Slide1.Shapes(ShapeName).Left = Slide1.Shapes(ShapeName).Left + 5
        ElseIf Mode = "Size" Then
            Slide1.Shapes(ShapeName).Width = Slide1.Shapes(ShapeName).Width + 5
        ElseIf Mode = "PMove" Then
            Slide1.Shapes(ShapeName).Left = Slide1.Shapes(ShapeName).Left + 1
        ElseIf Mode = "PSize" Then
            Slide1.Shapes(ShapeName).Width = Slide1.Shapes(ShapeName).Width + 1
        ElseIf Mode = "Rotate" Then
            Slide1.Shapes(ShapeName).Rotation = Slide1.Shapes(ShapeName).Rotation + 1
        End If
    End If
End Sub
Sub ShpMoveUp(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        ModeSplit = Split(Slide1.Shapes("ButtonModeSelectAppPresentator:" & AppID).TextFrame.TextRange.Text, " ")
        Mode = ModeSplit(1)
        If Mode = "Move" Then
            Slide1.Shapes(ShapeName).Top = Slide1.Shapes(ShapeName).Top - 5
        ElseIf Mode = "PMove" Then
            Slide1.Shapes(ShapeName).Top = Slide1.Shapes(ShapeName).Top - 1
        ElseIf Mode = "PSize" Then
            Slide1.Shapes(ShapeName).Height = Slide1.Shapes(ShapeName).Height - 1
        ElseIf Mode = "Size" Then
            Slide1.Shapes(ShapeName).Height = Slide1.Shapes(ShapeName).Height - 5
        ElseIf Mode = "Rotate" Then
            Slide1.Shapes(ShapeName).Rotation = Slide1.Shapes(ShapeName).Rotation - 45
        End If
    End If
End Sub

Sub ShpMoveDown(Shp As Shape)
    Dim AppID As String
    Dim Mode As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        ModeSplit = Split(Slide1.Shapes("ButtonModeSelectAppPresentator:" & AppID).TextFrame.TextRange.Text, " ")
        Mode = ModeSplit(1)
        If Mode = "Move" Then
            Slide1.Shapes(ShapeName).Top = Slide1.Shapes(ShapeName).Top + 5
        ElseIf Mode = "Size" Then
            Slide1.Shapes(ShapeName).Height = Slide1.Shapes(ShapeName).Height + 5
        ElseIf Mode = "PMove" Then
            Slide1.Shapes(ShapeName).Top = Slide1.Shapes(ShapeName).Top + 1
        ElseIf Mode = "PSize" Then
            Slide1.Shapes(ShapeName).Height = Slide1.Shapes(ShapeName).Height + 1
        ElseIf Mode = "Rotate" Then
            Slide1.Shapes(ShapeName).Rotation = Slide1.Shapes(ShapeName).Rotation + 45
        End If
    End If
End Sub

Sub ToggleMode(Shp As Shape)
    Dim AppID As String
    Dim Mode As String
    AppID = GetAppID(Shp)
    ModeSplit = Split(Slide1.Shapes("ButtonModeSelectAppPresentator:" & AppID).TextFrame.TextRange.Text, " ")
    Mode = ModeSplit(1)
    If Mode = "Move" Then
        Mode = "PMove"
    ElseIf Mode = "PMove" Then
        Mode = "Size"
    ElseIf Mode = "Size" Then
        Mode = "PSize"
    ElseIf Mode = "PSize" Then
        Mode = "Rotate"
    Else
        Mode = "Move"
    End If
    Slide1.Shapes("ButtonModeSelectAppPresentator:" & AppID).TextFrame.TextRange.Text = "Mode: " & Mode
End Sub

Sub ShpDelFromSld(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    ShapeName = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    If ShapeName <> "void" Then
        Slide1.Shapes(ShapeName).Delete
    End If
    Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text = "void"
End Sub

Sub AppPresentatorLaunchLoadPicDialog(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim Filename As String
    SetVar "Macro", "AppPresentatorImportPic"
    SetVar "AppID", AppID
    SetVar "LaunchDir", "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/"
    UnsetVar "Save"
    AppModalFiles
End Sub

Sub PresLaunchSaveMan(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim Filename As String
    Filename = Replace(Slide1.Shapes("WindowTitleAppPresentator:" & AppID).TextFrame.TextRange.Text, "Presentator – ", "")
    If Filename = "Untitled presentation" Or Shp.TextFrame.TextRange.Text = "Save as.." Then
        SetVar "Macro", "PresSetTitleAndSave"
        SetVar "AppID", AppID
        SetVar "LaunchDir", "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/"
        SetVar "Save", "Yes"
        AppModalFiles
        'AppInputBox "Please enter filename", "Presentator"
    Else
        SavePresentatorFile Filename, AppID
    End If
End Sub

Sub AppPresentatorImportPic()
    Dim AppID As String
    Dim PicFile As String
    Dim SelShape As Shape
    AppID = CheckVars("%AppID%")
    PicFile = CheckVars("%InputValue%")
    Set SelShape = Slide1.Shapes(Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text)
    If Left(PicFile, 3) = "C:\" Then
        SelShape.Fill.UserPicture PicFile
    Else
        PreparePic PicFile
        SelShape.Fill.UserPicture Environ("TEMP") & "\UserPic.PNG"
    End If
End Sub

Sub ClearPresentationFile(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim Limit As Integer
    Limit = Slide1.Shapes("RegularApp:" & AppID).GroupItems.Count
    Dim I As Integer
    For I = Limit To 1 Step -1
        If InStr(Slide1.Shapes("RegularApp:" & AppID).GroupItems(I).Name, "PresSld") = 1 Then
            Slide1.Shapes("RegularApp:" & AppID).GroupItems(I).Delete
        End If
    Next I
    
    Slide1.Shapes("WindowTitleAppPresentator:" & AppID).TextFrame.TextRange.Text = "Presentator – Untitled presentation"
    Slide1.Shapes("Shape15AppPresentator:" & AppID).TextFrame.TextRange.Text = "Slide 1"
    Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text = "void"
End Sub

Sub SavePresentatorFile(ByVal Filename As String, ByVal AppID As String)
    ' Creates a shape range of all of the shapes in the presentation, then writes it as a group to ShapeFS
    Slide1.Shapes("WindowTitleAppPresentator:" & AppID).TextFrame.TextRange.Text = "Presentator – Saving..."
    WaitCursor Slide1.Shapes("WindowAppPresentator:" & AppID), ""
    If Right(Filename, 5) <> ".pres" Then
        Filename = Filename & ".pres"
    End If
    HideCursor
    Dim Shp As Shape
    Dim Shp2 As Shape
    Dim Shapes As String
    Shapes = ""
    For Each Shp2 In Slide1.Shapes("RegularApp:" & AppID).GroupItems()
        If InStr(1, Shp2.Name, "PresSld") = 1 Then
            Shapes = Shapes & Shp2.Name & ","
        ElseIf Shp2.Name = "SlideAppPresentator:" & AppID Then
            Shapes = Shapes & Shp2.Name & ","
        End If
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
    Slide1.Shapes("WindowTitleAppPresentator:" & AppID).TextFrame.TextRange.Text = "Presentator – " & Filename
    WriteGroup Filename, Slide1.Shapes.Range(ShapesX), "SlideAppPresentator:" & AppID
End Sub

Sub PresLaunchLoadMan(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim Filename As String
    SetVar "Macro", "PresLoad"
    SetVar "AppID", AppID
    SetVar "LaunchDir", "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/"
    UnsetVar "Save"
    AppModalFiles
End Sub

Sub PresLoad()
    Dim AppID As String
    Dim TextValue As String
    Dim OffX As Integer
    Dim OffY As Integer
    AppID = CheckVars("%AppID%")
    Slide1.Shapes("WindowTitleAppPresentator:" & AppID).TextFrame.TextRange.Text = "Presentator – Loading..."
    WaitCursor Slide1.Shapes("WindowAppPresentator:" & AppID), ""
    HideCursor
    TextValue = CheckVars("%InputValue%")
    ClearPresentationFile Slide1.Shapes("RegularApp:" & AppID)
    OffX = Slide1.Shapes("SlideAppPresentator:" & AppID).Left
    OffY = Slide1.Shapes("SlideAppPresentator:" & AppID).Top
    sizeX = Slide1.Shapes("SlideAppPresentator:" & AppID).Width
    sizeY = Slide1.Shapes("SlideAppPresentator:" & AppID).Height
    If FileStreamsExist(TextValue) = False Then
        AppMessage "Load error. File does not exist!", "Presentator", "Error", True
        Exit Sub
    End If
    If Left(GetFileRef(TextValue).GroupItems(1).Name, 7) <> "SizeKey" Or Left(GetFileRef(TextValue).GroupItems(2).Name, 7) <> "PresSld" Then
        AppMessage "Load error. File may be corrupt or unsupported.", "Presentator", "Error", True
        Exit Sub
    End If
    ReadGroup TextValue, Slide1, OffX, OffY, Slide1.Shapes("RegularApp:" & AppID).GroupItems(1), sizeX, sizeY
    Slide1.Shapes("WindowTitleAppPresentator:" & AppID).TextFrame.TextRange.Text = "Presentator – " & TextValue
    GotoFirstSlide Slide1.Shapes("RegularApp:" & AppID)
End Sub

Sub PresSetTitleAndSave(Shp As Shape)
    Dim AppID As String
    Dim TextValue As String
    AppID = CheckVars("%AppID%")
    TextValue = CheckVars("%InputValue%")
    Slide1.Shapes("WindowTitleAppPresentator:" & AppID).TextFrame.TextRange.Text = "Presentator – " & TextValue
    SavePresentatorFile TextValue, AppID
End Sub

Sub PresLaunchTextEdit(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    SetVar "Macro", "PresChangeText"
    SetVar "AppID", AppID
    AppInputBox "Please enter a text value", "Presentator"
End Sub

Sub PresChangeText(Shp As Shape)
    Dim AppID As String
    Dim TextShp As String
    AppID = CheckVars("%AppID%")
    TextValue = CheckVars("%InputValue%")
    UnsetVar "AppID"
    UnsetVar "InputValue"
    TextShp = Slide1.Shapes("ShapeSelAppPresentator:" & AppID).TextFrame.TextRange.Text
    Slide1.Shapes(TextShp).TextFrame.TextRange.Text = TextValue
End Sub

Sub GotoFirstSlide(Shp As Shape)
    Dim AppID As String
    Dim Sld As String
    Dim NextSld As Integer
    Dim SNextSld As String
    Dim SShp As Shape
    AppID = GetAppID(Shp)
    For Each SShp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
        If InStr(1, SShp.Name, "PresSld1") = 1 Then
            SShp.Visible = msoTrue
        ElseIf InStr(1, SShp.Name, "PresSld") = 1 Then
            SShp.Visible = msoFalse
        End If
    Next SShp
    Slide1.Shapes("Shape15AppPresentator:" & AppID).TextFrame.TextRange.Text = "Slide 1"
End Sub

Sub NextSlide(Shp As Shape)
    Dim AppID As String
    Dim Sld As String
    Dim NextSld As Integer
    Dim SNextSld As String
    Dim SShp As Shape
    AppID = GetAppID(Shp)
    Sld = GetCSld(AppID)
    NextSld = CInt(Sld) + 1
    SNextSld = CStr(NextSld)
    For Each SShp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
        If InStr(1, SShp.Name, "PresSld" & Sld) = 1 Then
            SShp.Visible = msoFalse
        End If
        If InStr(1, SShp.Name, "PresSld" & SNextSld) = 1 Then
            SShp.Visible = msoTrue
        End If
    Next SShp
    Slide1.Shapes("Shape15AppPresentator:" & AppID).TextFrame.TextRange.Text = "Slide " & SNextSld
End Sub

Sub PrevSlide(Shp As Shape)
    Dim AppID As String
    Dim Sld As String
    Dim PreSld As Integer
    Dim SPreSld As String
    Dim SShp As Shape
    AppID = GetAppID(Shp)
    Sld = GetCSld(AppID)
    PreSld = CInt(Sld) - 1
    If PreSld < 1 Then
        Exit Sub
    End If
    SPreSld = CStr(PreSld)
    For Each SShp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
        If InStr(1, SShp.Name, "PresSld" & Sld) = 1 Then
            SShp.Visible = msoFalse
        End If
        If InStr(1, SShp.Name, "PresSld" & SPreSld) = 1 Then
            SShp.Visible = msoTrue
        End If
    Next SShp
    Slide1.Shapes("Shape15AppPresentator:" & AppID).TextFrame.TextRange.Text = "Slide " & SPreSld
End Sub

Sub DelSlide(Shp As Shape)
    Dim AppID As String
    Dim SShp As Shape
    Dim Sld As String
    Dim ShapeCount As Integer
    Dim I As Integer
    AppID = GetAppID(Shp)
    Sld = GetCSld(AppID)
    ShapeCount = Slide1.Shapes("RegularApp:" & AppID).GroupItems.Count
    For I = ShapeCount To 0 Step -1
        Set SShp = Slide1.Shapes("RegularApp:" & AppID).GroupItems(I)
        If InStr(1, SShp.Name, "PresSld" & Sld) = 1 Then
            SShp.Delete
        End If
    Next I
End Sub

Sub StartShow(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Sld = GetCSld(AppID)
    ActivePresentation.SlideShowWindow.View.GotoSlide 28
    ReadSlide Sld, AppID
End Sub

Sub ReadSlide(ByVal SlideNo As String, ByVal AppID As String)
    ' Clear any slides, that are being already displayed
    Dim I As Integer
    For I = Slide27.Shapes.Count To 1 Step -1
        If Slide27.Shapes(I).Name <> "SlideShowWindow" Then
            Slide27.Shapes(I).Delete
        End If
    Next
    ' Create an array of shapes, which match the criteria
    Dim Shp As Shape
    Dim Shp2 As Shape
    Dim Shapes As String
    Shapes = ""
    For Each Shp2 In Slide1.Shapes("RegularApp:" & AppID).GroupItems()
        If InStr(1, Shp2.Name, "PresSld" & SlideNo) = 1 Then
            Shapes = Shapes & Shp2.Name & ","
        ElseIf Shp2.Name = "SlideAppPresentator:" & AppID Then
            Shapes = Shapes & Shp2.Name & ","
        End If
    Next Shp2
    
    SplitShapes = Split(Shapes, ",")
    ' End of show, return to Slide1
    If UBound(SplitShapes) < 2 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 4
        UpdateTime
        Exit Sub
    End If
    UJ = CInt(UBound(SplitShapes))
    Dim ShapesX() As String
    
    ReDim ShapesX(UJ)
    For I = 0 To CInt(UBound(SplitShapes) - 1)
        CShape = SplitShapes(I)
        If Not IsInArray(CStr(CShape), ShapesX) Then
            ShapesX(I) = SplitShapes(I)
        End If
    Next
    ' Copy all selected shapes
    Slide1.Shapes.Range(ShapesX).Copy
    ' Paste and group selected shapes to slide 27
    With Slide27.Shapes.Paste.Group
        .Name = "_Slide" & SlideNo & "App:" & AppID
        .Visible = msoTrue
        Dim GI As Shape
        ' Replace macro for each group shape
        For Each GI In .GroupItems
            With GI.ActionSettings(ppMouseClick)
                .Run = "AdvanceShow"
            End With
            If GI.Name = "SlideAppPresentator:" & AppID Then
                GI.Visible = msoFalse
            End If
            GI.TextFrame.TextRange.Font.Size = GI.TextFrame.TextRange.Font.Size * 2.725
        Next GI
    End With
    ' Stretch the slide to full screen
    Slide27.Shapes("_Slide" & SlideNo & "App:" & AppID).Left = 0
    Slide27.Shapes("_Slide" & SlideNo & "App:" & AppID).Top = 0
    Slide27.Shapes("_Slide" & SlideNo & "App:" & AppID).Width = Slide27.Shapes("SlideShowWindow").Width
    Slide27.Shapes("_Slide" & SlideNo & "App:" & AppID).Height = Slide27.Shapes("SlideShowWindow").Height
End Sub

Sub AdvanceShow()
    Dim Shp As Shape
    Dim Ref As Shape
    For Each Shp In Slide27.Shapes
        If InStr(1, Shp.Name, "_") = 1 Then
            Set Ref = Shp
        End If
    Next Shp
    If Ref Is Nothing Then Exit Sub
    Dim RefSplit() As String
    RefSplit = Split(Ref.Name, ":")
    AppID = RefSplit(1)
    FirstPart = RefSplit(0)
    FirstPart2 = Replace(FirstPart, "_Slide", "")
    SlideNo = CInt(Replace(FirstPart2, "App", "")) + 1
    ReadSlide CStr(SlideNo), AppID
End Sub

Sub AppPresentatorSizeChanged(AppID As String)
    Dim B1 As Shape
    Dim B2 As Shape
    Dim B3 As Shape
    Dim B4 As Shape
    Dim B5 As Shape
    Set B1 = Slide1.Shapes("66*20*SW*Shape7AppPresentator:" & AppID)
    Set B2 = Slide1.Shapes("66*20*NW*Shape10AppPresentator:" & AppID)
    Set B3 = Slide1.Shapes("66*20*NW*Shape11AppPresentator:" & AppID)
    Set B4 = Slide1.Shapes("66*20*NW*SaveAsAppPresentator:" & AppID)
    Set B5 = Slide1.Shapes("66*20*NW*ShowButtonAppPresentator:" & AppID)
    B2.Left = B1.Left + B1.Width + 7.83
    B3.Left = B2.Left + B2.Width + 7.83
    B4.Left = B3.Left + B3.Width + 7.83
    B5.Left = B4.Left + B4.Width + 7.83
    B2.Top = B1.Top
    B3.Top = B1.Top
    B4.Top = B1.Top
    B5.Top = B1.Top
End Sub