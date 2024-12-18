' Words app (Generated from devCreateApp)

' This is executed when the application is launched
Sub AppWords(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Words"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    WordsCheckAssociations
    Slide1.Shapes("WindowTitleAppWords:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "Words – New Document"
    UpdateTime
End Sub

' This gets executed when a user clicks a file, which is associated with this application
Sub AssocWords(Shp As Shape)
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Words"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide1.Shapes("WindowTitleAppWords:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "Words – " & Filename
    SetVar "AppID", Slide1.Shapes("AppID").TextFrame.TextRange.Text
    SetVar "InputValue", Filename
    AppWordsOpen2
    UpdateTime
End Sub

' This gets executed when a user clicks icon of a file, which is associated with this application
Sub AssocIWords(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocWords Slide1.Shapes(ShapeName)
End Sub

Sub AppWordsSwapBgFg(Shp As Shape)
    If Shp.TextFrame.TextRange.Text = "BG" Then
        Shp.TextFrame.TextRange.Text = "FG"
    Else
        Shp.TextFrame.TextRange.Text = "BG"
    End If
End Sub

Sub WordsCheckAssociations()
    If GetSysConfig("IconWords") = "*" Then
        SaveSysConfig "IconWords", "/System/Icons/Words.emf"
    End If
    If GetSysConfig("assocwdoc") = "*" Then
        SaveSysConfig "assocwdoc", "Words"
    End If
End Sub

Sub AppWordsFontPopulate()
    If Not AAX Then Exit Sub
    Slide1.AxComboBox.Clear
    Slide1.AxComboBox.AddItem ("Arial")
    Slide1.AxComboBox.AddItem ("Arial Black")
    Slide1.AxComboBox.AddItem ("Arial Narrow")
    Slide1.AxComboBox.AddItem ("Caladea")
    Slide1.AxComboBox.AddItem ("Calibri")
    Slide1.AxComboBox.AddItem ("Cambria")
    Slide1.AxComboBox.AddItem ("Candara")
    Slide1.AxComboBox.AddItem ("Comic Sans MS")
    Slide1.AxComboBox.AddItem ("Consolas")
    Slide1.AxComboBox.AddItem ("Constantia")
    Slide1.AxComboBox.AddItem ("Corbel")
    Slide1.AxComboBox.AddItem ("Courier")
    Slide1.AxComboBox.AddItem ("Courier New")
    Slide1.AxComboBox.AddItem ("DejaVu Sans")
    Slide1.AxComboBox.AddItem ("DejaVu Sans Light")
    Slide1.AxComboBox.AddItem ("EmojiOne Color")
    Slide1.AxComboBox.AddItem ("Fixedsys")
    Slide1.AxComboBox.AddItem ("Franklin Gothic")
    Slide1.AxComboBox.AddItem ("Gabriola")
    Slide1.AxComboBox.AddItem ("Gadugi")
    Slide1.AxComboBox.AddItem ("Georgia")
    Slide1.AxComboBox.AddItem ("Harlow Solid")
    Slide1.AxComboBox.AddItem ("Impact")
    Slide1.AxComboBox.AddItem ("Liberation Mono")
    Slide1.AxComboBox.AddItem ("Liberation Sans")
    Slide1.AxComboBox.AddItem ("Liberation Serif")
    Slide1.AxComboBox.AddItem ("Lucida Console")
    Slide1.AxComboBox.AddItem ("Lucida Sans Unicode")
    Slide1.AxComboBox.AddItem ("Microsoft JhengHei UI")
    Slide1.AxComboBox.AddItem ("Microsoft Sans Serif")
    Slide1.AxComboBox.AddItem ("Microsoft YaHei UI")
    Slide1.AxComboBox.AddItem ("Modern")
    Slide1.AxComboBox.AddItem ("MS Sans Serif")
    Slide1.AxComboBox.AddItem ("MS Serif")
    Slide1.AxComboBox.AddItem ("Open Sans")
    Slide1.AxComboBox.AddItem ("OpenSymbol")
    Slide1.AxComboBox.AddItem ("Segoe Print")
    Slide1.AxComboBox.AddItem ("Segoe Script")
    Slide1.AxComboBox.AddItem ("Segoe UI")
    Slide1.AxComboBox.AddItem ("Segoe UI Emoji")
    Slide1.AxComboBox.AddItem ("Segoe UI Light")
    Slide1.AxComboBox.AddItem ("Segoe UI Semilight")
    Slide1.AxComboBox.AddItem ("Segoe UI Symbol")
    Slide1.AxComboBox.AddItem ("Small Fonts")
    Slide1.AxComboBox.AddItem ("Symbol")
    Slide1.AxComboBox.AddItem ("System")
    Slide1.AxComboBox.AddItem ("Tahoma")
    Slide1.AxComboBox.AddItem ("Terminal")
    Slide1.AxComboBox.AddItem ("Times New Roman")
    Slide1.AxComboBox.AddItem ("Trebuchet MS")
    Slide1.AxComboBox.AddItem ("Webdings")
    Slide1.AxComboBox.AddItem ("Verdana")
    Slide1.AxComboBox.AddItem ("Wingdings")
    Slide1.AxComboBox.AddItem ("Vivaldi")
End Sub



Sub AppWordsLaunchSaveMan(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim Filename As String
    Filename = Replace(Slide1.Shapes("WindowTitleAppWords:" & AppID).TextFrame.TextRange.Text, "Words – ", "")
    If Filename = "New Document" Or Shp.TextFrame.TextRange.Text = "Save as.." Then
        SetVar "Macro", "AppWordsSetTitleAndSave"
        SetVar "AppID", AppID
        SetVar "Save", "Yes"
        SetVar "LaunchDir", "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/"
        AppModalFiles
    Else
        AppWordsSave Filename, AppID
    End If
End Sub

Sub AppWordsSetTitleAndSave()
    If Not AAX Then Exit Sub
    Dim AppID As String
    Dim Filename As String
    AppID = CheckVars("%AppID%")
    Filename = CheckVars("%InputValue%")
    If Right(Filename, 5) <> ".wdoc" Or Right(Filename, 4) <> ".txt" Then
        Filename = Filename & ".wdoc"
    End If
    UnsetVar "AppID"
    UnsetVar "InputValue"
    Slide1.Shapes("WindowTitleAppWords:" & AppID).TextFrame.TextRange.Text = "Words – " & Filename
    AppWordsSave Filename, AppID
End Sub

Sub AppWordsSave(Filename As String, AppID As String)
    If Right(Filename, 5) = ".wdoc" Then
        Dim AppName As String
        AppName = "Words"
        Dim Shp2 As Shape
        Dim Shapes As String
        Shapes = ""
        For Each Shp2 In Slide1.Shapes("RegularApp:" & AppID).GroupItems()
            ' Check if the shape name starts with "AXTextBox2"
            If InStr(1, Shp2.Name, "AXTextBox2") = 1 Then
                Shapes = Shapes & Shp2.Name & ","
            ' Add a dummy shape to group for size key
            ElseIf InStr(1, Shp2.Name, "DummyShape") = 1 Then
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
        
        WriteGroup Filename, Slide1.Shapes.Range(ShapesX), "DummyShapeAppWords:" & AppID
    Else
        ' Just write plain text if we're not using the .wdoc extension
        SetFileContent Filename, Slide1.Shapes("AXTextBox2AppWords:" & AppID).TextFrame.TextRange.Text
    End If
    AppMessage "File saved as '" & Filename & "'", "Words", "Info", False
End Sub

Sub AppWordsShowFonts(Shp As Shape)
    If Not AAX Then Exit Sub
    Dim AppID As String
    AppID = GetAppID(Shp)
    AppWordsFontPopulate
    SetVar "Macro", "AppWordsSetFont"
    SetVar "AppID", AppID
    Slide1.AxComboBox.Visible = True
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    Slide1.AxComboBox.Left = Slide1.Shapes("FontButtonAppWords:" & AppID).Left
    Slide1.AxComboBox.Top = Slide1.Shapes("FontButtonAppWords:" & AppID).Top
    Slide1.AxComboBox.Width = 200
    Slide1.AxComboBox.DropDown
    Slide1.AxComboBox.Visible = False
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
End Sub

Sub AppWordsSetFont()
    UnsetVar "Macro"
    If AAX Then
        Slide1.AxTextBox.Font = Slide1.AxComboBox.SelText
    End If
    Slide1.Shapes("AXTextBox2AppWords:" & CheckVars("%AppID%")).TextFrame.TextRange.Font.Name = Slide1.AxComboBox.SelText
    UnsetVar "AppID"
End Sub

Sub AppWordsSwapColor(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim TextBox As Shape
    Set TextBox = Slide1.Shapes("AxTextBox2AppWords:" & AppID)
    Dim BgFg As String
    BgFg = GetShapeText(Slide1, AppID, "Words", "BgFgButton")
    If BgFg = "BG" Then
        TextBox.Fill.ForeColor = Shp.Fill.ForeColor
        If AAX Then Slide1.AxTextBox.BackColor = Shp.Fill.ForeColor
    Else
        TextBox.TextFrame.TextRange.Font.Color.RGB = Shp.Fill.ForeColor
        If AAX Then Slide1.AxTextBox.ForeColor = Shp.Fill.ForeColor
    End If
End Sub

Sub AppWordsClear(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim TextBox As Shape
    Set TextBox = Slide1.Shapes("AxTextBox2AppWords:" & AppID)
    TextBox.TextFrame2.TextRange.Font.Strikethrough = msoFalse
    TextBox.TextFrame.TextRange.Font.Bold = msoFalse
    TextBox.TextFrame.TextRange.Font.Italic = msoFalse
    TextBox.TextFrame.TextRange.Font.Underline = msoFalse
    TextBox.TextFrame.TextRange.Font.Name = "Candara"
    TextBox.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    TextBox.TextFrame.TextRange.Font.Size = 14
    TextBox.TextFrame.TextRange.Text = ""
    TextBox.Fill.ForeColor.RGB = RGB(255, 255, 255)
    TextBox.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
    ApplyTbAttribs TextBox
    If AAX Then Slide1.AxTextBox.Text = ""
    Slide1.Shapes("BgFgButtonAppWords:" & AppID).TextFrame.TextRange.Text = "BG"
    Slide1.Shapes("SizeLabelAppWords:" & AppID).TextFrame.TextRange.Text = "Size:14"
    Slide1.Shapes("WindowTitleAppWords:" & AppID).TextFrame.TextRange.Text = "Words – New Document"
End Sub

Sub AppWordsOpen(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    SetVar "Macro", "AppWordsOpen2"
    SetVar "AppID", AppID
    SetVar "LaunchDir", "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/"
    UnsetVar "Save"
    AppModalFiles
End Sub

Function AppWordsGetTextBox(AppID As String) As Shape
    Set AppWordsGetTextBox = Slide1.Shapes("AxTextBox2AppWords:" & AppID)
End Function

Sub AppWordsOpen2()
    Dim AppID As String
    Dim Filename As String
    Dim ShapeRef As Shape
    AppID = CheckVars("%AppID%")
    Filename = CheckVars("%InputValue%")
    Set ShapeRef = GetFileRef(Filename)
    
    Dim TextBox As Shape
    Set TextBox = AppWordsGetTextBox(AppID)
    
    ' no textframe = is group
    If Not ShapeRef.HasTextFrame Then
        Dim OffX As Integer
        Dim OffY As Integer
        Dim sizeX As Integer
        Dim sizeY As Integer
        Dim AppName As String
        AppName = "Words"
        OffX = Slide1.Shapes("AXTextBox2App" & AppName & ":" & AppID).Left
        OffY = Slide1.Shapes("AXTextBox2App" & AppName & ":" & AppID).Top
        sizeX = Slide1.Shapes("AXTextBox2App" & AppName & ":" & AppID).Width
        sizeY = Slide1.Shapes("AXTextBox2App" & AppName & ":" & AppID).Height
        ' Delete the current document to avoid issues regrouping
        Slide1.Shapes("AXTextBox2App" & AppName & ":" & AppID).Delete
        ' Paste group to the same location the document used to be at
        ReadGroup Filename, Slide1, OffX, OffY, Slide1.Shapes("RegularApp:" & AppID).GroupItems(1), sizeX, sizeY
    Else
        Dim FileContent As String
        FileContent = GetFileContent(Filename)
        If FileContent = "*" Then
            AppMessage "Access is denied", "Words", "Error", True
            Exit Sub
        End If
        AppWordsClear Slide1.Shapes("RegularApp:" & AppID)
        TextBox.TextFrame.TextRange.Text = FileContent
        If AAX Then Slide1.AxTextBox.Text = FileContent
    End If
    FocusWindow AppID
    Slide1.Shapes("WindowTitleAppWords:" & AppID).TextFrame.TextRange.Text = "Words – " & Filename
End Sub


Sub AppWordsAlignLeft(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim TextBox As Shape
    Set TextBox = AppWordsGetTextBox(AppID)
    TextBox.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    If AAX Then Slide1.AxTextBox.TextAlign = fmTextAlignLeft
End Sub

Sub AppWordsAlignCenter(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim TextBox As Shape
    Set TextBox = AppWordsGetTextBox(AppID)
    TextBox.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    If AAX Then Slide1.AxTextBox.TextAlign = fmTextAlignCenter
End Sub

Sub AppWordsAlignRight(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim TextBox As Shape
    Set TextBox = AppWordsGetTextBox(AppID)
    TextBox.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignRight
    If AAX Then Slide1.AxTextBox.TextAlign = fmTextAlignRight
End Sub

Sub AppWordsToggleBold(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim TextBox As Shape
    Set TextBox = AppWordsGetTextBox(AppID)
    If TextBox.TextFrame.TextRange.Font.Bold = msoTrue Then
        TextBox.TextFrame.TextRange.Font.Bold = msoFalse
    Else
        TextBox.TextFrame.TextRange.Font.Bold = msoTrue
    End If
    If AAX Then Slide1.AxTextBox.Font.Bold = MsoTristateToBool(TextBox.TextFrame.TextRange.Font.Bold)
End Sub

Sub AppWordsToggleItalic(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim TextBox As Shape
    Set TextBox = AppWordsGetTextBox(AppID)
    If TextBox.TextFrame.TextRange.Font.Italic = msoTrue Then
        TextBox.TextFrame.TextRange.Font.Italic = msoFalse
    Else
        TextBox.TextFrame.TextRange.Font.Italic = msoTrue
    End If
    If AAX Then Slide1.AxTextBox.Font.Italic = MsoTristateToBool(TextBox.TextFrame.TextRange.Font.Italic)
End Sub


Sub AppWordsToggleUnderline(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim TextBox As Shape
    Set TextBox = AppWordsGetTextBox(AppID)
    If TextBox.TextFrame.TextRange.Font.Underline = msoTrue Then
        TextBox.TextFrame.TextRange.Font.Underline = msoFalse
    Else
        TextBox.TextFrame.TextRange.Font.Underline = msoTrue
    End If
    If AAX Then Slide1.AxTextBox.Font.Underline = MsoTristateToBool(TextBox.TextFrame.TextRange.Font.Underline)
End Sub

Sub AppWordsToggleStrikethrough(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim TextBox As Shape
    Set TextBox = AppWordsGetTextBox(AppID)
    If TextBox.TextFrame2.TextRange.Font.Strikethrough = msoTrue Then
        TextBox.TextFrame2.TextRange.Font.Strikethrough = msoFalse
    Else
        TextBox.TextFrame2.TextRange.Font.Strikethrough = msoTrue
    End If
    If AAX Then Slide1.AxTextBox.Font.Strikethrough = MsoTristateToBool(TextBox.TextFrame2.TextRange.Font.Strikethrough)
End Sub

Sub AppWordsInsertDate(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim TextBox As Shape
    Set TextBox = AppWordsGetTextBox(AppID)
    TextBox.TextFrame.TextRange.Text = TextBox.TextFrame.TextRange.Text & vbNewLine & Time & " " & Date
    If AAX Then Slide1.AxTextBox.Text = TextBox.TextFrame.TextRange.Text
End Sub

Sub AppWordsGrow(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim TextBox As Shape
    Set TextBox = AppWordsGetTextBox(AppID)
    TextBox.TextFrame.TextRange.Font.Size = TextBox.TextFrame.TextRange.Font.Size + 1
    Slide1.Shapes("SizeLabelAppWords:" & AppID).TextFrame.TextRange.Text = "Size:" & CStr(TextBox.TextFrame.TextRange.Font.Size)
    ApplyTbAttribs TextBox
End Sub

Sub AppWordsShrink(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim TextBox As Shape
    Set TextBox = AppWordsGetTextBox(AppID)
    TextBox.TextFrame.TextRange.Font.Size = TextBox.TextFrame.TextRange.Font.Size - 1
    Slide1.Shapes("SizeLabelAppWords:" & AppID).TextFrame.TextRange.Text = "Size:" & CStr(TextBox.TextFrame.TextRange.Font.Size)
    ApplyTbAttribs TextBox
End Sub

Sub AppWordsNormalSize(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim TextBox As Shape
    Set TextBox = AppWordsGetTextBox(AppID)
    TextBox.TextFrame.TextRange.Font.Size = 14
    Slide1.Shapes("SizeLabelAppWords:" & AppID).TextFrame.TextRange.Text = "Size:14"
    ApplyTbAttribs TextBox
End Sub