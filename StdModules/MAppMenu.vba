' Application launcher

Sub Menu()
    If Slide1.Shapes("Username").TextFrame.TextRange.Text = "Nobody" Then
        Logout
        Exit Sub
    End If
    DispApps
    If Hour(Time) < 12 And Hour(Time) > 4 Then
        Slide2.Shapes("GreetingAppMenu_").TextFrame.TextRange.Text = "Good morning, " & Slide1.Shapes("Username").TextFrame.TextRange.Text & "!"
    ElseIf Hour(Time) > 11 And Hour(Time) < 17 Then
        Slide2.Shapes("GreetingAppMenu_").TextFrame.TextRange.Text = "Good afternoon, " & Slide1.Shapes("Username").TextFrame.TextRange.Text & "!"
    Else
        Slide2.Shapes("GreetingAppMenu_").TextFrame.TextRange.Text = "Good evening, " & Slide1.Shapes("Username").TextFrame.TextRange.Text & "!"
    End If
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Menu"
    CopyFillFormat GetTheme().GroupItems("WindowFrame"), Slide2.Shapes("Shape5AppMenu_"), False
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    If AAX Then Slide1.AxTextBox.Visible = False
    MenuShowPage 1
End Sub

Sub AppMenu()
    Menu
End Sub


Sub MenuNextPage(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim NextPage As Integer
    NextPage = MenuGetPageNumber(AppID) + 1
    If MenuGetVisibleItems(AppID) < 18 Then Exit Sub
    MenuShowPage NextPage
End Sub

Function GetTheme() As Shape
    If FileStreamsExist("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm") Then
        Set GetTheme = GetFileRef("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm")
    Else
        Set GetTheme = GetFileRef("/Defaults/Themes/Default.thm")
    End If
End Function

Sub MenuLastPage(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim LastPage As Integer
    LastPage = MenuGetPageNumber(AppID) - 1
    If LastPage = 0 Then Exit Sub
    MenuShowPage LastPage
End Sub

Function MenuGetVisibleItems(AppID As String) As Integer
    For Each Shp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
        If InStr(1, Shp.Name, "Label") = 1 Then
            If Shp.Visible = msoTrue Then
                MenuGetVisibleItems = MenuGetVisibleItems + 1
            End If
        End If
    Next Shp
End Function

Function MenuGetPageNumber(AppID As String) As Integer
    Dim IDX As Integer
    Dim Shp As Shape
    MenuGetPageNumber = 1
    For IDX = 1 To 32767 Step 18
        Dim PageNum As Integer
        PageNum = ((IDX - 1) / 18) + 1
        For Each Shp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
            If InStr(1, Shp.Name, "Label" & IDX & "App") = 1 Then
                If Shp.Visible = msoTrue Then
                    MenuGetPageNumber = PageNum
                    Exit Function
                End If
            End If
        Next Shp
    Next IDX
End Function

Sub MenuShowPage(Page As Integer)
    Dim AppID As String
    Dim En As Integer
    Dim St As Integer
    Dim Shp As Shape
    AppID = Slide1.Shapes("AppID").TextFrame.TextRange.Text
    St = (Page - 1) * 18 + 1
    En = St + 17
    For J = Slide1.Shapes("RegularApp:" & AppID).GroupItems.Count To 1 Step -1
        Set Shp = Slide1.Shapes("RegularApp:" & AppID).GroupItems(J)
        If InStr(1, Shp.Name, "Icon") = 1 Or InStr(1, Shp.Name, "Label") = 1 Then
            Slide1.Shapes("RegularApp:" & AppID).GroupItems(J).Visible = msoFalse
        End If
    Next J
    For I = En To St Step -1
        For J = Slide1.Shapes("RegularApp:" & AppID).GroupItems.Count To 1 Step -1
            Set Shp = Slide1.Shapes("RegularApp:" & AppID).GroupItems(J)
            Dim II As String
            Dim LI As String
            II = "Icon" & I & "Part"
            LI = "Label" & I & "App"
            If (InStr(1, Shp.Name, II) = 1) Or (InStr(1, Shp.Name, LI) = 1) Then
                Slide1.Shapes("RegularApp:" & AppID).GroupItems(J).Visible = msoTrue
            End If
        Next J
    Next I
End Sub

Sub DispApps()
    ' A macro for drawing a list of applications to a test slide
    Dim guestSession As Boolean
    ' Set the slide, where the target menu is located at
    Dim Sld As Slide
    Set Sld = Slide2
    ' Ungroup the menu, so that we can work with it better
    If ShapeExists(Sld, "AppMenu") Then
        Sld.Shapes("AppMenu").Ungroup
    End If
    Limit = Sld.Shapes.Count
    ' Treverse list of shapes on slide backwards to avoid errors
    For IDX = Limit To 1 Step -1
        Dim Shp As Shape
        Set Shp = Sld.Shapes(IDX)
        If InStr(Shp.Name, "Icon") Then
            Shp.Delete
        ElseIf InStr(Shp.Name, "Label") Then
            Shp.Delete
        End If
    Next IDX
    ' Check if we're in a guest session
    guestSession = False
    If Slide1.Shapes("Username").TextFrame.TextRange.Text = "Guest" Then
        guestSession = True
    End If
    ' Define margins and positional offsets for menu icons
    Dim Left As Double
    Dim Top As Double
    Dim MarginX As Double
    Dim MarginY As Double
    Left = Sld.Shapes("BackgroundAppMenu_").Left + 10
    Top = Sld.Shapes("BackgroundAppMenu_").Top + 10
    'Dim PasteLeft As Integer
    'Dim PasteTop As Integer
    'Dim PasteWidth As Integer
    'Dim PasteHeight As Integer
    'With Sld.Shapes("Shape5AppMenu_")
    '    PasteLeft = .Left
    '    PasteTop = .Top
    '    PasteWidth = .Width
    '    PasteHeight = .Height
    '    .Delete
    'End With
    'GetTheme().GroupItems("WindowFrame").Copy
    'With Sld.Shapes.Paste
    '    .Name = "Shape5AppMenu_"
    '    .Left = PasteLeft
    '    .Top = PasteTop
    '    .Width = PasteWidth
    '    .Height = PasteHeight
    '    .Visible = msoTrue
    '    .Fill.Transparency = 0
    '    .ZOrder msoSendToBack
    'End With
    MarginX = 6
    MarginY = 40
    I = 1
    Dim IsFirstPage As Boolean
    IsFirstPage = True
    For Each Shp In Slide25.Shapes
        If InStr(Shp.Name, ":Properties") Then
            ' Full name of the app (e.g. AppCalc)
            FullName = Replace(Shp.Name, ":Properties", "")
            ' Application properties (access control and user friendly name)
            Props = Split(Shp.TextFrame.TextRange.Text, ":")
            Access = Props(0)
            ' Draw the icon only if access is to everyone, our username or we're not in a guest session
            If Access = "Everyone" Or guestSession = False Or Access = Slide1.Shapes("Username").TextFrame.TextRange.Text Then
                FriendlyName = Props(1)
                PackageName = Right(FullName, Len(FullName) - 3)
                ' Copy, Paste, Align the App Icon
                Slide25.Shapes("App" & PackageName & ":Icon").Copy
                Sld.Shapes.Paste
                Sld.Shapes("App" & PackageName & ":Icon").Left = Left
                Sld.Shapes("App" & PackageName & ":Icon").Top = Top
                Sld.Shapes("App" & PackageName & ":Icon").Visible = msoTrue
                Sld.Shapes("App" & PackageName & ":Icon").Name = "Icon" & CStr(I) & "AppMenu_"
                ' Same, but for App Label
                Slide25.Shapes("App" & PackageName & ":Properties").Copy
                Sld.Shapes.Paste
                Sld.Shapes("App" & PackageName & ":Properties").Left = Left
                Sld.Shapes("App" & PackageName & ":Properties").Top = Top + Sld.Shapes("Icon" & CStr(I) & "AppMenu_").Height
                Sld.Shapes("App" & PackageName & ":Properties").Width = Sld.Shapes("Icon" & CStr(I) & "AppMenu_").Width
                Sld.Shapes("App" & PackageName & ":Properties").Visible = msoTrue
                Sld.Shapes("App" & PackageName & ":Properties").TextFrame.TextRange.Text = FriendlyName
                Sld.Shapes("App" & PackageName & ":Properties").TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
                Sld.Shapes("App" & PackageName & ":Properties").TextFrame.TextRange.Font.Size = 8.3
                Sld.Shapes("App" & PackageName & ":Properties").TextFrame.TextRange.Font.Color.SchemeColor = ppForeground
                With Sld.Shapes("App" & PackageName & ":Properties").ActionSettings(ppMouseClick)
                   .Run = "App" & PackageName
                End With
                Sld.Shapes("App" & PackageName & ":Properties").Name = "Label" & CStr(I) & "AppMenu_"
                ' Set position for the next icon
                Left = Left + Sld.Shapes("Icon" & CStr(I) & "AppMenu_").Width + MarginX
                If I Mod 6 = 0 Then
                    ' If the index of current icon is divisible by 6, move it down a bit
                    Left = Sld.Shapes("BackgroundAppMenu_").Left + 10
                    Top = Top + Sld.Shapes("Icon" & CStr(I) & "AppMenu_").Height + MarginY
                End If
                ' Ungroup any subgroups to avoid errors later (rename all group children as well)
                sIdx = 1
                For Each Shp2 In Sld.Shapes("Icon" & CStr(I) & "AppMenu_").GroupItems
                    Shp2.Name = "Icon" & CStr(I) & "Part" & CStr(sIdx) & "AppMenu_"
                    sIdx = sIdx + 1
                Next Shp2
                Sld.Shapes("Icon" & CStr(I) & "AppMenu_").Ungroup
                I = I + 1
                If I Mod 19 = 0 Then
                    Left = Sld.Shapes("BackgroundAppMenu_").Left + 10
                    Top = Sld.Shapes("BackgroundAppMenu_").Top + 10
                    IsFirstPage = False
                End If
            End If
        End If
    Next Shp
    Sld.Shapes("ButtonNextAppMenu_").Fill.Transparency = 0
    Sld.Shapes("ButtonNextAppMenu_").TextFrame2.TextRange.Font.Fill.Transparency = 0
    Sld.Shapes("ButtonNextAppMenu_").ActionSettings(ppMouseClick).Action = ppActionNone
    Sld.Shapes("ButtonBackAppMenu_").Fill.Transparency = 0
    Sld.Shapes("ButtonBackAppMenu_").TextFrame2.TextRange.Font.Fill.Transparency = 0
    Sld.Shapes("ButtonBackAppMenu_").ActionSettings(ppMouseClick).Action = ppActionNone
    If I < 19 Then
        Sld.Shapes("ButtonNextAppMenu_").Fill.Transparency = 1
        Sld.Shapes("ButtonNextAppMenu_").TextFrame2.TextRange.Font.Fill.Transparency = 1
        Sld.Shapes("ButtonBackAppMenu_").Fill.Transparency = 1
        Sld.Shapes("ButtonBackAppMenu_").TextFrame2.TextRange.Font.Fill.Transparency = 1
    Else
        Sld.Shapes("ButtonNextAppMenu_").ActionSettings(ppMouseClick).Action = ppActionRunMacro
        Sld.Shapes("ButtonNextAppMenu_").ActionSettings(ppMouseClick).Run = "MenuNextPage"
        Sld.Shapes("ButtonBackAppMenu_").ActionSettings(ppMouseClick).Action = ppActionRunMacro
        Sld.Shapes("ButtonBackAppMenu_").ActionSettings(ppMouseClick).Run = "MenuLastPage"
    End If
    ' Regroup everything back
    Dim groupableShapes As String
    groupableShapes = ""
    For Each Shp In Sld.Shapes
        If Right(Shp.Name, Len("AppMenu_")) = "AppMenu_" Then
            groupableShapes = groupableShapes & ":" & Shp.Name
        End If
    Next Shp
    ' Removes ":" from the beginning
    groupableShapes = Right(groupableShapes, Len(groupableShapes) - 1)
    With Sld.Shapes.Range(Split(groupableShapes, ":")).Group
        .Name = "AppMenu"
    End With
    ' Hide only the group and not the group items
    ' Due to the weirdness of how PowerPoint works, we first need to hide the group,
    ' then go through each group item and set them to be visible
    Sld.Shapes("AppMenu").Visible = msoFalse
    For Each Shp In Sld.Shapes("AppMenu").GroupItems
        Shp.Visible = msoTrue
    Next Shp
End Sub

