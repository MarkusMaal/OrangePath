Public Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As LongPtr
Public disableClock As Boolean


' OrangePath/OS system module

Sub Pause(Length As LongPtr)
    Dim NowTime As LongPtr
    Dim EndTime As LongPtr

    EndTime = GetTickCount + (Length * 1000)
    Do
        NowTime = GetTickCount
        DoEvents
    Loop Until NowTime >= EndTime
End Sub

Sub ExitAppMenu()
    ActivePresentation.SlideShowWindow.View.GotoSlide (3)
    UpdateTime
End Sub


Sub CreateNewWindow()
    ' Variables
    'ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    Slide1.Shapes("AppID").TextFrame.TextRange.Text = CStr(CInt(Slide1.Shapes("AppID").TextFrame.TextRange.Text) + 1)
    AppID = Slide1.Shapes("AppID").TextFrame.TextRange.Text
    Dim Shp As shape
    ' Skip creating taskbar icon or any other nonsense if we're just opening the launch menu
    If Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Menu" Then GoTo ModalSkip
    If Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Message" Then GoTo ModalSkip
    If Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "InputBox" Then GoTo ModalSkip
    If InStr(1, Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text, "Modal") = 1 Then GoTo ModalSkip
    
    'Taskbar icon
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 4 Then
    
        LoopCount = 0
        Do
            TaskIconCount = 0
            Offset = 0
            For Each Shp In Slide1.Shapes()
                If InStr(Shp.Name, "TaskIcon:") And Shp.Left > 0 And Shp.Left < ActivePresentation.SlideShowWindow.Width Then
                    Offset = Offset + Shp.Width
                    TaskIconCount = TaskIconCount + 1
                End If
            Next
            If TaskIconCount = 5 Then
                If LoopCount < 4 Then
                    CheckSize Slide1.Shapes("SwitchWorkspace")
                    LoopCount = LoopCount + 1
                Else
                    AppMessage "Maximum number of windows reached", "Cannot create window", "Error", True
                    Exit Sub
                End If
            Else
                Exit Do
            End If
        Loop
        FocusWindow AppID
    
        Slide1.Shapes("TaskbarButtonSample").Copy
        With Slide1.Shapes.Paste
        .Name = "TaskIcon:" & AppID
        .Left = .Left + Offset
        .Top = Slide1.Shapes("Taskbar").Top
        .TextFrame.TextRange.Text = Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text
        If CheckVars("%Animations%") <> "True" Then
            .Visible = msoTrue
        End If
        End With
    End If
ModalSkip:
    FocusWindow AppID
    ' Scripts
    Slide2.Shapes("App" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text).Copy
    currentSlide = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    With ActivePresentation.Slides(currentSlide).Shapes.Paste
    .Name = "RegularApp:" & AppID
    .Visible = msoTrue
    End With
    ActivePresentation.Slides(currentSlide).Shapes("RegularApp:" & AppID).Ungroup
    
    For Each Shp In ActivePresentation.Slides(currentSlide).Shapes()
        If InStr(Shp.Name, "_") Then
            If CheckVars("%Animations%") = "True" Then
                Shp.Visible = msoFalse
            End If
            SplitName = Split(Shp.Name, "_")
            Shp.Name = SplitName(0) & ":" & AppID
            If InStr(Shp.Name, "AXTextBox") Then ApplyTbAttribs Shp
            Shapes = Shapes & Shp.Name & ","
        End If
    Next
    
    SplitShapes = Split(Shapes, ",")
    UJ = CInt(UBound(SplitShapes))
    Dim ShapesX() As String
    
    ReDim ShapesX(UJ)
    For i = 0 To CInt(UBound(SplitShapes) - 1)
        CShape = SplitShapes(i)
        If Not IsInArray(CStr(CShape), ShapesX) Then
            ShapesX(i) = SplitShapes(i)
        End If
    Next
    
    With ActivePresentation.Slides(currentSlide).Shapes.Range(ShapesX).Group
        .Name = "RegularApp:" & AppID
    End With
    On Error Resume Next
    If Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text <> "Menu" Then ActivePresentation.Slides(currentSlide).Shapes("TaskIcon:" & AppID).Visible = msoTrue

    ' Special case for Taskmgr
    RefreshTaskmgrs
    If Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Taskmgr" Then
        TaskmgrRefresh Slide1.Shapes("RegularApp:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text)
    End If
    Dim eff As Effect
    For i = ActivePresentation.Slides(currentSlide).TimeLine.MainSequence.Count To 1 Step -1
        Set eff = Slide1.TimeLine.MainSequence(i)
        eff.Delete
    Next i
    ActivePresentation.SlideShowWindow.View.GotoSlide (currentSlide)
    If CheckVars("%Animations%") = "True" Then
        Dim oeff As Effect
        Set oeff = ActivePresentation.Slides(currentSlide).TimeLine.MainSequence.AddEffect(shape:=ActivePresentation.Slides(currentSlide).Shapes("RegularApp:" & AppID), effectId:=msoAnimEffectStrips, trigger:=msoAnimTriggerAfterPrevious)
        oeff.Exit = msoFalse
        oeff.EffectParameters.Direction = msoAnimDirectionBottomRight
        oeff.Timing.Duration = 0.5
        ActivePresentation.Slides(currentSlide).Shapes("RegularApp:" & AppID).Visible = msoTrue
        ActivePresentation.SlideShowWindow.Activate
        Pause 2
        For i = ActivePresentation.Slides(currentSlide).TimeLine.MainSequence.Count To 1 Step -1
            Set eff = Slide1.TimeLine.MainSequence(i)
            eff.Delete
        Next i
        ActivePresentation.Slides(currentSlide).Shapes("RegularApp:" & AppID).Visible = msoFalse
        ActivePresentation.SlideShowWindow.Activate
        ActivePresentation.Slides(currentSlide).Shapes("RegularApp:" & AppID).Visible = msoTrue
        ActivePresentation.SlideShowWindow.Activate
        ActivePresentation.SlideShowWindow.View.GotoSlide (currentSlide)
    End If
End Sub

Public Function IsInArray(stringToBeFound As String, arr() As String) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Sub RefreshTaskmgrs()
    For Each Grp In Slide1.Shapes
        If InStr(Grp.Name, "RegularApp:") Then
            On Error Resume Next
            TaskmgrRefresh Slide1.Shapes(Grp.Name)
        End If
    Next Grp
End Sub

Sub CloseWindow(Shp As shape)
    AppID = 0
    If InStr(Shp.TextFrame.TextRange.Text, "PID") Then
        SplitA = Split(Shp.TextFrame.TextRange.Text, ":")
        StringA = SplitA(UBound(SplitA))
        SplitB = Split(StringA, " ")
        StringB = SplitB(1)
        SplitC = Split(StringB, ")")
        AppID = SplitC(0)
        If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 4 Then
            For Each S In Slide1.Shapes
                sName = Split(S.Name, ":")
                sId = sName(UBound(sName))
                If sId = AppID Then
                    Animate AppID, ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition)
                    If CheckVars("%Animations%") = "True" Then Pause 2
                    S.Delete
                End If
            Next S
            If ShapeExists(ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition), "TaskIcon:" & AppID) Then
                TaskIcon = CInt(Slide1.Shapes("TaskIcon:" & AppID).Left)
                ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition).Shapes("TaskIcon:" & AppID).Delete
                
                RefreshTaskmgrs
                If TaskIcon = 338 Then
                    TaskIcon = 3
                ElseIf TaskIcon = 474 Then
                    TaskIcon = 4
                ElseIf TaskIcon = 202 Then
                    TaskIcon = 2
                ElseIf TaskIcon = 65 Then
                    TaskIcon = 1
                Else
                    TaskIcon = 5
                End If
                ReorganizeTaskIcons TaskIcon
            End If
        End If
    Else
        CheckActiveX Shp
        SplitZ = Split(Shp.ParentGroup.Name, ":")
        AppID = SplitZ(1)
        Animate AppID, ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition)
        If CheckVars("%Animations%") = "True" Then Pause 2
        Shp.ParentGroup.Delete
        RefreshTaskmgrs
        If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 4 Then
            If ShapeExists(ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition), "TaskIcon:" & AppID) Then
                TaskIcon = CInt(Slide1.Shapes("TaskIcon:" & AppID).Left)
                ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition).Shapes("TaskIcon:" & AppID).Delete
                If TaskIcon = 338 Then
                    TaskIcon = 3
                ElseIf TaskIcon = 474 Then
                    TaskIcon = 4
                ElseIf TaskIcon = 202 Then
                    TaskIcon = 2
                ElseIf TaskIcon = 65 Then
                    TaskIcon = 1
                Else
                    TaskIcon = 5
                End If
                ReorganizeTaskIcons TaskIcon
            End If
        End If
    End If
End Sub


Sub Animate(ByVal AppID As String, ByVal Sld As slide)
    If CheckVars("%Animations%") = "True" Then
        Dim eff As Effect
        For i = Slide1.TimeLine.MainSequence.Count To 1 Step -1
            Set eff = Slide1.TimeLine.MainSequence(i)
            eff.Delete
        Next i
        Dim oeff As Effect
        Set oeff = Sld.TimeLine.MainSequence.AddEffect(shape:=Sld.Shapes("RegularApp:" & AppID), effectId:=msoAnimEffectStrips, trigger:=msoAnimTriggerAfterPrevious)
        oeff.Exit = msoTrue
        oeff.EffectParameters.Direction = msoAnimDirectionTopLeft
        oeff.Timing.Duration = 0.5
        ActivePresentation.SlideShowWindow.Activate
    End If
End Sub

Sub CloseTest()
    AppID = 18
    Dim Shp As shape
    Set Shp = Slide1.Shapes("CloseAppCalc:18")
    If Shp.TextFrame.TextRange.Text = "X" Then
        MsgBox "X"
        CheckActiveX Shp
        SplitZ = Split(Shp.ParentGroup.Name, ":")
        AppID = SplitZ(1)
        Shp.ParentGroup.Delete
        If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 4 Then
            ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition).Shapes("TaskIcon:" & AppID).Delete
            
            TaskIcon = CInt(Slide1.Shapes("TaskIcon:" & AppID).Left)
            If TaskIcon = 338 Then
                TaskIcon = 3
            ElseIf TaskIcon = 474 Then
                TaskIcon = 4
            ElseIf TaskIcon = 202 Then
                TaskIcon = 2
            ElseIf TaskIcon = 65 Then
                TaskIcon = 1
            Else
                TaskIcon = 5
            End If
            ReorganizeTaskIcons TaskIcon
        End If
    ElseIf InStr(Shp.TextFrame.TextRange.Text, "PID") Then
        MsgBox "PIDCLOSE"
        SplitA = Split(Shp.TextFrame.TextRange.Text, ":")
        StringA = SplitA(UBound(SplitA))
        SplitB = Split(StringA, " ")
        StringB = SplitB(1)
        SplitC = Split(StringB, ")")
        AppID = SplitC(0)
        If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 4 Then
            ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition).Shapes("TaskIcon:" & AppID).Delete
            For Each S In Slide1.Shapes
                sName = Split(S.Name, ":")
                sId = sName(UBound(sName))
                If sId = AppID Then
                    S.Delete
                End If
            Next S
            TaskmgrRefresh Shp
            
            TaskIcon = CInt(Slide1.Shapes("TaskIcon:" & AppID).Left)
            If TaskIcon = 338 Then
                TaskIcon = 3
            ElseIf TaskIcon = 474 Then
                TaskIcon = 4
            ElseIf TaskIcon = 202 Then
                TaskIcon = 2
            ElseIf TaskIcon = 65 Then
                TaskIcon = 1
            Else
                TaskIcon = 5
            End If
            ReorganizeTaskIcons TaskIcon
        End If
    End If
End Sub


Sub FixOrder(ByVal AppID As String)
    TaskIcon = CInt(Slide1.Shapes("TaskIcon:" & AppID).Left)
    If TaskIcon = 338 Then
        TaskIcon = 3
    ElseIf TaskIcon = 474 Then
        TaskIcon = 4
    ElseIf TaskIcon = 202 Then
        TaskIcon = 2
    ElseIf TaskIcon = 65 Then
        TaskIcon = 1
    Else
        TaskIcon = 5
    End If
    ReorganizeTaskIcons TaskIcon
End Sub

Sub CheckActiveX(ByVal Shp As shape)
    For Each SubShp In Shp.ParentGroup.GroupItems
        If InStr(SubShp.Name, "AXTextBox") Then
            Slide1.AxTextBox.Visible = False
            If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then Slide13.AxTextBox.Visible = False
        End If
    Next SubShp
End Sub

Sub CheckActiveXShow(Shp As shape)
    Dim SubShp As shape
    For Each SubShp In Shp.ParentGroup.GroupItems
        If InStr(SubShp.Name, "AXTextBox") Then ApplyTbAttribs SubShp
    Next SubShp
End Sub

Sub MovableWindow(Shp As shape)
'Sub MovableWindow()
    ' If moving, exit sub
    If Slide1.Shapes("MoveEvent").TextFrame.TextRange.Text = "True" Then
        Exit Sub
    End If
    
    Dim LoopState As Boolean
    SplitZ = Split(Shp.Name, ":")
    AppID = SplitZ(1)
    For i = Slide1.TimeLine.MainSequence.Count To 1 Step -1
        Set oeff = Slide1.TimeLine.MainSequence(i)
        oeff.Delete
    Next i
    
    Do
        currentSlide = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
        GetCursorPositionX1 = GetCursorX
        GetCursorPositionY1 = GetCursorY
        GetAsyncKeyState1 = GetAsyncKeyState(1)
        ' TitleBar clicked, window moving
        If LoopState And GetAsyncKeyState1 Then
            ActivePresentation.Slides(currentSlide).Shapes("RegularApp:" & AppID).Top = GetCursorPositionY1 - dy
            ActivePresentation.Slides(currentSlide).Shapes("RegularApp:" & AppID).Left = GetCursorPositionX1 - dx
            ActivePresentation.Slides(currentSlide).Shapes("RegularApp:" & AppID).ZOrder msoBringToFront
            Slide1.AxTextBox.Visible = False
            If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then Slide13.AxTextBox.Visible = False
            
            ActivePresentation.SlideShowWindow.View.GotoSlide (currentSlide)
        ' TitleBar clicked, window hasn't moved
        ElseIf LoopState = False And GetAsyncKeyState1 Then
            dx = GetCursorPositionX1 - ActivePresentation.Slides(currentSlide).Shapes("RegularApp:" & AppID).Left
            dy = GetCursorPositionY1 - ActivePresentation.Slides(currentSlide).Shapes("RegularApp:" & AppID).Top - 5
            Slide1.Shapes("MoveEvent").TextFrame.TextRange.Text = "True"
        ' TitleBar not clicked, window was moving
        ElseIf LoopState = True And GetAsyncKeyState1 = False Then
            Slide1.Shapes("MoveEvent").TextFrame.TextRange.Text = "False"
            If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 4 Then FocusWindow AppID
            For Each Shp In ActivePresentation.Slides(currentSlide).Shapes("RegularApp:" & AppID).GroupItems
                If InStr(Shp.Name, "AXTextBox") Then ApplyTbAttribs Shp
            Next Shp
            ActivePresentation.SlideShowWindow.View.GotoSlide (currentSlide)
            Exit Sub
        End If
        DoEvents
        LoopState = GetAsyncKeyState1
    Loop
End Sub

Sub SetTextBox()
    With Slide1.AxTextBox
        .Visible = True
        .Left = Slide1.Shapes("AxTextBox1AppNotes:21").Left
        .Top = Slide1.Shapes("AxTextBox1AppNotes:21").Top
        .Width = Slide1.Shapes("AxTextBox1AppNotes:21").Width
        .Height = Slide1.Shapes("AxTextBox1AppNotes:21").Height
        .Text = Slide1.Shapes("AxTextBox1AppNotes:21").TextFrame.TextRange.Text
    End With
End Sub

Sub SetTextBoxVal(Shp As shape)
    currentSlide = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    If currentSlide <> 13 Then
        Shp.TextFrame.TextRange.Text = Slide1.AxTextBox.Text
    Else
        Shp.TextFrame.TextRange.Text = Slide13.AxTextBox.Text
    End If
End Sub

Function ShapeExists(ByVal oSl As slide, ByVal ShapeName As String) As Boolean
   Dim oSh As shape
   For Each oSh In oSl.Shapes
     If oSh.Name = ShapeName Then
        ShapeExists = True
        Exit Function
     End If
   Next ' Shape
   ' No shape here, so though it's not strictly necessary:
   ShapeExists = False
End Function

Function GroupItemExists(ByVal oSl As shape, ByVal ShapeName As String) As Boolean
   Dim oSh As shape
   For Each oSh In oSl.GroupItems
     If oSh.Name = ShapeName Then
        GroupItemExists = True
        Exit Function
     End If
   Next ' Shape
   ' No shape here, so though it's not strictly necessary:
   GroupItemExists = False
End Function


Sub ResizingWindow(Shp As shape)
    ' If reszing, exit sub
    If Slide1.Shapes("ResizeEvent").TextFrame.TextRange.Text = "True" Then
       Exit Sub
    End If
    
    Dim LoopState As Boolean
    SplitZ = Split(Shp.Name, ":")
    AppID = SplitZ(1)
    SplitN = Split(Shp.Name, "App")
    SplitO = Split(SplitN(1), ":")
    AppName = SplitO(0)
    Dim O As shape
    Set O = Slide1.Shapes("WindowApp" & AppName & ":" & AppID)
    Do
        GetCursorPositionX1 = GetCursorX
        GetCursorPositionY1 = GetCursorY
        GetAsyncKeyState1 = GetAsyncKeyState(1)
        ' GrabArea clicked, window resizing
        If LoopState And GetAsyncKeyState1 Then
            Slide1.Shapes("RegularApp:" & AppID).Height = GetCursorPositionY1 - dy
            Slide1.Shapes("RegularApp:" & AppID).Width = GetCursorPositionX1 - dx
            Slide1.Shapes("RegularApp:" & AppID).ZOrder msoBringToFront
            Slide1.AxTextBox.Visible = False
            If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then Slide13.AxTextBox.Visible = False
            ActivePresentation.SlideShowWindow.View.GotoSlide (4)
        ' GrabArea clicked, window hasn't resized
        ElseIf LoopState = False And GetAsyncKeyState1 Then
            dx = Slide1.Shapes("RegularApp:" & AppID).Left
            dy = Slide1.Shapes("RegularApp:" & AppID).Top
            Slide1.Shapes("ResizeEvent").TextFrame.TextRange.Text = "True"
        ' GrabArea not clicked, window was resized
        ElseIf LoopState = True And GetAsyncKeyState1 = False Then
            Slide1.Shapes("ResizeEvent").TextFrame.TextRange.Text = "False"
            FocusWindow AppID
            For Each Shp In Slide1.Shapes("RegularApp:" & AppID).GroupItems
                If InStr(Shp.Name, "AXTextBox") Then ApplyTbAttribs Shp
            Next Shp
            ActivePresentation.SlideShowWindow.View.GotoSlide (4)
            With Slide1.Shapes("RegularApp:" & AppID)
                For i = Slide1.Shapes("RegularApp:" & AppID).GroupItems.Count To 0 Step -1
                    With .GroupItems(i)
                        If InStr(.Name, "+") Then
                            NameSplit = Split(.Name, "+")
                            ReferenceShape = NameSplit(0)
                            Slide1.Shapes(.Name).Width = Slide1.Shapes(ReferenceShape).Width
                            Slide1.Shapes(.Name).Height = Slide1.Shapes(ReferenceShape).Height
                            Slide1.Shapes(.Name).Left = Slide1.Shapes(ReferenceShape).Left
                            Slide1.Shapes(.Name).Top = Slide1.Shapes(ReferenceShape).Top + Slide1.Shapes(ReferenceShape).Height + 1
                        ElseIf InStr(.Name, "*") Then
                            SplitName = Split(.Name, "*")
                            Anchor = SplitName(2)
                            PreviousWidth = Slide1.Shapes(.Name).Width
                            PreviousHeight = Slide1.Shapes(.Name).Height
                            NewPosX = Slide1.Shapes(.Name).Left
                            NewPosY = Slide1.Shapes(.Name).Top
                            NewWidth = CDbl(SplitName(0))
                            NewHeight = CDbl(SplitName(1))
                            If InStr(Anchor, "S") Then
                                NewPosY = (NewPosY + PreviousHeight) - NewHeight
                            End If
                            If InStr(Anchor, "E") Then
                                NewPosX = (NewPosX + PreviousWidth) - NewWidth
                            End If
                            Slide1.Shapes(.Name).Width = CDbl(NewWidth)
                            Slide1.Shapes(.Name).Height = CDbl(NewHeight)
                            Slide1.Shapes(.Name).Top = CDbl(NewPosY)
                            Slide1.Shapes(.Name).Left = CDbl(NewPosX)
                        ElseIf InStr(.Name, "Close") Then
                            Slide1.Shapes(.Name).Width = 42.5
                            Slide1.Shapes(.Name).Height = 22.7
                            Slide1.Shapes(.Name).Left = Slide1.Shapes("RegularApp:" & AppID).Left + Slide1.Shapes("RegularApp:" & AppID).Width - .Width - 11.7
                        ElseIf InStr(.Name, "Minimize") Then
                            Slide1.Shapes(.Name).Width = 42.5
                            Slide1.Shapes(.Name).Height = 22.7
                            Slide1.Shapes(.Name).Left = Slide1.Shapes("RegularApp:" & AppID).Left + Slide1.Shapes("RegularApp:" & AppID).Width - (2 * .Width) - 11.7
                        End If
                    End With
                Next
                For i = Slide1.Shapes("RegularApp:" & AppID).GroupItems.Count To 0 Step -1
                    With .GroupItems(i)
                        If InStr(.Name, "+") Then
                            NameSplit = Split(.Name, "+")
                            ReferenceShape = NameSplit(0)
                            Slide1.Shapes(.Name).Width = Slide1.Shapes(ReferenceShape).Width
                            Slide1.Shapes(.Name).Height = Slide1.Shapes(ReferenceShape).Height
                            Slide1.Shapes(.Name).Left = Slide1.Shapes(ReferenceShape).Left
                            Slide1.Shapes(.Name).Top = Slide1.Shapes(ReferenceShape).Top + Slide1.Shapes(ReferenceShape).Height + 1
                        End If
                    End With
                Next
            End With
            Exit Sub
        End If
        DoEvents
        LoopState = GetAsyncKeyState1
    Loop
End Sub

Sub GetScale()
    MsgBox Slide2.Shapes("Shape8AppPaint_").Width & " " & Slide2.Shapes("Shape8AppPaint_").Height
    'MsgBox Slide2.Shapes("CloseAppSettings_").Width & " " & Slide2.Shapes("CloseAppSettings_").Height
End Sub


Sub ApplyTbAttribs(Shp As shape)
    AppID = GetAppID(Shp)
    With Slide1.AxTextBox
        .Left = Shp.Left
        .Top = Shp.Top
        .Width = Shp.Width
        .Height = Shp.Height
        .Text = Shp.TextFrame.TextRange.Text
        If InStr(Shp.Name, "Shell") Then
            .ScrollBars = fmScrollBarsNone
            .EnterKeyBehavior = False
            .BackColor = RGB(0, 0, 0)
            .ForeColor = RGB(254, 254, 254)
        ElseIf InStr(Shp.Name, "Notes") Then
            .ScrollBars = fmScrollBarsBoth
            .EnterKeyBehavior = True
            .ForeColor = Slide1.Shapes("WindowAppNotes:" & AppID).TextFrame.TextRange.Font.Color.RGB
            .BackColor = Slide1.Shapes("WindowAppNotes:" & AppID).TextFrame2.TextRange.Font.Line.ForeColor.RGB
        End If
        .Visible = True
    End With
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then
        With Slide13.AxTextBox
            .Left = Shp.Left
            .Top = Shp.Top
            .Width = Shp.Width
            .Height = Shp.Height
            .Text = Shp.TextFrame.TextRange.Text
            .Visible = True
        End With
    End If
End Sub


Sub CheckSize(Shp As shape)
    Slide1.AxTextBox.Visible = False
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then Slide13.AxTextBox.Visible = False
    If Shp.TextFrame.TextRange.Text = "Workspace 1" Then
        For Each Shp In Slide1.Shapes
            If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                AppID = GetAppID(Shp)
                If Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.4 Then
                    FocusWindow (AppID)
                End If
            End If
        Next Shp
        Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 2"
    ElseIf Shp.TextFrame.TextRange.Text = "Workspace 2" Then
        For Each Shp In Slide1.Shapes
            If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                AppID = GetAppID(Shp)
                If Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.4 Then
                    FocusWindow (AppID)
                End If
            End If
        Next Shp
        Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 3"
    ElseIf Shp.TextFrame.TextRange.Text = "Workspace 3" Then
        For Each Shp In Slide1.Shapes
            If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                AppID = GetAppID(Shp)
                If Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.4 Then
                    FocusWindow (AppID)
                End If
            End If
        Next Shp
        Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 4"
    Else
        For Each Shp In Slide1.Shapes
            If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                Shp.Left = Shp.Left + ActivePresentation.SlideShowWindow.Width * 3
                AppID = GetAppID(Shp)
                If Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.4 Then
                    FocusWindow (AppID)
                End If
            End If
        Next Shp
        Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 1"
    End If
End Sub


Sub Hibernate()
    Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "True"
    Slide3.Shapes("BootText").TextFrame.TextRange.Text = "Resuming from hibernation..."
    Slide2.Shapes("BootText").TextFrame.TextRange.Text = "Window templates"
    Slide3.Shapes("Bootlogo").Visible = msoTrue
    Slide3.Shapes("Bootwarning").Visible = msoFalse
    SavePresentation
    ActivePresentation.SlideShowWindow.View.Exit
End Sub

Sub MacroTest()
    Slide8.Shapes("BootText").TextFrame.TextRange.Text = "Macro test success"
    Pause (1)
    Slide8.Shapes("BootText").TextFrame.TextRange.Text = "Recovery mode"
End Sub

Sub RecoverSession()
    Dim CanRecover As Boolean
    For Each oShp In Slide1.Shapes
        If InStr(oShp.Name, "RegularApp:") Then
            CanRecover = True
        End If
    Next
    If CanRecover Then
        Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "True"
        Slide3.Shapes("BootText").TextFrame.TextRange.Text = "Attempting session recovery..."
        Slide3.Shapes("Bootlogo").Visible = msoTrue
        Slide3.Shapes("Bootwarning").Visible = msoFalse
        ActivePresentation.SlideShowWindow.View.GotoSlide (1)
    Else
        Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "False"
        Slide3.Shapes("BootText").TextFrame.TextRange.Text = "No recoverable data found, booting normally..."
        Slide3.Shapes("Bootlogo").Visible = msoTrue
        Slide3.Shapes("Bootwarning").Visible = msoFalse
        ActivePresentation.SlideShowWindow.View.GotoSlide (1)
    End If
End Sub

Sub TaskmgrRefresh(ByVal Shp As shape)
    AppID = GetAppID(Shp)
    TaskMgrID = AppID
    With Slide1.Shapes("RegularApp:" & TaskMgrID)
        For Each GI In .GroupItems
            If InStr(GI.Name, "Proc") And GI.TextFrame.TextRange.Text <> "" Then
                GI.TextFrame.TextRange.Text = ""
                With GI.ActionSettings(ppMouseClick)
                    .Action = ppActionNone
                End With
            End If
        Next
    End With
    i = 1
    For Each oShp In Slide1.Shapes
        If InStr(oShp.Name, "RegularApp:") Then
            SplitZ = Split(oShp.Name, ":")
            AppID = SplitZ(1)
            With Slide1.Shapes("RegularApp:" & AppID)
                If .GroupItems.Count > 0 Then
                    AppNameSplit = Split(.GroupItems(1).Name, ":")
                    AppNameSplit2 = Split(AppNameSplit(0), "App")
                    AppName = AppNameSplit2(1)
                    If i < 11 Then
                        Dim hp As Boolean
                        hp = False
                        With Slide1.Shapes("RegularApp:" & TaskMgrID)
                            For Each GI In .GroupItems
                                If InStr(GI.Name, "Proc") And GI.TextFrame.TextRange.Text = "" And Not hp Then
                                    GI.TextFrame.TextRange.Text = AppName & " (PID: " & CStr(AppID) & ")"
                                    i = i + 1
                                    With GI.ActionSettings(ppMouseClick)
                                        .Run = "CloseWindow"
                                    End With
                                    hp = True
                                End If
                            Next
                        End With
                    End If
                End If
            End With
        End If
    Next
End Sub


Sub OnSlideShowPageChange(ByVal oSW As SlideShowWindow)
    'On Error GoTo ReportIssue
    PageChange oSW
    Exit Sub
ReportIssue:
    Slide22.Shapes("Details").TextFrame.TextRange.Text = "Error details: " & Err.Description
    ActivePresentation.SlideShowWindow.View.GotoSlide 22
End Sub


Sub Slide2Run()
    Slide1.Shapes("MoveEvent").TextFrame.TextRange.Text = "False"
    Slide1.Shapes("ResizeEvent").TextFrame.TextRange.Text = "False"
    Slide1.Shapes("AppID").TextFrame.TextRange.Text = "1"
    Slide1.Shapes("Username").TextFrame.TextRange.Text = "Nobody"
    DeleteDir "/Temp/"
    NewFolder "/Temp"
    Slide1.AxTextBox.Visible = False
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then Slide13.AxTextBox.Visible = False
    ResetWindows
    Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 1"
    Dim Factory As Boolean
    Dim Shp As shape
    Factory = Not FileStreamsExist("/System/Settings.cnf")
    If Factory Then
        ActivePresentation.SlideShowWindow.View.GotoSlide (18)
    Else
        If FileExists("/System/Settings.cnf", "Autologin") Then
            Slide1.Shapes("Username").TextFrame.TextRange.Text = GetFileContent("/System/Settings.cnf", "Autologin")
            If Slide1.Shapes("Username").TextFrame.TextRange.Text <> "Nobody" Then
                InitUserspace
            Else
                ActivePresentation.SlideShowWindow.View.GotoSlide 12
            End If
        Else
            Slide22.Shapes("Details").TextFrame.TextRange.Text = "Error details: SYSTEM_CONFIGURATION_IS_CORRUPT"
            ActivePresentation.SlideShowWindow.View.GotoSlide 22
        End If
    End If
    Slide7.Shapes("BootText").TextFrame.TextRange.Text = "Shutting down..."
    Slide2.Shapes("BootText").TextFrame.TextRange.Text = "Shutting down..."
End Sub

Sub PageChange(ByVal oSW As SlideShowWindow)
    If Slide3.Shapes("ErrorTest").TextFrame.TextRange.Text = "True" Then
        Exit Sub
    End If
    If oSW.View.CurrentShowPosition <> 13 Then
        Slide13.UsernameFIeld.Text = ""
        Slide13.PasswordField.Text = ""
    End If
    If oSW.View.CurrentShowPosition = 12 Then
        InitLogon
    End If
    CleanPopups
    Unlight
    If oSW.View.CurrentShowPosition = 1 Then
        Dim buildNo As String
        Slide3.Shapes("Wdym").Visible = msoFalse
        Slide3.Shapes("RecoveryModeBtn").Visible = msoFalse
        buildNo = "953 *EVERYTHING BREAKS*"
        Slide5.Shapes("MacroTest").TextFrame.TextRange.Text = "Macros enabled"
        Slide1.Shapes("TextBox 2").TextFrame.TextRange.Text = "Codename OrangePath/OS" + vbNewLine + "Build " + buildNo + vbNewLine + "For evaluation purposes only"
        Slide3.Shapes("TextBox 5").TextFrame.TextRange.Text = "Codename OrangePath/OS" + vbNewLine + "Build " + buildNo + vbNewLine + "For evaluation purposes only"
        Slide5.Shapes("TextBox 8").TextFrame.TextRange.Text = "Codename OrangePath/OS" + vbNewLine + "Build " + buildNo + vbNewLine + "For evaluation purposes only"
        Slide7.Shapes("TextBox 4").TextFrame.TextRange.Text = "Codename OrangePath/OS" + vbNewLine + "Build " + buildNo + vbNewLine + "For evaluation purposes only"
        Slide22.Shapes("Details").TextFrame.TextRange.Text = " "
        If InStr(Application.ActivePresentation.Name, "ForceRecovery") Then
            Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "Recovery"
        End If
        Slide2.Shapes("BootText").TextFrame.TextRange.Text = "Shutting down..."
        If Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "True" Then
            Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "False"
            ActivePresentation.SlideShowWindow.View.GotoSlide (4)
            UpdateTime
            Slide3.Shapes("BootText").TextFrame.TextRange.Text = "Macros not enabled"
            Slide5.Shapes("BootText").TextFrame.TextRange.Text = "Macros not enabled"
            Slide3.Shapes("Wdym").Visible = msoTrue
            Slide3.Shapes("RecoveryModeBtn").Visible = msoTrue
            Slide3.Shapes("Bootwarning").Visible = msoTrue
            Slide3.Shapes("Bootlogo").Visible = msoFalse
            Exit Sub
        ElseIf Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "Recovery" Then
            Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "False"
            ActivePresentation.SlideShowWindow.View.GotoSlide (8)
            Exit Sub
        Else
            Slide3.Shapes("BootText").TextFrame.TextRange.Text = "Starting OrangePath/OS"
            Slide5.Shapes("BootText").TextFrame.TextRange.Text = "Starting OrangePath/OS"
            Slide3.Shapes("Bootwarning").Visible = msoFalse
            Slide3.Shapes("Bootlogo").Visible = msoTrue
            With Slide5.Shapes("Windows Flora Startup").AnimationSettings
                .AdvanceMode = ppAdvanceOnTime
                .PlaySettings.PlayOnEntry = True
                .PlaySettings.PauseAnimation = False
                .PlaySettings.StopAfterSlides = 999
            End With
        End If
    ElseIf oSW.View.CurrentShowPosition = 2 Then
        Slide5.Shapes("MacroTest").TextFrame.TextRange.Text = "Macros disabled"
        Slide3.Shapes("Wdym").Visible = msoFalse
        Slide3.Shapes("RecoveryModeBtn").Visible = msoFalse
        Slide2Run
    ElseIf oSW.View.CurrentShowPosition = 3 Then
        Slide3.Shapes("BootText").TextFrame.TextRange.Text = "Macros not enabled"
        Slide5.Shapes("BootText").TextFrame.TextRange.Text = "Macros not enabled"
        Slide3.Shapes("Wdym").Visible = msoTrue
        Slide3.Shapes("RecoveryModeBtn").Visible = msoTrue
        Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "Recovery"
        Slide8.Shapes("BootText").TextFrame.TextRange.Text = "System was not shut down correctly. How would you like to proceed?"
        Slide3.Shapes("Bootwarning").Visible = msoTrue
        Slide3.Shapes("Bootlogo").Visible = msoFalse
        Dim oeff As Effect
        For i = Slide3.TimeLine.MainSequence.Count To 1 Step -1
            Set oeff = Slide3.TimeLine.MainSequence(i)
            oeff.Delete
        Next i
        For i = Slide6.TimeLine.MainSequence.Count To 1 Step -1
            Set oeff = Slide3.TimeLine.MainSequence(i)
            oeff.Delete
        Next i
    ElseIf oSW.View.CurrentShowPosition = 3 Then
        UpdateTime
    ElseIf oSW.View.CurrentShowPosition = 5 Then
        Slide7.Shapes("EndShowClickarea").Visible = msoFalse
    ElseIf oSW.View.CurrentShowPosition = 6 Then
        With Slide7.Shapes("Windows Flora Shutdown (build 632-893)").AnimationSettings
            .AdvanceMode = ppAdvanceOnTime
            .PlaySettings.PlayOnEntry = True
            .PlaySettings.PauseAnimation = False
            .PlaySettings.StopAfterSlides = 999
        End With
        ResetWindows
        SavePresentation
        Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "False"
        Slide8.Shapes("BootText").TextFrame.TextRange.Text = "Recovery mode"
        Slide2.Shapes("BootText").TextFrame.TextRange.Text = "Window templates"
        If Slide7.Shapes("Restart").TextFrame.TextRange.Text = "True" Then
            Slide7.Shapes("Restart").TextFrame.TextRange.Text = "False"
            Slide7.SlideShowTransition.AdvanceOnTime = msoTrue
            SavePresentation
            Slide3.Shapes("BootText").TextFrame.TextRange.Text = "Starting OrangePath/OS"
            Slide3.Shapes("Bootwarning").Visible = msoFalse
            Slide3.Shapes("Bootlogo").Visible = msoTrue
            ActivePresentation.SlideShowWindow.View.Next
        ElseIf Slide7.Shapes("Restart").TextFrame.TextRange.Text = "Recovery" Then
            Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "Recovery"
            Slide7.Shapes("Restart").TextFrame.TextRange.Text = "False"
            Slide7.SlideShowTransition.AdvanceOnTime = msoTrue
            SavePresentation
            ActivePresentation.SlideShowWindow.View.Next
        Else
            Slide7.SlideShowTransition.AdvanceOnTime = msoFalse
            Slide7.Shapes("BootText").TextFrame.TextRange.Text = "It's now safe to close the presentation"
            Slide7.Shapes("EndShowClickarea").Visible = msoTrue
        End If
    ElseIf oSW.View.CurrentShowPosition = 11 Then
        If Slide12.Shapes("FirmwareSource").TextFrame.TextRange.Text <> "*" Then
            Slide12.Shapes("StatusText").TextFrame.TextRange.Text = "Updating system..."
            Slide12.Shapes("Notice").TextFrame.TextRange.Text = "Do not close the presentation!"
            ActivePresentation.SlideShowWindow.Activate
            ResetWindows
            Pause (1)
            UpdateTest
            Pause (4)
            Restart
        Else
            Slide12.Shapes("StatusText").TextFrame.TextRange.Text = "Firmware location not specified"
            Slide12.Shapes("Notice").TextFrame.TextRange.Text = "Returning to recovery mode in 5 seconds..."
            Slide8.Shapes("BootText").TextFrame.TextRange.Text = "System update failed. How would you like to proceed?"
            ActivePresentation.SlideShowWindow.Activate
            Pause 5
            ActivePresentation.SlideShowWindow.View.GotoSlide 8
        End If
    ElseIf oSW.View.CurrentShowPosition = 20 Then
        If Slide19.Shapes("MsgDisplayed").TextFrame.TextRange.Text = "False" Then
            Slide19.Shapes("MsgDisplayed").TextFrame.TextRange.Text = "True"
            AppMessage "This version of Codename OrangePath/OS is not finished and therefore may be unstable. Please report any detected issues to the developer as soon as possible.", "Evaluation copy", "Info", False
        End If
    ElseIf oSW.View.CurrentShowPosition = 23 Then
        ' oh dear
        Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "Recovery"
        Slide8.Shapes("BootText").TextFrame.TextRange.Text = "System was shut down because of an error. How would you like to proceed?"
        Pause 2
        SavePresentation
        Pause 2
        ActivePresentation.SlideShowWindow.View.GotoSlide (24)
    End If
End Sub


Sub CleanPopups()
    Dim Sld As slide
    Dim Shp As shape
    IDX = 1
    For Each Sld In ActivePresentation.Slides
        If IDX <> 4 And IDX <> 10 And IDX <> 13 And IDX <> 9 And IDX <> 24 And IDX <> 27 And IDX <> 26 And IDX <> ActivePresentation.SlideShowWindow.View.CurrentShowPosition Then
            For Each Shp In Sld.Shapes
                If InStr(Shp.Name, ":") Then
                    Shp.Delete
                End If
            Next Shp
        End If
        IDX = IDX + 1
    Next Sld
End Sub

Sub Restart()
    Slide2.Shapes("BootText").TextFrame.TextRange.Text = "Restarting..."
    Slide7.Shapes("BootText").TextFrame.TextRange.Text = "Restarting..."
    Slide7.Shapes("Restart").TextFrame.TextRange.Text = "True"
    ActivePresentation.SlideShowWindow.View.GotoSlide (5)
End Sub

Sub RestartRecovery()
    Slide2.Shapes("BootText").TextFrame.TextRange.Text = "Please wait..."
    Slide7.Shapes("BootText").TextFrame.TextRange.Text = "Please wait..."
    Slide7.Shapes("Restart").TextFrame.TextRange.Text = "Recovery"
    ActivePresentation.SlideShowWindow.View.GotoSlide (5)
End Sub

' Set global variable
Sub SetVar(ByVal Key As String, ByVal Value As String)
    If ShapeExists(Slide21, Key) = False Then
        With Slide21.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 0, 0)
            .Name = Key
            .TextFrame.TextRange.Text = Value
            .Visible = msoFalse
        End With
    Else
        If Value <> "" Then
            Slide21.Shapes(Key).TextFrame.TextRange.Text = Value
        Else
            Slide21.Shapes(Key).Delete
        End If
    End If
End Sub

' Unset global variable
Sub UnsetVar(ByVal Key As String)
    If ShapeExists(Slide21, Key) = True Then
        Slide21.Shapes(Key).Delete
    End If
End Sub


Function CheckVars(ByVal str As String)
    outStr = str
    For Each Shp In Slide21.Shapes
        outStr = Replace(outStr, "%" & Shp.Name & "%", Shp.TextFrame.TextRange.Text)
    Next Shp
    CheckVars = outStr
End Function

Sub Highlight(Shp As shape)
    Unlight
    Shp.Fill.ForeColor = Slide8.Shapes("Highlight").Fill.ForeColor
    Shp.TextFrame.TextRange.Font.Color = Slide8.Shapes("Highlight").TextFrame.TextRange.Font.Color
End Sub

Sub Unlight()
    For Each Shp In Slide8.Shapes
        If InStr(1, Shp.Name, "Btn") = 1 Then
            Shp.Fill.ForeColor = Slide8.Shapes("Unlight").Fill.ForeColor
            Shp.TextFrame.TextRange.Font.Color = Slide8.Shapes("Unlight").TextFrame.TextRange.Font.Color
        End If
    Next Shp
End Sub

Sub AddUserOOBE()
    On Error GoTo Crash
    If Slide19.PassField.Text <> Slide19.ConfirmPassField.Text Then
        Slide19.Shapes("MsgDisplayed").TextFrame.TextRange.Text = "True"
        AppMessage "Passwords do not match", "Out of box experience", "Error", False
        Exit Sub
    End If
    If Slide19.UsernameFIeld.Text = "Nobody" Then
        AppMessage "This is a reserved username, which cannot be used", "Out of box experience", "Error", False
        Exit Sub
    End If
    Slide15.Export Environ("TEMP") & "\Userpic.PNG", "PNG"
    SetFileContent "/Users/" & Slide19.UsernameFIeld.Text & "/Password.txt", Slide19.PassField.Text
    SetFileContent "/Users/" & Slide19.UsernameFIeld.Text & "/Theme.txt", "0"
    SetFilePic "/Users/" & Slide19.UsernameFIeld.Text & "/Background.png", Environ("TEMP") & "\Userpic.PNG"
    If Slide19.Shapes("AutologinCheck").Fill.ForeColor.RGB = Slide1.ColorScheme.Colors(ppFill) Then
        SetFileContent "/System/Settings.cnf", Slide19.UsernameFIeld.Text, "Autologin"
    Else
        SetFileContent "/System/Settings.cnf", "Nobody", "Autologin"
    End If
    ' Reset autosave interval
    SetFileContent "/System/Settings.cnf", "5", "AutosaveInterval"
    
    Slide19.Shapes("AutologinCheck").Fill.ForeColor.RGB = RGB(218, 227, 243)
    Slide19.UsernameFIeld.Text = ""
    Slide19.PassField.Text = ""
    Slide19.ConfirmPassField.Text = ""
    AppMessage "The system will now restart.", "Out of box experience", "Info", False
    With Slide19.Shapes("Shape6AppMessage:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).ActionSettings(ppMouseClick)
        .Run = "Restart"
    End With
Done:
    Exit Sub
Crash:
    Slide22.Shapes("Details").TextFrame.TextRange.Text = "Error details: OUT_OF_BOX_EXPERIENCE_FAIL"
    ActivePresentation.SlideShowWindow.View.GotoSlide 22
End Sub

Sub CheckUncheck(Shp As shape)

    If Shp.Fill.ForeColor.RGB = Slide1.ColorScheme.Colors(ppBackground).RGB Then
        Shp.Fill.ForeColor.RGB = Slide1.ColorScheme.Colors(ppFill)
    Else
        Shp.Fill.ForeColor.RGB = Slide1.ColorScheme.Colors(ppBackground)
    End If
End Sub

Sub DeleteAllUsers()
    Users = GetFiles("/Users/")
    UsersList = Split(Users, "/")
    For i = UBound(UsersList) To 0 Step -1
        User = Replace(UsersList(i), vbNewLine, "")
        Slide1.Shapes("Username").TextFrame.TextRange.Text = User
        DeleteDir "/Users/" & User & "/"
    Next i
End Sub

Sub AddUser()
    If Slide17.PassField.Text <> Slide17.ConfirmPassField.Text Then
        AppMessage "Passwords do not match", "Add user", "Error", False
        Exit Sub
    End If
    If Slide17.UsernameFIeld.Text = "Nobody" Then
        AppMessage "This is a reserved username, which cannot be used.", "Add user", "Error", False
        Exit Sub
    End If
    Slide15.Export Environ("TEMP") & "\Userpic.PNG", "PNG"
    SetFileContent "/Users/" & Slide17.UsernameFIeld.Text & "/Password.txt", Slide17.PassField.Text
    SetFileContent "/Users/" & Slide17.UsernameFIeld.Text & "/Theme.txt", "0"
    SetFilePic "/Users/" & Slide17.UsernameFIeld.Text & "/Background.png", Environ("TEMP") & "\Userpic.PNG"
    Slide17.UsernameFIeld.Text = ""
    Slide17.PassField.Text = ""
    Slide17.ConfirmPassField.Text = ""
    AppMessage "User account has been added.", "Add user", "Info", True
End Sub

Sub TestAdd20Users()
    For i = 1 To 20 Step 1
        Slide15.Export Environ("TEMP") & "\Userpic.PNG", "PNG"
        SetFileContent "/Users/Test" & i & "/Password.txt", ""
        SetFileContent "/Users/Test" & i & "/Theme.txt", "0"
        SetFilePic "/Users/Test" & i & "/Background.png", Environ("TEMP") & "\Userpic.PNG"
    Next i
End Sub

Sub TestFixBackgrounds()
    Users = GetFiles("/Users/")
    UsersList = Split(Users, "/")
    For i = UBound(UsersList) To 0 Step -1
        User = Replace(UsersList(i), vbNewLine, "")
        Slide1.Shapes("Username").TextFrame.TextRange.Text = User
        DeleteFile "/Users/" & User & "/Background.png"
        If FileExists("/Users/" & User & "/Background.png") Then
            DeleteFile "/Users/" & User & "/Background.png"
        End If
        If FileExists("/Users/" & User & "/Background.pngBackground.png") Then
            DeleteFile "/Users/" & User & "/Background.pngBackground.png"
        End If
        CopyFile "/Defaults/Images/Background.png", "/Users/" & User & "/"
    Next i
End Sub


Sub UpdatePass()
    CurrentPass = GetFileContent("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Password.txt")
    EnteredPass = Slide16.OldPassField.Text
    If CurrentPass <> EnteredPass Then
        AppMessage "Current password incorrect", "Password wasn't updated", "Error", False
        Exit Sub
    End If
    NewPass = Slide16.NewPassField.Text
    ConfirmPass = Slide16.ConfirmPassField.Text
    If NewPass <> ConfirmPass Then
        AppMessage "Passwords do not match", "Password wasn't updated", "Error", False
        Exit Sub
    End If
    SetFileContent "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Password.txt", NewPass
    Slide16.NewPassField.Text = ""
    Slide16.OldPassField.Text = ""
    Slide16.ConfirmPassField.Text = ""
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
End Sub

Sub HardReset()
    ' Reset boot logos
    Dim PreShutdown As shape
    Dim Splash As shape
    Dim PreSplash As shape
    Set PreShutdown = GetFileRef("/Defaults/BootlogoPreShutdown.3d")
    Set Splash = GetFileRef("/Defaults/BootlogoSplash.3d")
    Set PreSplash = GetFileRef("/Defaults/BootlogoPreSplash.3d")
    Slide2.Shapes("Bootlogo").ThreeD.RotationZ = PreShutdown.ThreeD.RotationZ
    Slide2.Shapes("Bootlogo").ThreeD.RotationY = PreShutdown.ThreeD.RotationY
    Slide2.Shapes("Bootlogo").ThreeD.RotationX = PreShutdown.ThreeD.RotationX
    Slide3.Shapes("Bootlogo").ThreeD.RotationZ = Splash.ThreeD.RotationZ
    Slide3.Shapes("Bootlogo").ThreeD.RotationY = Splash.ThreeD.RotationY
    Slide3.Shapes("Bootlogo").ThreeD.RotationX = Splash.ThreeD.RotationX
    Slide5.Shapes("Bootlogo").ThreeD.RotationZ = PreSplash.ThreeD.RotationZ
    Slide5.Shapes("Bootlogo").ThreeD.RotationY = PreSplash.ThreeD.RotationY
    Slide5.Shapes("Bootlogo").ThreeD.RotationX = PreSplash.ThreeD.RotationX
    Slide7.Shapes("Bootlogo").ThreeD.RotationZ = Splash.ThreeD.RotationZ
    Slide7.Shapes("Bootlogo").ThreeD.RotationY = Splash.ThreeD.RotationY
    Slide7.Shapes("Bootlogo").ThreeD.RotationX = Splash.ThreeD.RotationX
    
    ' Disable debug mode
    Slide2.Shapes("106*20*E*Shape6AppSettings_").TextFrame.TextRange.Text = "Enable"
    Slide1.Shapes("AppID").Visible = msoFalse
    Slide1.Shapes("MoveEvent").Visible = msoFalse
    Slide1.Shapes("ResizeEvent").Visible = msoFalse
    Slide1.Shapes("AppCreatingEvent").Visible = msoFalse
    Slide1.Shapes("AutosaveTime").Visible = msoFalse
    
    ' Enable OOBE message
    Slide19.Shapes("MsgDisplayed").TextFrame.TextRange.Text = "False"
    
    ' Reset theme
    With Slide1.Master.Theme
        .ThemeColorScheme(msoThemeAccent1) = RGB(68, 114, 196)
        .ThemeColorScheme(msoThemeAccent2) = RGB(237, 125, 49)
        .ThemeColorScheme(msoThemeAccent3) = RGB(165, 165, 165)
        .ThemeColorScheme(msoThemeAccent4) = RGB(255, 192, 0)
        .ThemeColorScheme(msoThemeAccent5) = RGB(91, 155, 213)
        .ThemeColorScheme(msoThemeAccent6) = RGB(112, 173, 71)
        .ThemeColorScheme(msoThemeDark1) = RGB(0, 0, 0)
        .ThemeColorScheme(msoThemeDark2) = RGB(68, 84, 106)
        .ThemeColorScheme(msoThemeLight1) = RGB(255, 255, 255)
        .ThemeColorScheme(msoThemeLight2) = RGB(231, 230, 230)
    End With
    
    ' Log in as root
    Slide1.Shapes("Username").TextFrame.TextRange.Text = "Nobody"
    
    ' Clear temporary files
    DeleteDir "/Temp/"
    NewFolder "/Temp"
    
    ' Delete user accounts
    DeleteDir "/Users/"
    
    ' Delete variables
    DeleteDir "/System/"
    
    ' Close all open windows
    ResetWindows
    Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "Factory"
    ' Display message in recovery mode
    ActivePresentation.SlideShowWindow.View.GotoSlide 8
    Slide8.Shapes("BootText").TextFrame.TextRange.Text = "Factory reset success"
    Pause (1)
    Slide8.Shapes("BootText").TextFrame.TextRange.Text = "Recovery mode"
End Sub

Sub UpdateTime()
    'On Error GoTo Crash
    Slide1.Shapes("ResizeEvent").TextFrame.TextRange.Text = "False"
    Slide1.Shapes("MoveEvent").TextFrame.TextRange.Text = "False"
    invervalStr = GetFileContent("/System/Settings.cnf", "AutosaveInterval")
    Dim target As Long
    target = Minute(Time) + CLng(intervalStr)
    If target >= 60 Then
        target = target - 60
    End If
    disableClock = True
    Do
        fullClock = Split(Time, ":")
        hrMin = fullClock(0) & ":" & fullClock(1) & ":" & fullClock(2)
        'hrMin = Time
        Slide1.Shapes("Clock").TextFrame.TextRange.Text = hrMin
        Slide1.Shapes("AutosaveTime").TextFrame.TextRange.Text = CStr(target)
        If CStr(Minute(Time)) = CStr(target) Then
            intervalStr = GetFileContent("/System/Settings.cnf", "AutosaveInterval")
            target = Minute(Time) + CLng(intervalStr)
            If target >= 60 Then
                target = target - 60
            End If
            SavePresentation
        End If
        If ActivePresentation.SlideShowWindow.View.CurrentShowPosition > 4 Or ActivePresentation.SlideShowWindow.View.CurrentShowPosition < 3 Then
            Exit Do
        End If
        DoEvents
    Loop
Done:
    Exit Sub
Crash:
    Slide22.Shapes("Details").TextFrame.TextRange.Text = "Error details: SYSTEM_WATCHDOG_ERROR"
    ActivePresentation.SlideShowWindow.View.GotoSlide 22
End Sub


Sub ResetWindows()
    For i = 0 To 3
        Dim Shp As shape
        Set Shp = Slide1.Shapes("SwitchWorkspace")
        'MsgBox (Shp.TextFrame.TextRange.Text)
        If Shp.TextFrame.TextRange.Text = "Workspace 1" Then
            For Each Shp In Slide1.Shapes
                If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                    Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                End If
            Next Shp
            Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 2"
        ElseIf Shp.TextFrame.TextRange.Text = "Workspace 2" Then
            For Each Shp In Slide1.Shapes
                If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                    Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                End If
            Next Shp
            Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 3"
        ElseIf Shp.TextFrame.TextRange.Text = "Workspace 3" Then
            For Each Shp In Slide1.Shapes
                If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                    Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                End If
            Next Shp
            Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 4"
        Else
            For Each Shp In Slide1.Shapes
                If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                    Shp.Left = Shp.Left + ActivePresentation.SlideShowWindow.Width * 3
                End If
            Next Shp
            Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 1"
        End If
        For Each Shp In Slide1.Shapes
            If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                Shp.Delete
            End If
        Next Shp
    Next i
End Sub

Function FocusWindow(ByVal AppID As String)
    For Each Shp In Slide1.Shapes
        If InStr(Shp.Name, "TaskIcon:") Then
            If InStr(Shp.Name, AppID) Then
                Shp.Fill.Transparency = 0.4
            Else
                Shp.Fill.Transparency = 0.8
            End If
        End If
        Dim HasAx As Boolean
        HasAx = False
        If InStr(Shp.Name, "RegularApp:" & AppID) Then
            For X = 1 To Shp.GroupItems.Count
                With Shp.GroupItems(X)
                    If InStr(.Name, "AXTextBox") Then
                        ApplyTbAttribs Shp.GroupItems(X)
                        HasAx = True
                    End If
                End With
            Next
        End If
        Slide1.AxTextBox.Visible = HasAx
        ActivePresentation.SlideShowWindow.View.GotoSlide (ActivePresentation.SlideShowWindow.View.CurrentShowPosition)
    Next Shp
End Function

Sub MinimizeWindow(Shp As shape)
    AppID = GetAppID(Shp)
    FocusWindow AppID
    MinimizeRestore Slide1.Shapes("TaskIcon:" & AppID)
End Sub

Sub MinimizeRestore(Shp As shape)
    AppID = GetAppID(Shp)
    If Slide1.Shapes("RegularApp:" & AppID).Visible = msoTrue Then
        If Shp.Fill.Transparency = 0.8 Then
            Slide1.Shapes("RegularApp:" & AppID).ZOrder msoBringToFront
            FocusWindow AppID
            UpdateTime
        Else
            Slide1.Shapes("RegularApp:" & AppID).Visible = msoFalse
            Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.8
            Slide1.AxTextBox.Visible = False
            ActivePresentation.SlideShowWindow.View.GotoSlide (4)
            UpdateTime
        End If
    Else
        Slide1.Shapes("RegularApp:" & AppID).Visible = msoTrue
        Slide1.Shapes("RegularApp:" & AppID).ZOrder msoBringToFront
        FocusWindow AppID
        Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.4
        For X = 1 To Shp.GroupItems.Count
            With Slide1.Shapes("RegularApp:" & AppID).GroupItems(X)
                If InStr(.Name, "AXTextBox") Then
                    ApplyTbAttribs Shp.GroupItems(X)
                    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
                End If
            End With
        Next
        UpdateTime
    End If
End Sub

Sub InvertValue(Shp As shape)
    If Shp.TextFrame.TextRange.Text = "True" Then
        Shp.TextFrame.TextRange.Text = "False"
    Else
        Shp.TextFrame.TextRange.Text = "True"
    End If
End Sub

Sub ReorganizeTaskIcons(ByVal IDX As Integer)
    ' Values derived from GetTaskbarLefts macro
    If IDX = 1 Then
        MoveLeft Slide1.Shapes("TaskbarButtonSample").Left + 136
    ElseIf IDX = 2 Then
        MoveLeft Slide1.Shapes("TaskbarButtonSample").Left + 139
    ElseIf IDX = 3 Then
        MoveLeft Slide1.Shapes("TaskbarButtonSample").Left + 408
    ElseIf IDX = 4 Then
        MoveLeft Slide1.Shapes("TaskbarButtonSample").Left + 544
    ElseIf IDX = 5 Then
        MoveLeft Slide1.Shapes("TaskbarButtonSample").Left + 680
    End If
End Sub

Sub MoveLeft(Left As Integer)
    For Each Shp In Slide1.Shapes
        If InStr(Shp.Name, "TaskIcon:") And Shp.Left > Left And Shp.Left < ActivePresentation.SlideShowWindow.Width Then
            Shp.Left = Shp.Left - Shp.Width
        End If
    Next Shp
End Sub

Function CheckShape(ByVal Left As Integer)
    For Each Shp In Slide1.Shapes
        If InStr(Shp.Name, "TaskIcon:") Then
            If CInt(Shp.Left) >= Left And CInt(Shp.Left) < Left + Shp.Width Then
                CheckShape = True
                Exit Function
            End If
        End If
    Next Shp
    CheckShape = False
End Function

Function SaveSysConfig(ByVal Key As String, ByVal Value As String)
    SetFileContent "/System/Settings.cnf", Value, Key
End Function

Function GetSysConfig(ByVal Name As String) As String
    GetSysConfig = GetFileContent("/System/Settings.cnf", Name)
End Function

Sub SavePresentation()
    On Error GoTo Crash
    
    If Slide7.Shapes("BootText").TextFrame.TextRange.Text = "It's now safe to close the presentation" Then Exit Sub
    With Application.ActivePresentation
        ' save only if the save path is known and there are unsaved changes
        If Not .Saved And .Path <> "" Then .Save
    End With
    
Done:
    Exit Sub
Crash:
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition <> 22 Then
        Slide22.Shapes("Details").TextFrame.TextRange.Text = "Error details: SAVE_ERROR"
        ActivePresentation.SlideShowWindow.View.GotoSlide 22
    End If
End Sub

' Returns the App ID based on the clicked/hovered shape name
Function GetAppID(ByVal Shp As shape) As String
    SplitZ = Split(Shp.Name, ":")
    AppID = SplitZ(1)
    GetAppID = AppID
End Function

' Regroups ungrouped windows
Sub Regroup(ByVal AppID As String, ByVal Sld As slide)
    Dim Shapes As String
    Shapes = ""
    For Each Shp2 In Sld.Shapes()
        If InStr(Shp2.Name, ":" & AppID) Then
            If InStr(Shp2.Name, "TaskIcon") = False Then
            'If InStr(Shp2.Name, "AXTextBox") Then ApplyTbAttribs Shp2
                Shapes = Shapes & Shp2.Name & ","
            End If
        End If
    Next Shp2
    
    SplitShapes = Split(Shapes, ",")
    UJ = CInt(UBound(SplitShapes))
    Dim ShapesX() As String
    
    ReDim ShapesX(UJ)
    For i = 0 To CInt(UBound(SplitShapes) - 1)
        CShape = SplitShapes(i)
        If Not IsInArray(CStr(CShape), ShapesX) Then
            ShapesX(i) = SplitShapes(i)
        End If
    Next
    
    With Sld.Shapes.Range(ShapesX).Group
        .Name = "RegularApp:" & AppID
    End With
End Sub

' Regroups any shape with _ at the end
Sub RegroupShapes(ByVal ShapeName As String, ByVal Sld As slide)
    Dim Shapes As String
    Shapes = ""
    For Each Shp2 In Sld.Shapes()
        If InStr(Shp2.Name, "_") Then
            'If InStr(Shp2.Name, "AXTextBox") Then ApplyTbAttribs Shp2
            Shapes = Shapes & Shp2.Name & ","
        End If
    Next Shp2
    
    SplitShapes = Split(Shapes, ",")
    UJ = CInt(UBound(SplitShapes))
    Dim ShapesX() As String
    
    ReDim ShapesX(UJ)
    For i = 0 To CInt(UBound(SplitShapes) - 1)
        CShape = SplitShapes(i)
        If Not IsInArray(CStr(CShape), ShapesX) Then
            ShapesX(i) = SplitShapes(i)
        End If
    Next
    With Sld.Shapes.Range(ShapesX).Group
        .Name = ShapeName
    End With
End Sub

' Pastes a shape to a window
Sub PasteToGroup(ByVal Ref As shape, _
                ByVal Shp As shape, _
                ByVal Name As String, _
                ByVal OffsetX As Integer, _
                ByVal OffsetY As Integer, _
                ByVal Sld As slide, _
                Optional ByVal Macro As String = "")
    ' Declarations
    Dim AppID As String
    Dim Left As Integer
    Dim Top As Integer
    ' Get AppID from reference shape
    AppID = GetAppID(Ref)
    ' Set position for the shape based on offsets specified
    Left = OffsetX
    Top = OffsetY
    ' Ungroup the current app shape
    Sld.Shapes("RegularApp:" & AppID).Ungroup
    ' Copy source shape
    Shp.Copy
    ' Paste to the specified position
    With Sld.Shapes.Paste
        .Left = Left
        .Top = Top
        .Name = Name
        If Macro <> "" Then
                On Error GoTo IsGroup
                .ActionSettings(ppMouseClick).Run = Macro
                GoTo ContinueP
IsGroup:
                Dim Shp2 As shape
                For Each Shp2 In .GroupItems
                    Shp2.ActionSettings(ppMouseClick).Run = Macro
                Next Shp2
ContinueP:
        End If
    End With
    
    ' Regroup shapes
    Regroup AppID, Sld
End Sub

' Erase a shape from window
Sub EraseFromGroup(ByVal Ref As shape, ByVal ShpName As String, ByVal Sld As slide)
    ' Declarations
    Dim AppID As String
    ' Get AppID from reference shape
    AppID = GetAppID(Ref)
    ' Ungroup the current app shape
    Sld.Shapes("RegularApp:" & AppID).Ungroup
    ' Delete shape specified
    Sld.Shapes(ShpName).Delete
    ' Regroup shapes
    Regroup AppID, Sld
End Sub

Sub SkinShapeLo(ByVal Sld As slide, ByVal Parent As String, ByVal Replacable As String, ByVal ReplacedBy As shape)
    Sld.Shapes(Parent).Ungroup
    Sld.Shapes(Replacable).Delete
    ReplacedBy.Copy
    With Sld.Shapes.Paste
        .Name = Replacable
        .ZOrder msoSendToBack
    End With
    RegroupShapes Parent, Sld
End Sub

Sub SkinShape(ByVal Ref As shape, ByVal control As String, ByVal Thm As shape, ByVal Sld As slide)
    Replacable = control & Ref.Name & "_"
    On Error Resume Next
    X = Sld.Shapes(Replacable).Left
    Y = Sld.Shapes(Replacable).Top
    W = Sld.Shapes(Replacable).Width
    H = Sld.Shapes(Replacable).Height
    T = Sld.Shapes(Replacable).TextFrame.TextRange.Text
    M = Sld.Shapes(Replacable).ActionSettings(ppMouseClick).Run
    MO = Sld.Shapes(Replacable).ActionSettings(ppMouseOver).Run
    SkinShapeLo Sld, Ref.Name, Replacable, Thm.GroupItems(control)
    With Sld.Shapes(Replacable)
        If control <> "WindowFrame" And control <> "WindowTitle" Then
            .Width = Thm.GroupItems(control).Width
            .Height = Thm.GroupItems(control).Height
        Else
            .Width = W
            .Height = H
        End If
        .Left = X
        .Top = Y
        If control = "WindowTitle" Then
            .TextFrame.TextRange.Text = T
        End If
        .ActionSettings(ppMouseClick).Run = M
        .ActionSettings(ppMouseOver).Run = MO
    End With
SkipThisShape:
    Exit Sub
End Sub


Sub ApplyThemeTest()
    Filename = "/Defaults/Themes/Flat.thm"
    ApplyTheme Filename
End Sub

Sub ApplyTheme(ByVal Filename As String)
    'Slide2.Shapes("Icon4Part1AppMenu_").ZOrder msoSendToBack
    Dim Thm As shape
    Dim W As Integer
    Dim H As Integer
    Dim X As Integer
    Dim Y As Integer
    Set Thm = GetFileRef(Filename)
    Dim ShpArr() As shape
    
    
    For i = Slide2.Shapes.Count To 1 Step -1
        If Slide2.Shapes(i).Type = msoGroup Then
            If InStr(1, Slide2.Shapes(i).Name, "App") = 1 Then
                If Slide2.Shapes(i).Name <> "AppMenu" Then
                    SkinShape Slide2.Shapes(i), "Close", Thm, Slide2
                    SkinShape Slide2.Shapes(i), "Minimize", Thm, Slide2
                    SkinShape Slide2.Shapes(i), "WindowTitle", Thm, Slide2
                    SkinShape Slide2.Shapes(i), "WindowFrame", Thm, Slide2
                    Slide2.Shapes(i).Visible = msoFalse
                End If
            End If
        End If
    Next i
End Sub