Public Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As LongPtr
Public disableClock As Boolean


' OrangePath/OS system module

Sub Pause(Length As LongPtr)
    Dim NowTime As LongPtr
    Dim endtime As LongPtr

    endtime = GetTickCount + (Length * 1000)
    Do
        NowTime = GetTickCount
        DoEvents
    Loop Until NowTime >= endtime
End Sub

'
' Takes care of animation triggers that appear "empty" in the animation pane (what's really going on is that some shapes that have animations there are hidden by ShapeObject.Visible = msoFalse),
' which can in some cases corrupt and/or crash the presentation (seems to commonly occur with sound files for some reason). Normally, anything that has a colon in the shape name should get
' cleared if the CleanPopups subroutine is run, but it is designed with certain exceptions in mind.
'
' DO NOT save the presentation during an active slideshow unless you have run this subroutine first! If you use SavePresentation routine (as you really should), it will automatically run this
' subroutine for you. We will not fix presentation files that have been corrupted due to having run badly coded custom applications.
'
Sub DeleteOrphanedTriggers()
    Dim Slide As Slide
    Dim timeLine As timeLine
    Dim mainSeq As Sequences
    Dim animEffect As Effect
    Dim triggerShape As Shape
    Dim i As Integer
    
    Set Slide = Slide1
    Set timeLine = Slide.timeLine
    If timeLine.InteractiveSequences.Count > 0 Then
        Set mainSeq = timeLine.InteractiveSequences
        Dim seq As Sequence
        For i = mainSeq.Count To 1 Step -1
            Set seq = mainSeq.Item(i)
            If Slide.Shapes(seq.Item(1).DisplayName).Visible = msoFalse Then
                Slide.Shapes(seq.Item(1).DisplayName).Delete
            End If
        Next i
    End If
End Sub

Sub IntResetDragSize()
    If Slide1.Shapes("MoveEvent").Visible = msoTrue Then
        Dim Backup As Long
        Dim Backup2 As Long
        Backup = Slide1.Shapes("MoveEvent").Fill.ForeColor.RGB
        Backup2 = Slide1.Shapes("ResizeEvent").Fill.ForeColor.RGB
        Slide1.Shapes("MoveEvent").Fill.ForeColor.RGB = Backup2
        Slide1.Shapes("ResizeEvent").Fill.ForeColor.RGB = Backup
        ActivePresentation.SlideShowWindow.Activate
    End If
    Slide1.Shapes("ResizeEvent").TextFrame.TextRange.Text = "False"
    Slide1.Shapes("MoveEvent").TextFrame.TextRange.Text = "False"
End Sub

Sub LogData(ByVal Data As String)
    If Slide1.Shapes("Username").Visible = msoTrue Then
        Debug.Print "[Trace] " & Data
    End If
End Sub

Sub ExitAppMenu()
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    UpdateTime
End Sub

Function GetShape(ByVal Sld As Slide, ByVal AppID As String, ByVal AppName As String, ByVal ShapeLabel As String) As Shape
    GetShape = Sld.Shapes(ShapeLabel & "App" & AppName & ":" & AppID)
End Function

Function GetShapeText(ByVal Sld As Slide, ByVal AppID As String, ByVal AppName As String, ByVal ShapeLabel As String) As String
    GetShapeText = Sld.Shapes(ShapeLabel & "App" & AppName & ":" & AppID).TextFrame.TextRange.Text
End Function

Sub WaitCursor(RootShape As Shape, Optional ByVal WaitText As String = "Please wait...")
    If GetSysConfig("Loaders") = "False" Then Exit Sub
    If RootShape Is Nothing Then
        Slide1.Shapes("WaitPlease").Left = 0
        Slide1.Shapes("WaitPlease").Top = 0
    Else
        Slide1.Shapes("WaitPlease").Left = RootShape.Left
        Slide1.Shapes("WaitPlease").Top = RootShape.Top
    End If
    
    Slide1.Shapes("WaitLabel").TextFrame.TextRange.Text = vbNewLine & vbNewLine & WaitText
    Slide1.Shapes("WaitPlease").Visible = msoTrue
    Slide1.Shapes("WaitPlease").ZOrder msoBringToFront
    ActivePresentation.SlideShowWindow.View.GotoSlide ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    ActivePresentation.SlideShowWindow.Activate
    IntHideCursor
End Sub

' Internal function, DO NOT USE OUTSIDE KERNEL SUBROUTINES!
Sub IntHideCursor()
    Slide1.Shapes("WaitPlease").Visible = msoFalse
    Slide1.Shapes("WaitPlease").Left = 0
    Slide1.Shapes("WaitPlease").Top = 0
End Sub

' For backwards compatibility
Sub HideCursor()
    ActivePresentation.SlideShowWindow.Activate
End Sub

'
' You must use this function when initializing a windowed application. A fullscreen exclusive application (e.g. game) should only be coded to use the special slide
' (Slide27 in position 28), however please note that doing this will disable any multitasking capability. This can also happen during runtime. Note that before
' leaving Slide27, ALL shapes must be deleted before doing so. If you are using colon in the name of all shapes, you can use CleanPopups before leaving the slide
' to accomplish this.
'
Sub CreateNewWindow()
    ' Variables
    'ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    
    If GetSysConfig("Loaders") <> "False" Then
        Slide1.Shapes("WaitPlease").Left = 0
        Slide1.Shapes("WaitPlease").Top = 0
        Slide1.Shapes("WaitLabel").TextFrame.TextRange.Text = vbNewLine & vbNewLine & "Launching " & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & "..."
        Slide1.Shapes("WaitPlease").Visible = msoTrue
        Slide1.Shapes("WaitPlease").ZOrder msoBringToFront
    End If
    ActivePresentation.SlideShowWindow.Activate
    Slide1.Shapes("AppID").TextFrame.TextRange.Text = CStr(CInt(Slide1.Shapes("AppID").TextFrame.TextRange.Text) + 1)
    AppID = Slide1.Shapes("AppID").TextFrame.TextRange.Text
    Dim Shp As Shape
    ' Skip creating taskbar icon or any other nonsense if we're opening modal windows
    Dim IsModal As Boolean
    IsModal = True
    If Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Menu" Then GoTo ModalSkip
    If Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Message" Then GoTo ModalSkip
    If Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "InputBox" Then GoTo ModalSkip
    If InStr(1, Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text, "Modal") = 1 Then GoTo ModalSkip
    IsModal = False
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
    CurrentSlide = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    With ActivePresentation.Slides(CurrentSlide).Shapes.Paste
    .Name = "RegularApp:" & AppID
    .Visible = msoTrue
    End With
    Dim CurrentSld As Slide
    Set CurrentSld = ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition)
    CurrentSld.Shapes("RegularApp:" & AppID).Ungroup
    For Each Shp In CurrentSld.Shapes
        If Shp.Name = "AssocIconApp" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & "_" Then
            With Shp
                If Not FileStreamsExist("/System/Icons/") Then
                    NewFolder "/System/Icons"
                End If
                If Not FileStreamsExist("/System/Icons/" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & ".emf") Then
                    Shp.Copy
                    With Slide9.Shapes.Paste
                        .Name = "/System/Icons/" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & ".emf"
                        .Visible = msoFalse
                        .LockAspectRatio = msoTrue
                    End With
                End If
                Shp.Delete
            End With
        End If
    Next Shp
    If Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text <> "Menu" Then
        If FileStreamsExist("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm") Then
            GetFileRef("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm").Copy
        Else
            GetFileRef("/Defaults/Themes/Default.thm").Copy
        End If
        With CurrentSld.Shapes.Paste
            .GroupItems("Handle").Delete
            .Ungroup
        End With
        Dim IsTrafficLight As Boolean
        IsTrafficLight = False
        Dim RootShp As Shape
        Set RootShp = CurrentSld.Shapes("WindowApp" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & "_")
        If InStr(1, CurrentSld.Shapes("Window").TextFrame.TextRange.Text, "[Trafficlight]") Then
            IsTrafficLight = True
        End If
        If Not IsTrafficLight Then
            CurrentSld.Shapes("Close").Height = 22.7
            CurrentSld.Shapes("Close").Width = 42.5
            CurrentSld.Shapes("Minimize").Height = 22.7
            CurrentSld.Shapes("Minimize").Width = 42.5
        End If
        CurrentSld.Shapes("Window").Delete
        CurrentSld.Shapes("WindowFrame").ActionSettings(ppMouseOver).Run = "IntResetDragSize"
        CurrentSld.Shapes("Minimize").ActionSettings(ppMouseOver).Run = "IntResetDragSize"
        CurrentSld.Shapes("Close").ActionSettings(ppMouseOver).Run = "IntResetDragSize"
        With CurrentSld.Shapes("WindowFrame")
            .Left = RootShp.Left - 6.837
            .Top = RootShp.Top - CurrentSld.Shapes("Close").Height
            .Width = RootShp.Width + (6.837 * 2)
            .Height = RootShp.Height + CurrentSld.Shapes("Close").Height + 6.837
            .Name = "WindowFrameApp" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & "_"
        End With
        Dim FrameOfRef As Shape
        Set FrameOfRef = CurrentSld.Shapes("WindowFrameApp" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & "_")
        With CurrentSld.Shapes("WindowTitle")
            .Left = RootShp.Left
            If IsTrafficLight Then .Height = 22.4
            If Not IsTrafficLight Then .Top = FrameOfRef.Top
            If IsTrafficLight Then .Top = RootShp.Top - .Height
            .TextFrame.TextRange.Text = Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text
            .Name = "WindowTitleApp" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & "_"
            .ZOrder msoSendToBack
            If IsModal Then
                .ActionSettings(ppMouseOver).Action = ppActionNone
            End If
        End With
        Dim AppTitleShp As Shape
        Set AppTitleShp = ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition).Shapes("WindowTitleApp" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & "_")
        With CurrentSld.Shapes("Close")
            If Not IsTrafficLight Then
                .Left = RootShp.Left + RootShp.Width - .Width
                .Top = FrameOfRef.Top
            Else
                .Left = RootShp.Left + 5
                .Top = AppTitleShp.Top + (AppTitleShp.Height / 2) - (.Height / 2)
            End If
            .Name = "CloseApp" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & "_"
            .ZOrder msoBringToFront
        End With
        With CurrentSld.Shapes("Minimize")
            If Not IsTrafficLight Then
                .Left = RootShp.Left + RootShp.Width - (.Width * 2)
                .Top = FrameOfRef.Top
            Else
                .Left = RootShp.Left + 10 + .Width
                .Top = AppTitleShp.Top + (AppTitleShp.Height / 2) - (.Height / 2)
            End If
            .Name = "MinimizeApp" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & "_"
            .ZOrder msoBringToFront
            If IsModal Then
                .Delete
            End If
        End With
        FrameOfRef.ZOrder msoSendToBack
        With CurrentSld.Shapes("WindowTitleApp" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & "_")
            If Not IsTrafficLight Then
                If Not IsModal Then
                    .Width = CurrentSld.Shapes("MinimizeApp" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & "_").Left - .Left
                Else
                    .Width = CurrentSld.Shapes("CloseApp" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & "_").Left - .Left
                End If
            Else
                .Width = RootShp.Width
            End If
        End With
        If IsModal Then
            With CurrentSld.Shapes.AddShape(msoShapeRectangle, 0, 0, ActivePresentation.PageSetup.SlideWidth, ActivePresentation.PageSetup.SlideHeight)
                .Name = "ModalBackdropApp" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & "_"
                .Fill.ForeColor.RGB = RGB(0, 0, 0)
                .Fill.Transparency = 0.7
                .Line.Weight = 0
                .Line.Transparency = 1
                .ZOrder msoSendToBack
            End With
        End If
    End If
    For Each Shp In ActivePresentation.Slides(CurrentSlide).Shapes()
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
    
    With ActivePresentation.Slides(CurrentSlide).Shapes.Range(ShapesX).Group
        .Name = "RegularApp:" & AppID
    End With
    On Error Resume Next
    If Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text <> "Menu" Then ActivePresentation.Slides(CurrentSlide).Shapes("TaskIcon:" & AppID).Visible = msoTrue

    ' Special case for Taskmgr
    RefreshTaskmgrs
    If Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Taskmgr" Then
        TaskmgrRefresh Slide1.Shapes("RegularApp:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text)
    End If
    Dim eff As Effect
    For i = ActivePresentation.Slides(CurrentSlide).timeLine.MainSequence.Count To 1 Step -1
        Set eff = Slide1.timeLine.MainSequence(i)
        eff.Delete
    Next i
    ActivePresentation.SlideShowWindow.View.GotoSlide (CurrentSlide)
    Slide1.Shapes("WaitPlease").Visible = msoFalse
    If CheckVars("%Animations%") = "True" Then
        Dim oeff As Effect
        Set oeff = ActivePresentation.Slides(CurrentSlide).timeLine.MainSequence.AddEffect(Shape:=ActivePresentation.Slides(CurrentSlide).Shapes("RegularApp:" & AppID), effectId:=msoAnimEffectStrips, trigger:=msoAnimTriggerAfterPrevious)
        oeff.Exit = msoFalse
        oeff.EffectParameters.Direction = msoAnimDirectionBottomRight
        oeff.Timing.Duration = 0.5
        ActivePresentation.Slides(CurrentSlide).Shapes("RegularApp:" & AppID).Visible = msoTrue
        ActivePresentation.SlideShowWindow.Activate
        Pause 2
        For i = ActivePresentation.Slides(CurrentSlide).timeLine.MainSequence.Count To 1 Step -1
            Set eff = Slide1.timeLine.MainSequence(i)
            eff.Delete
        Next i
        ActivePresentation.Slides(CurrentSlide).Shapes("RegularApp:" & AppID).Visible = msoFalse
        ActivePresentation.SlideShowWindow.Activate
        ActivePresentation.Slides(CurrentSlide).Shapes("RegularApp:" & AppID).Visible = msoTrue
        ActivePresentation.SlideShowWindow.Activate
        ActivePresentation.SlideShowWindow.View.GotoSlide (CurrentSlide)
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

Sub CloseWindow(Shp As Shape)
    AppID = 0
    ' Note for app developers
    ' The following condition is only designed to be used by Taskmgr. For third party applications, we recommend calling CloseWindow with a shape directly
    If InStr(Shp.TextFrame.TextRange.Text, "PID") Then
        SplitA = Split(Shp.TextFrame.TextRange.Text, ":")
        StringA = SplitA(UBound(SplitA))
        SplitB = Split(StringA, " ")
        StringB = SplitB(1)
        SplitC = Split(StringB, ")")
        AppID = SplitC(0)
        ' some neat recursion :)
        CloseWindow Slide1.Shapes("RegularApp:" & AppID).GroupItems(1)
    Else
        CheckActiveX Shp
        SplitZ = Split(Shp.ParentGroup.Name, ":")
        AppID = SplitZ(1)
        LogData "Closing window with ID " & AppID
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
    If Not ShapeExists(Slide1, "RegularApp:-1") Then
        GetFileRef("/Defaults/Desktop.grp").Copy
        With Slide1.Shapes.Paste
            .Name = "RegularApp:-1"
            .Left = 0
            .Top = 0
            .Visible = msoTrue
            .ZOrder msoSendToBack
            Slide1.Shapes("BackgroundImg").ZOrder msoSendToBack
            Slide1.Shapes("AnimationRect").ZOrder msoSendToBack
            .GroupItems("PathAppFiles:-1").TextFrame.TextRange.Text = "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Desktop/"
        End With
        Reload "-1"
    End If
    UpdateTime
End Sub

' Internal function, DO NOT USE! Will break things if used incorrectly!
Sub Animate(ByVal AppID As String, ByVal Sld As Slide)
    If CheckVars("%Animations%") = "True" Then
        LogData "Running animation for window with ID " & AppID
        Dim eff As Effect
        For i = Slide1.timeLine.MainSequence.Count To 1 Step -1
            Set eff = Slide1.timeLine.MainSequence(i)
            eff.Delete
        Next i
        Dim oeff As Effect
        Set oeff = Sld.timeLine.MainSequence.AddEffect(Shape:=Sld.Shapes("RegularApp:" & AppID), effectId:=msoAnimEffectStrips, trigger:=msoAnimTriggerAfterPrevious)
        oeff.Exit = msoTrue
        oeff.EffectParameters.Direction = msoAnimDirectionTopLeft
        oeff.Timing.Duration = 0.5
        ActivePresentation.SlideShowWindow.Activate
    End If
End Sub

Sub CloseTest()
    AppID = 18
    Dim Shp As Shape
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
            For Each s In Slide1.Shapes
                sName = Split(s.Name, ":")
                sId = sName(UBound(sName))
                If sId = AppID Then
                    s.Delete
                End If
            Next s
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

Function AAX() As Boolean
    If GetSysConfig("NoActiveX") <> "True" Then
        AAX = True
    Else
        AAX = False
    End If
End Function

Sub CheckActiveX(ByVal Shp As Shape)
    For Each SubShp In Shp.ParentGroup.GroupItems
        If InStr(SubShp.Name, "AXTextBox") And AAX Then
            If AAX Then Slide1.AxTextBox.Visible = False
            If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then Slide13.AxTextBox.Visible = False
        End If
    Next SubShp
End Sub

Sub CheckActiveXShow(Shp As Shape)
    Dim SubShp As Shape
    For Each SubShp In Shp.ParentGroup.GroupItems
        If InStr(SubShp.Name, "AXTextBox") Then ApplyTbAttribs SubShp
    Next SubShp
End Sub

Sub MovableWindow(Shp As Shape)
    ' If moving, exit sub
    If Slide1.Shapes("MoveEvent").TextFrame.TextRange.Text = "True" Then
        Exit Sub
    End If
    If AAX Then
        If Slide1.AxComboBox.Visible Then
            Slide1.AxComboBox.Visible = False
        End If
    End If
    Dim LoopState As Boolean
    SplitZ = Split(Shp.Name, ":")
    AppID = SplitZ(1)
    LogData "Starting move event for window with ID " & AppID
    For i = Slide1.timeLine.MainSequence.Count To 1 Step -1
        Set oeff = Slide1.timeLine.MainSequence(i)
        oeff.Delete
    Next i
    
    Do
        CurrentSlide = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
        GetCursorPositionX1 = GetCursorX
        GetCursorPositionY1 = GetCursorY
        GetAsyncKeyState1 = GetAsyncKeyState(1)
        ' TitleBar clicked, window moving
        If LoopState And GetAsyncKeyState1 Then
            ActivePresentation.Slides(CurrentSlide).Shapes("RegularApp:" & AppID).Top = GetCursorPositionY1 - dy
            ActivePresentation.Slides(CurrentSlide).Shapes("RegularApp:" & AppID).Left = GetCursorPositionX1 - dx
            ActivePresentation.Slides(CurrentSlide).Shapes("RegularApp:" & AppID).ZOrder msoBringToFront
            If AAX Then
                Slide1.AxTextBox.Visible = False
            End If
            If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then Slide13.AxTextBox.Visible = False
            
            ActivePresentation.SlideShowWindow.View.GotoSlide (CurrentSlide)
        ' TitleBar clicked, window hasn't moved
        ElseIf LoopState = False And GetAsyncKeyState1 Then
            dx = GetCursorPositionX1 - ActivePresentation.Slides(CurrentSlide).Shapes("RegularApp:" & AppID).Left
            dy = GetCursorPositionY1 - ActivePresentation.Slides(CurrentSlide).Shapes("RegularApp:" & AppID).Top - 5
            Slide1.Shapes("MoveEvent").TextFrame.TextRange.Text = "True"
        ' TitleBar not clicked, window was moving
        ElseIf LoopState = True And GetAsyncKeyState1 = False Then
            Slide1.Shapes("MoveEvent").TextFrame.TextRange.Text = "False"
            If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 4 Then FocusWindow AppID
            For Each Shp In ActivePresentation.Slides(CurrentSlide).Shapes("RegularApp:" & AppID).GroupItems
                If InStr(Shp.Name, "AXTextBox") Then ApplyTbAttribs Shp
            Next Shp
            ActivePresentation.SlideShowWindow.View.GotoSlide (CurrentSlide)
            Exit Sub
        End If
        DoEvents
        LoopState = GetAsyncKeyState1
    Loop
End Sub

Sub SetTextBox()
    If Not AAX Then Exit Sub
    With Slide1.AxTextBox
        .Visible = True
        .Left = Slide1.Shapes("AxTextBox1AppNotes:21").Left
        .Top = Slide1.Shapes("AxTextBox1AppNotes:21").Top
        .Width = Slide1.Shapes("AxTextBox1AppNotes:21").Width
        .Height = Slide1.Shapes("AxTextBox1AppNotes:21").Height
        .Text = Slide1.Shapes("AxTextBox1AppNotes:21").TextFrame.TextRange.Text
    End With
End Sub

Sub SetTextBoxVal(Shp As Shape)
    If Not AAX Then Exit Sub
    CurrentSlide = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    If CurrentSlide <> 13 Then
        Shp.TextFrame.TextRange.Text = Slide1.AxTextBox.Text
    Else
        Shp.TextFrame.TextRange.Text = Slide13.AxTextBox.Text
    End If
    
End Sub

Function ShapeExists(ByVal oSl As Slide, ByVal ShapeName As String) As Boolean
   Dim oSh As Shape
   For Each oSh In oSl.Shapes
     If oSh.Name = ShapeName Then
        ShapeExists = True
        Exit Function
     End If
   Next ' Shape
   ' No shape here, so though it's not strictly necessary:
   ShapeExists = False
End Function

Function GroupItemExists(ByVal oSl As Shape, ByVal ShapeName As String) As Boolean
   Dim oSh As Shape
   For Each oSh In oSl.GroupItems
     If oSh.Name = ShapeName Then
        GroupItemExists = True
        Exit Function
     End If
   Next ' Shape
   ' No shape here, so though it's not strictly necessary:
   GroupItemExists = False
End Function

Sub ResizingWindow(Shp As Shape)
    On Error Resume Next
    ' If reszing, exit sub
    If Slide1.Shapes("ResizeEvent").TextFrame.TextRange.Text = "True" Then
       Exit Sub
    End If
    LogData "Starting resize event for window with ID " & AppID
    
    Dim LoopState As Boolean
    SplitZ = Split(Shp.Name, ":")
    AppID = SplitZ(1)
    SplitN = Split(Shp.Name, "App")
    SplitO = Split(SplitN(1), ":")
    AppName = SplitO(0)
    Dim O As Shape
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
            If AAX Then
                Slide1.AxTextBox.Visible = False
            End If
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
            Dim RootShp As Shape
            Set RootShp = Slide1.Shapes("WindowApp" & AppName & ":" & AppID)
            Dim IsTrafficLight As Boolean
            IsTrafficLight = False
            If FileStreamsExist("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm") Then
                If InStr(1, GetFileRef("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm").GroupItems("Window").TextFrame.TextRange.Text, "[Trafficlight]") Then
                    IsTrafficLight = True
                End If
            End If
            ActivePresentation.SlideShowWindow.View.GotoSlide (4)
            With Slide1.Shapes("RegularApp:" & AppID)
                Dim FrameOfRef As Shape
                Set FrameOfRef = .GroupItems("WindowFrameApp" & AppName & ":" & AppID)
                With FrameOfRef
                    .Left = RootShp.Left - 6.837
                    .Top = RootShp.Top - 22.7
                    .Width = RootShp.Width + (6.837 * 2)
                    .Height = RootShp.Height + 22.7 + 6.837
                End With
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
                            If Not IsTrafficLight Then
                                Slide1.Shapes(.Name).Width = 42.5
                                Slide1.Shapes(.Name).Height = 22.7
                                Slide1.Shapes(.Name).Left = RootShp.Left + RootShp.Width - .Width
                                Slide1.Shapes(.Name).Top = Slide1.Shapes("WindowFrameApp" & AppName & ":" & AppID).Top
                            Else
                                Slide1.Shapes(.Name).Width = 8.5
                                Slide1.Shapes(.Name).Height = Slide1.Shapes(.Name).Width
                                Slide1.Shapes(.Name).Left = RootShp.Left + 5
                                Slide1.Shapes(.Name).Top = RootShp.Top - (22.7 / 2) - (8.5 / 2)
                            End If
                        ElseIf InStr(.Name, "Minimize") Then
                            If Not IsTrafficLight Then
                                Slide1.Shapes(.Name).Width = 42.5
                                Slide1.Shapes(.Name).Height = 22.7
                                Slide1.Shapes(.Name).Left = RootShp.Left + RootShp.Width - (.Width * 2)
                                Slide1.Shapes(.Name).Top = Slide1.Shapes("WindowFrameApp" & AppName & ":" & AppID).Top
                            Else
                                Slide1.Shapes(.Name).Width = 8.5
                                Slide1.Shapes(.Name).Height = Slide1.Shapes(.Name).Width
                                Slide1.Shapes(.Name).Left = RootShp.Left + 10 + Slide1.Shapes(.Name).Width
                                Slide1.Shapes(.Name).Top = RootShp.Top - (22.7 / 2) - (8.5 / 2)
                            End If
                        ElseIf InStr(.Name, "WindowFrameApp") Then
                            Slide1.Shapes(.Name).Left = RootShp.Left - 6.837
                            Slide1.Shapes(.Name).Top = RootShp.Top - 22.7
                            Slide1.Shapes(.Name).Width = RootShp.Width + (6.837 * 2)
                            Slide1.Shapes(.Name).Height = RootShp.Height + 22.7 + 6.837
                        ElseIf InStr(.Name, "WindowTitle") Then
                            Slide1.Shapes(.Name).Left = RootShp.Left
                            Slide1.Shapes(.Name).Top = Slide1.Shapes("WindowFrameApp" & AppName & ":" & AppID).Top
                            Slide1.Shapes(.Name).Width = RootShp.Width
                            Slide1.Shapes(.Name).Height = 22.7
                            If Not IsTrafficLight Then
                                Slide1.Shapes(.Name).Width = RootShp.Width - (2 * 42.5)
                            Else
                                Slide1.Shapes(.Name).Top = RootShp.Top - Slide1.Shapes(.Name).Height
                            End If
                        End If
                    End With
                Next
                TryRunMacro Slide1.Shapes("TaskIcon:" & AppID).TextFrame.TextRange.Text, "SizeChanged", CInt(AppID)
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

' Converts Microsoft Office TriState value to a standard boolean
Function MsoTristateToBool(ByVal State As MsoTriState, Optional ByVal Strict As Boolean = False) As Boolean
    If State = msoTrue Then
        MsoTristateToBool = True
    ElseIf State = msoCTrue Then
        MsoTristateToBool = Not Strict
    ElseIf State = msoFalse Then
        MsoTristateToBool = False
    ElseIf State = msoTriStateMixed Then
        MsoTristateToBool = Not Strict
    ElseIf State = msoTriStateToggle Then
        MsoTristateToBool = Not Strict
    End If
End Function

' Internal function, do not call this subroutine outside the kernel!
Sub ApplyTbAttribs(Shp As Shape)
    If Not AAX Then Exit Sub
    AppID = GetAppID(Shp)
    LogData "Displaying ActiveX control for window with ID " & AppID
    Dim isBold As Boolean
    Dim isItalic As Boolean
    Dim isUnderline As Boolean
    Dim isStriken As Boolean
    isBold = MsoTristateToBool(Shp.TextFrame.TextRange.Font.Bold)
    isItalic = MsoTristateToBool(Shp.TextFrame.TextRange.Font.Italic)
    isUnderline = MsoTristateToBool(Shp.TextFrame.TextRange.Font.Underline)
    isStriken = MsoTristateToBool(Shp.TextFrame2.TextRange.Font.Strikethrough)
    With Slide1.AxTextBox
        .Left = Shp.Left
        .Top = Shp.Top
        .Width = Shp.Width
        .Height = Shp.Height
        .BackColor = Shp.Fill.ForeColor
        .ForeColor = Shp.TextFrame.TextRange.Font.Color
        .TextAlign = Shp.TextFrame.TextRange.ParagraphFormat.Alignment
        .Text = Shp.TextFrame.TextRange.Text
        .Font.Bold = isBold
        .Font.Italic = isItalic
        .Font.Underline = isUnderline
        .Font.Strikethrough = isStriken
        .Font.Name = Shp.TextFrame.TextRange.Font.Name
        .Font.Size = Shp.TextFrame.TextRange.Font.Size
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


Sub CheckSize(Shp As Shape)
    On Error GoTo Crash
    If AAX Then
        Slide1.AxTextBox.Visible = False
    End If
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then Slide13.AxTextBox.Visible = False
    If Shp.TextFrame.TextRange.Text = "Workspace 1" Then
        For Each Shp In Slide1.Shapes
            If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                If Shp.Name <> "RegularApp:-1" Then
                    Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                    AppID = GetAppID(Shp)
                    If Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.4 Then
                        FocusWindow (AppID)
                    End If
                End If
            End If
        Next Shp
        Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 2"
        LogData "Switched to Workspace 2"
    ElseIf Shp.TextFrame.TextRange.Text = "Workspace 2" Then
        For Each Shp In Slide1.Shapes
            If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                If Shp.Name <> "RegularApp:-1" Then
                    Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                    AppID = GetAppID(Shp)
                    If Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.4 Then
                        FocusWindow (AppID)
                    End If
                End If
            End If
        Next Shp
        Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 3"
        LogData "Switched to Workspace 3"
    ElseIf Shp.TextFrame.TextRange.Text = "Workspace 3" Then
        For Each Shp In Slide1.Shapes
            If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                If Shp.Name <> "RegularApp:-1" Then
                    Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                    AppID = GetAppID(Shp)
                    If Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.4 Then
                        FocusWindow (AppID)
                    End If
                End If
            End If
        Next Shp
        Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 4"
        LogData "Switched to Workspace 4"
    Else
        For Each Shp In Slide1.Shapes
            If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                If Shp.Name <> "RegularApp:-1" Then
                    Shp.Left = Shp.Left + ActivePresentation.SlideShowWindow.Width * 3
                    AppID = GetAppID(Shp)
                    If Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.4 Then
                        FocusWindow (AppID)
                    End If
                End If
            End If
        Next Shp
        Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 1"
        LogData "Switched to Workspace 1"
    End If
    Exit Sub
Crash:
    OSCrash "WORKSPACE_SWITCHER_FAILED", Err
End Sub

Sub Hibernate()
    ' Display "Hibernating..." screen
    LogData "Hibernation requested"
    ActivePresentation.SlideShowWindow.View.GotoSlide 21
    ActivePresentation.SlideShowWindow.Activate
    ' Set hibernation flags
    Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "True"
    Slide3.Shapes("BootText").TextFrame.TextRange.Text = "Resuming from hibernation..."
    Slide2.Shapes("BootText").TextFrame.TextRange.Text = "Window templates"
    Slide3.Shapes("Bootlogo").Visible = msoTrue
    Slide3.Shapes("Bootwarning").Visible = msoFalse
    ' Save and exit
    SavePresentation
    ActivePresentation.SlideShowWindow.View.Exit
End Sub

Sub MacroTest()
    Slide8.Shapes("BootText").TextFrame.TextRange.Text = "Macro test success"
    Pause (1)
    Slide8.Shapes("BootText").TextFrame.TextRange.Text = "Recovery mode"
End Sub


Sub RecoveryAbout()
    Slide8.Shapes("BootText").TextFrame.TextRange.Text = "Recovery mode for Sunlight 1.0 by mmaal [Fading stars]" ' f***ing stars?
    Pause (5)
    Slide8.Shapes("BootText").TextFrame.TextRange.Text = "Recovery mode"
End Sub

' Only use in single user mode!
Sub RecoverSession()
    Dim CanRecover As Boolean
    For Each oshp In Slide1.Shapes
        If InStr(oshp.Name, "RegularApp:") And oshp.Name <> "RegularApp:-1" Then
            CanRecover = True
        End If
    Next
    If CanRecover Then
        Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "True"
        Slide3.Shapes("BootText").TextFrame.TextRange.Text = "Attempting session recovery..."
        ActivePresentation.SlideShowWindow.Activate
        Slide3.Shapes("Bootlogo").Visible = msoTrue
        Slide3.Shapes("Bootwarning").Visible = msoFalse
        LogData "Session recovered"
        ActivePresentation.SlideShowWindow.View.GotoSlide (1)
    Else
        Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "False"
        Slide3.Shapes("BootText").TextFrame.TextRange.Text = "No recoverable data found, booting normally..."
        ActivePresentation.SlideShowWindow.Activate
        Slide3.Shapes("Bootlogo").Visible = msoTrue
        Slide3.Shapes("Bootwarning").Visible = msoFalse
        ActivePresentation.SlideShowWindow.View.GotoSlide (1)
    End If
End Sub

Sub TaskmgrRefresh(ByVal Shp As Shape)
    AppID = GetAppID(Shp)
    LogData "Task list refreshed for window with ID " & AppID
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
    For Each oshp In Slide1.Shapes
        If InStr(oshp.Name, "RegularApp:") Then
            SplitZ = Split(oshp.Name, ":")
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
    On Error Resume Next
    PageChange oSW
    Exit Sub
ReportIssue:
    OSCrash Err.Description, Err
End Sub


Sub Slide2Run()
    Slide1.Shapes("MoveEvent").TextFrame.TextRange.Text = "False"
    Slide1.Shapes("ResizeEvent").TextFrame.TextRange.Text = "False"
    Slide1.Shapes("AppID").TextFrame.TextRange.Text = "1"
    Slide1.Shapes("Username").TextFrame.TextRange.Text = "Nobody"
    DeleteDir "/Temp/"
    NewFolder "/Temp"
    If AAX Then
        Slide1.AxTextBox.Visible = False
    End If
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then Slide13.AxTextBox.Visible = False
    ResetWindows
    Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 1"
    Dim Factory As Boolean
    Dim Shp As Shape
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
            OSCrash "SYSTEM_CONFIGURATION_IS_CORRUPT", Err
        End If
    End If
    Slide7.Shapes("BootText").TextFrame.TextRange.Text = "Shutting down..."
    Slide2.Shapes("BootText").TextFrame.TextRange.Text = "Shutting down..."
End Sub

Sub DebugPageChange()
    PageChange ActivePresentation.SlideShowWindow
End Sub

Function GetBuildNo() As String
    GetBuildNo = "991"
End Function

Sub CheckDelay()
    Dim Delay As Integer
    Delay = GetBootDelay
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 1 Then
        Slide5.Shapes("BootloaderInfo").TextFrame.TextRange.Text = "Loader: LightBoot 3.0" & vbNewLine & "Date: 2024-12-08" & vbNewLine & "Boot delay: " & Delay & " second(s)" & vbNewLine & "Macros enabled"
    Else
        Slide5.Shapes("BootloaderInfo").TextFrame.TextRange.Text = "Loader: LightBoot 3.0" & vbNewLine & "Date: 2024-12-08" & vbNewLine & "Boot delay: " & Delay & " second(s)" & vbNewLine & "Macros disabled"
    End If
    For i = Slide5.timeLine.MainSequence.Count To 1 Step -1
        Dim eff As Effect
        Set eff = Slide5.timeLine.MainSequence(i)
        eff.Timing.TriggerDelayTime = Delay
    Next i
End Sub

Sub SetBootDelay(Secs As Integer)
    Slide5.Shapes("Bootdelay").TextFrame.TextRange.Text = CStr(Secs)
End Sub

Function GetBootDelay() As Integer
    GetBootDelay = CInt(Slide5.Shapes("Bootdelay").TextFrame.TextRange.Text)
End Function

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
        Dim BuildNo As String
        Slide3.Shapes("Wdym").Visible = msoFalse
        Slide3.Shapes("RecoveryModeBtn").Visible = msoFalse
        BuildNo = GetBuildNo
        CheckDelay
        Slide5.Shapes("MacroTest").TextFrame.TextRange.Text = "Macros enabled"
        If CheckVars("%ShowBuild%") <> "False" Then
            Dim BuildStr As String
            BuildStr = "Codename OrangePath OS" + vbNewLine + "Build " + BuildNo + vbNewLine + "For evaluation purposes only"
            Slide1.Shapes("BuildInfo").TextFrame.TextRange.Text = BuildStr
            Slide3.Shapes("BuildInfo").TextFrame.TextRange.Text = BuildStr
            Slide5.Shapes("BuildInfo").TextFrame.TextRange.Text = BuildStr
            Slide7.Shapes("BuildInfo").TextFrame.TextRange.Text = BuildStr
        Else
            BuildStr = "Sunlight OS" + vbNewLine + "Build " + BuildNo + vbNewLine
            Slide1.Shapes("BuildInfo").TextFrame.TextRange.Text = BuildStr
            Slide3.Shapes("BuildInfo").TextFrame.TextRange.Text = BuildStr
            Slide5.Shapes("BuildInfo").TextFrame.TextRange.Text = BuildStr
            Slide7.Shapes("BuildInfo").TextFrame.TextRange.Text = BuildStr
        End If
        Slide22.Shapes("Details").TextFrame.TextRange.Text = " "
        If InStr(Application.ActivePresentation.Name, "ForceRecovery") Then
            Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "Recovery"
        End If
        Slide2.Shapes("BootText").TextFrame.TextRange.Text = "Shutting down..."
        If Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "True" Then
            Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "False"
            ActivePresentation.SlideShowWindow.View.GotoSlide (4)
            UpdateTime
            Slide3.Shapes("BootText").TextFrame.TextRange.Text = "Startup error"
            Slide5.Shapes("BootText").TextFrame.TextRange.Text = "Startup error"
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
            Slide3.Shapes("BootText").TextFrame.TextRange.Text = "Starting Sunlight OS"
            Slide5.Shapes("BootText").TextFrame.TextRange.Text = "Starting Sunlight OS"
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
        CheckDelay
        Slide3.Shapes("Wdym").Visible = msoFalse
        Slide3.Shapes("RecoveryModeBtn").Visible = msoFalse
        Slide2Run
    ElseIf oSW.View.CurrentShowPosition = 3 Then
        Slide3.Shapes("BootText").TextFrame.TextRange.Text = "Startup error"
        Slide5.Shapes("BootText").TextFrame.TextRange.Text = "Startup error"
        Slide3.Shapes("Wdym").Visible = msoTrue
        Slide3.Shapes("RecoveryModeBtn").Visible = msoTrue
        Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "Recovery"
        Slide8.Shapes("BootText").TextFrame.TextRange.Text = "System was not shut down correctly. How would you like to proceed?"
        Slide8.Shapes("BtnSessionRecovery").TextFrame.TextRange.Text = "Attempt session recovery"
        Slide3.Shapes("Bootwarning").Visible = msoTrue
        Slide3.Shapes("Bootlogo").Visible = msoFalse
        Dim oeff As Effect
        For i = Slide3.timeLine.MainSequence.Count To 1 Step -1
            Set oeff = Slide3.timeLine.MainSequence(i)
            oeff.Delete
        Next i
        For i = Slide6.timeLine.MainSequence.Count To 1 Step -1
            Set oeff = Slide3.timeLine.MainSequence(i)
            oeff.Delete
        Next i
    ElseIf oSW.View.CurrentShowPosition = 4 Then
        ManualClockUpdate
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
        Slide8.Shapes("BtnSessionRecovery").TextFrame.TextRange.Text = "Continue boot"
        If Slide7.Shapes("Restart").TextFrame.TextRange.Text = "True" Then
            Slide7.Shapes("Restart").TextFrame.TextRange.Text = "False"
            Slide7.SlideShowTransition.AdvanceOnTime = msoTrue
            SavePresentation
            Slide3.Shapes("BootText").TextFrame.TextRange.Text = "Starting Sunlight OS"
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
            SavePresentation
            ActivePresentation.SlideShowWindow.View.Exit
        End If
    ElseIf oSW.View.CurrentShowPosition = 11 Then
        If Slide12.Shapes("FirmwareSource").TextFrame.TextRange.Text <> "*" Then
            Slide12.Shapes("StatusText").TextFrame.TextRange.Text = "Updating system..."
            Slide12.Shapes("Notice").TextFrame.TextRange.Text = "Do not close the presentation!"
            ActivePresentation.SlideShowWindow.Activate
            ResetWindows
            Pause (1)
            UpdateSystem
            Restart
        Else
            Slide12.Shapes("StatusText").TextFrame.TextRange.Text = "Firmware location not specified"
            Slide12.Shapes("Notice").TextFrame.TextRange.Text = "Returning to recovery mode in 5 seconds..."
            Slide8.Shapes("BootText").TextFrame.TextRange.Text = "System update failed. How would you like to proceed?"
            ActivePresentation.SlideShowWindow.Activate
            Pause 5
            ActivePresentation.SlideShowWindow.View.GotoSlide 8
        End If
    ElseIf oSW.View.CurrentShowPosition = 12 Then
        Slide3.Shapes("BootText").TextFrame.TextRange.Text = "Startup error"
        Slide5.Shapes("BootText").TextFrame.TextRange.Text = "Startup error"
        Slide3.Shapes("Wdym").Visible = msoTrue
        Slide3.Shapes("RecoveryModeBtn").Visible = msoTrue
        Slide3.Shapes("Bootwarning").Visible = msoTrue
        Slide3.Shapes("Bootlogo").Visible = msoFalse
    ElseIf oSW.View.CurrentShowPosition = 19 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide (31)
    ElseIf oSW.View.CurrentShowPosition = 20 Then
        If Slide19.Shapes("MsgDisplayed").TextFrame.TextRange.Text = "False" Then
            Slide19.Shapes("MsgDisplayed").TextFrame.TextRange.Text = "True"
            AppMessage "This version of Codename Sunlight OS is not finished and therefore may be unstable. Please report any detected issues to the developer as soon as possible.", "Evaluation copy", "Info", False
        End If
    ElseIf oSW.View.CurrentShowPosition = 23 Then
        ' oh dear
        Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "Recovery"
        Slide8.Shapes("BootText").TextFrame.TextRange.Text = "System was shut down because of an error. How would you like to proceed?"
        Slide8.Shapes("BtnSessionRecovery").TextFrame.TextRange.Text = "Attempt session recovery"
        Pause 2
        SavePresentation
        Pause 2
        ActivePresentation.SlideShowWindow.View.GotoSlide (24)
    ElseIf oSW.View.CurrentShowPosition = 26 Then
        ActivePresentation.SlideShowWindow.View.Last
    ElseIf oSW.View.CurrentShowPosition > 30 Then
        SetupPageChange oSW
    End If
End Sub


Sub CleanPopups()
    Dim Sld As Slide
    Dim Shp As Shape
    IDX = 1
    For Each Sld In ActivePresentation.Slides
        If IDX <> 4 And IDX <> 10 And IDX <> 13 And IDX <> 9 And IDX <> 24 And IDX <> 27 And IDX <> 26 And IDX <> 24 And IDX <> 29 And IDX <> 30 And IDX <> ActivePresentation.SlideShowWindow.View.CurrentShowPosition Then
            For Each Shp In Sld.Shapes
                If InStr(Shp.Name, ":") Then
                    Shp.Delete
                End If
            Next Shp
        End If
        IDX = IDX + 1
    Next Sld
    DeleteOrphanedTriggers
End Sub

Sub Restart()
    LogData "Restart requested"
    Slide2.Shapes("BootText").TextFrame.TextRange.Text = "Restarting..."
    Slide7.Shapes("BootText").TextFrame.TextRange.Text = "Restarting..."
    Slide7.Shapes("Restart").TextFrame.TextRange.Text = "True"
    ActivePresentation.SlideShowWindow.View.GotoSlide (5)
End Sub

Sub RestartRecovery()
    LogData "Restart to recovery requested"
    Slide2.Shapes("BootText").TextFrame.TextRange.Text = "Please wait..."
    Slide7.Shapes("BootText").TextFrame.TextRange.Text = "Please wait..."
    Slide7.Shapes("Restart").TextFrame.TextRange.Text = "Recovery"
    ActivePresentation.SlideShowWindow.View.GotoSlide (5)
End Sub

Sub EnterRecovery()
    LogData "Recovery request"
    Slide8.Shapes("BootText").TextFrame.TextRange.Text = "Recovery mode"
    Slide8.Shapes("BtnSessionRecovery").TextFrame.TextRange.Text = "Continue boot"
    ActivePresentation.SlideShowWindow.View.GotoSlide (8)
    CheckDelay
End Sub

' Set global variable
Sub SetVar(ByVal Key As String, ByVal Value As String)
    LogData "Setting global variable " & Key
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
    LogData "Unsetting variable  " & Key
    If ShapeExists(Slide21, Key) = True Then
        Slide21.Shapes(Key).Delete
    End If
End Sub


Function CheckVars(ByVal str As String)
    On Error Resume Next
    outStr = str
    For Each Shp In Slide21.Shapes
        If Shp.Name <> "__AxDummy" Then
            Dim preStr As String
            preStr = outStr
            outStr = Replace(outStr, "%" & Shp.Name & "%", Shp.TextFrame.TextRange.Text)
            If preStr <> outStr Then
                LogData "Found global variable " & Shp.Name
            End If
        End If
    Next Shp
    CheckVars = outStr
End Function

Sub Highlight(Shp As Shape)
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

Sub OPCrashTest()
    OSCrash "TESTING_EXCEPTION", Err
End Sub

Sub OSCrash(ByVal Details As String, Optional ByVal Error As ErrObject = Nothing)
    Slide22.Shapes("Details").TextFrame.TextRange.Text = "Error details: " & Details
    If CheckVars("%EnforceUserspaceExceptions%") <> "True" Then
        If Error Is Nothing Or Error.Description = "" Then
            Slide22.Shapes("Exception").TextFrame.TextRange.Text = "EXCEPTION DATA UNAVAILABLE"
        Else
            Slide22.Shapes("Exception").TextFrame.TextRange.Text = "Source: " & Error.Source & vbNewLine & "Description: " & Error.Description
        End If
        ActivePresentation.SlideShowWindow.View.GotoSlide 23
    Else
        If Details <> "MESSAGE_BOX_ERROR" Then
            If Error Is Nothing Or Error.Description = "" Then
                AppMessage "Fatal system error: " & Details & vbNewLine & Error.Description & vbNewLine, Error.Source, "Error", True
            Else
                AppMessage "Fatal system error: " & Details, "System error", "Error", True
            End If
        Else
            MsgBox "Fatal system error: " & Details & vbNewLine & Error.Description & vbNewLine, vbCritical, Error.Source
        End If
    End If
End Sub

' Deprecated, replaced by SetupExperience [StdModule]
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
    OSCrash "OUT_OF_BOX_EXPERIENCE_FAIL", Err
End Sub

Sub CheckUncheck(Shp As Shape)

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
    For i = 21 To 40 Step 1
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

Sub AboutDe()
    AppMessage "Sunlight Slide1 Environment" + vbNewLine + "Version 1.0 by mmaal" + vbNewLine + "What do we know about hidden dialogs?", "You found me!", "Info", True
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
    LogData "Factory reset requested"
    Dim PreShutdown As Shape
    Dim Splash As Shape
    Dim PreSplash As Shape
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
    Slide2.Shapes("ShpMisc6AppSettings_").TextFrame.TextRange.Text = "Enable"
    Slide1.Shapes("AppID").Visible = msoFalse
    Slide1.Shapes("MoveEvent").Visible = msoFalse
    Slide1.Shapes("ResizeEvent").Visible = msoFalse
    Slide1.Shapes("AppCreatingEvent").Visible = msoFalse
    Slide1.Shapes("AutosaveTime").Visible = msoFalse
    Slide1.Shapes("Username").Visible = msoFalse
    Slide1.Shapes("BuildInfo").Visible = msoFalse
    Slide3.Shapes("BuildInfo").Visible = msoFalse
    Slide5.Shapes("BuildInfo").Visible = msoFalse
    Slide7.Shapes("BuildInfo").Visible = msoFalse
    
    ' Clear desktop
    Slide1.Shapes("RegularApp:-1").Delete
    GetFileRef("/Defaults/Desktop.grp").Copy
    With Slide1.Shapes.Paste
        .ZOrder msoSendToBack
        .Name = "RegularApp:-1"
        .Visible = msoTrue
        Slide1.Shapes("BackgroundImg").ZOrder msoSendToBack
        Slide1.Shapes("AnimationRect").ZOrder msoSendToBack
    End With
    
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
    
    ' Reset boot delay
    SetBootDelay 1
    
    Slide3.Shapes("Hibernated").TextFrame.TextRange.Text = "Factory"
    ' Display message in recovery mode
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition <> 4 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 8
        Slide8.Shapes("BootText").TextFrame.TextRange.Text = "Factory reset success"
        Slide8.Shapes("BtnSessionRecovery").TextFrame.TextRange.Text = "Continue boot"
        Pause (1)
        Slide8.Shapes("BootText").TextFrame.TextRange.Text = "Recovery mode"
    End If
End Sub


Sub UpdateTime()
    On Error GoTo Crash
    'On Error Resume Next
    Slide1.Shapes("ResizeEvent").TextFrame.TextRange.Text = "False"
    Slide1.Shapes("MoveEvent").TextFrame.TextRange.Text = "False"
    invervalStr = GetFileContent("/System/Settings.cnf", "AutosaveInterval")
    'Dim target As Long
    'target = Minute(Time) + CLng(intervalStr)
    'If target >= 60 Then
    '    target = target - 60
    'End If
    'disableClock = True
    'Do
        fullClock = Split(Time, ":")
        hrMin = fullClock(0) & ":" & fullClock(1)
        'hrMin = Time
        Slide1.Shapes("Clock").TextFrame.TextRange.Text = hrMin
        'Slide1.Shapes("AutosaveTime").TextFrame.TextRange.Text = CStr(target)
        Slide1.Shapes("AutosaveTime").TextFrame.TextRange.Text = "N/A"
        intervalStr = GetFileContent("/System/Settings.cnf", "AutosaveInterval")
        'target = Minute(Time) + CLng(intervalStr)
        'If target >= 60 Then
        '    target = target - 60
        'End If
        If CInt(intervalStr) > 0 And CStr(Int(100 * Rnd) Mod 5) = "0" Then
            SavePresentation
        End If
        'If ActivePresentation.SlideShowWindow.View.CurrentShowPosition > 4 Or ActivePresentation.SlideShowWindow.View.CurrentShowPosition < 3 Then
            'Exit Do
        'End If
        'DoEvents
    'Loop
Done:
    Exit Sub
Crash:
    OSCrash "SYSTEM_WATCHDOG_ERROR", Err
End Sub

Sub ManualClockUpdate()
    fullClock = Split(Time, ":")
    hrMin = fullClock(0) & ":" & fullClock(1)
    Slide1.Shapes("Clock").TextFrame.TextRange.Text = hrMin
End Sub


Sub ResetWindows()
    If Slide7.Shapes("BootText").TextFrame.TextRange.Text = "It's now safe to close the presentation" Then Exit Sub
    LogData "Resetting windows"
    For i = 0 To 3
        Dim Shp As Shape
        Set Shp = Slide1.Shapes("SwitchWorkspace")
        'MsgBox (Shp.TextFrame.TextRange.Text)
        If Shp.TextFrame.TextRange.Text = "Workspace 1" Then
            For Each Shp In Slide1.Shapes
                If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                    If Shp.Name <> "RegularApp:-1" Then
                        Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                    End If
                End If
            Next Shp
            Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 2"
        ElseIf Shp.TextFrame.TextRange.Text = "Workspace 2" Then
            For Each Shp In Slide1.Shapes
                If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                    If Shp.Name <> "RegularApp:-1" Then
                        Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                    End If
                End If
            Next Shp
            Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 3"
        ElseIf Shp.TextFrame.TextRange.Text = "Workspace 3" Then
            For Each Shp In Slide1.Shapes
                If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                    If Shp.Name <> "RegularApp:-1" Then
                        Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                    End If
                End If
            Next Shp
            Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 4"
        Else
            For Each Shp In Slide1.Shapes
                If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                    If Shp.Name <> "RegularApp:-1" Then
                        Shp.Left = Shp.Left + ActivePresentation.SlideShowWindow.Width * 3
                    End If
                End If
            Next Shp
            Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 1"
        End If
        For Each Shp In Slide1.Shapes
            If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                If Shp.Name <> "RegularApp:-1" Then
                    Shp.Delete
                End If
            End If
        Next Shp
    Next i
End Sub

Function FocusWindow(ByVal AppID As String)
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 38 Then
        Exit Function
    ElseIf ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 13 Then
        Exit Function
    End If
    LogData "Focus " & AppID
    If AAX Then
        If Slide1.AxComboBox.Visible Then
            Slide1.AxComboBox.Visible = False
        End If
    End If
    For Each Shp In Slide1.Shapes
        If InStr(Shp.Name, "TaskIcon:") Then
            If InStr(Shp.Name, AppID) Then
                Shp.Fill.Transparency = 0.4
            Else
                Shp.Fill.Transparency = 0.8
            End If
        End If
        Dim HasAx As Boolean
        Dim MultiLineTb As Boolean
        HasAx = False
        MultiLineTb = False
        If InStr(Shp.Name, "RegularApp:" & AppID) Then
            For X = 1 To Shp.GroupItems.Count
                With Shp.GroupItems(X)
                    If InStr(.Name, "AXTextBox2") Then
                        MultiLineTb = True
                    End If
                    If InStr(.Name, "AXTextBox") Then
                        ApplyTbAttribs Shp.GroupItems(X)
                        HasAx = True
                    End If
                End With
            Next
        End If
        If AAX Then
            Slide1.AxTextBox.Visible = HasAx
            Slide1.AxTextBox.MultiLine = MultiLineTb
            Slide1.AxTextBox.EnterKeyBehavior = MultiLineTb
            Slide1.AxTextBox.TabKeyBehavior = MultiLineTb
            If MultiLineTb Then
                Slide1.AxTextBox.ScrollBars = fmScrollBarsBoth
            Else
                Slide1.AxTextBox.ScrollBars = fmScrollBarsNone
            End If
        End If
        ActivePresentation.SlideShowWindow.View.GotoSlide (ActivePresentation.SlideShowWindow.View.CurrentShowPosition)
    Next Shp
End Function

Sub MinimizeWindow(Shp As Shape)
    AppID = GetAppID(Shp)
    FocusWindow AppID
    LogData "Minimize requested for ID " & AppID
    MinimizeRestore Slide1.Shapes("TaskIcon:" & AppID)
End Sub

Sub MinimizeRestore(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    If Slide1.Shapes("RegularApp:" & AppID).Visible = msoTrue Then
        If Shp.Fill.Transparency = 0.8 Then
            Slide1.Shapes("RegularApp:" & AppID).ZOrder msoBringToFront
            FocusWindow AppID
            TryRunMacro Slide1.Shapes("TaskIcon:" & AppID).TextFrame.TextRange.Text, "Focus", AppID
            UpdateTime
        Else
            LogData "Minimized window with ID " & AppID
            Slide1.Shapes("RegularApp:" & AppID).Visible = msoFalse
            Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.8
            If AAX Then
                Slide1.AxTextBox.Visible = False
            End If
            TryRunMacro Slide1.Shapes("TaskIcon:" & AppID).TextFrame.TextRange.Text, "Minimize", AppID
            ActivePresentation.SlideShowWindow.View.GotoSlide (4)
            UpdateTime
        End If
    Else
        LogData "Restoring window with ID " & AppID
        Slide1.Shapes("RegularApp:" & AppID).Visible = msoTrue
        Slide1.Shapes("RegularApp:" & AppID).ZOrder msoBringToFront
        FocusWindow AppID
        Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.4
        
        TryRunMacro Slide1.Shapes("TaskIcon:" & AppID).TextFrame.TextRange.Text, "Restore", AppID
        
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

Sub TryRunMacro(AppLabel As String, RoutineLabel As String, AppID As String)
    On Error GoTo ExitSub
    Application.Run "App" & AppLabel & RoutineLabel, AppID
    LogData "Executed " & RoutineLabel & " hook for window with ID " & AppID
ExitSub:
    Exit Sub
End Sub

Sub InvertValue(Shp As Shape)
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
    LogData "Rearranged taskbar labels"
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
    LogData "Write system configuration with key " & Key
End Function

Function GetSysConfig(ByVal Name As String) As String
    GetSysConfig = GetFileContent("/System/Settings.cnf", Name)
    LogData "Read system configuration with key " & Name
End Function


Sub SavePresentation()
    On Error GoTo Crash
    With Application.ActivePresentation
        ' save only if the save path is known and there are unsaved changes
        If Not .Saved And .Path <> "" Then
            LogData "Saving presentation"
            If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 4 Then
                WaitCursor Nothing, "Saving..."
                HideCursor
            End If
            DeleteOrphanedTriggers
            .Save
        End If
    End With
    
Done:
    Exit Sub
Crash:
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition <> 22 Then
        OSCrash "SAVE_ERROR", Err
    End If
End Sub

' Returns the App ID based on the clicked/hovered shape name
Function GetAppID(ByVal Shp As Shape) As String
    On Error GoTo ReturnNothing
    SplitZ = Split(Shp.Name, ":")
    AppID = SplitZ(1)
    GetAppID = AppID
    Exit Function
ReturnNothing:
    GetAppID = ""
End Function

' Regroups ungrouped windows
Sub Regroup(ByVal AppID As String, ByVal Sld As Slide)
    On Error GoTo NoTY
    LogData "Regrouping window with ID " & AppID
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
NoTY:
End Sub

' Regroups any shape with _ at the end
Sub RegroupShapes(ByVal ShapeName As String, ByVal Sld As Slide)
    LogData "Grouping orphaned shape " & ShapeName
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
Sub PasteToGroup(ByVal Ref As Shape, _
                ByVal Shp As Shape, _
                ByVal Name As String, _
                ByVal OffsetX As Integer, _
                ByVal OffsetY As Integer, _
                ByVal Sld As Slide, _
                Optional ByVal Macro As String = "")
    ' Declarations
    Dim AppID As String
    Dim Left As Integer
    Dim Top As Integer
    ' Get AppID from reference shape
    AppID = GetAppID(Ref)
    LogData "Pasting group to window with ID " & AppID
    If AppID = "" Then Exit Sub
    If AppID = "Icon" Then Exit Sub
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
                Dim Shp2 As Shape
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
Sub EraseFromGroup(ByVal Ref As Shape, ByVal ShpName As String, ByVal Sld As Slide)
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

Sub SkinShapeLo(ByVal Sld As Slide, ByVal Parent As String, ByVal Replacable As String, ByVal ReplacedBy As Shape)
    Sld.Shapes(Parent).Ungroup
    Sld.Shapes(Replacable).Delete
    ReplacedBy.Copy
    With Sld.Shapes.Paste
        .Name = Replacable
        .ZOrder msoSendToBack
    End With
    RegroupShapes Parent, Sld
End Sub

Sub SkinShape(ByVal Ref As Shape, ByVal control As String, ByVal Thm As Shape, ByVal Sld As Slide)
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

' Deprecated, will break stuff
Sub ApplyTheme(ByVal Filename As String)
    'Slide2.Shapes("Icon4Part1AppMenu_").ZOrder msoSendToBack
    Dim Thm As Shape
    Dim W As Integer
    Dim H As Integer
    Dim X As Integer
    Dim Y As Integer
    Set Thm = GetFileRef(Filename)
    Dim ShpArr() As Shape
    
    
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

