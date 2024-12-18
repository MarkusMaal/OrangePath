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
    Dim I As Integer
    
    Set Slide = Slide1
    Set timeLine = Slide.timeLine
    If timeLine.InteractiveSequences.Count > 0 Then
        Set mainSeq = timeLine.InteractiveSequences
        Dim seq As Sequence
        For I = mainSeq.Count To 1 Step -1
            Set seq = mainSeq.Item(I)
            If Slide.Shapes(seq.Item(1).DisplayName).Visible = msoFalse Then
                Slide.Shapes(seq.Item(1).DisplayName).Delete
            End If
        Next I
    End If
End Sub

Sub CompensateText(Shp As Shape, Text As String)
    Shp.TextFrame.TextRange.Text = Text
    Do While Shp.Height > Shp.TextFrame.TextRange.Font.Size And Len(Text) > 2
        Text = Right(Text, Len(Text) - 1)
    Loop
End Sub

Sub TestReloadNeg1()
    Reload "-1"
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
    Slide17.UsernameFIeld.Text = ""
    Slide17.PassField.Text = ""
    Slide17.ConfirmPassField.Text = ""
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
    On Error GoTo HideCursor
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
    Exit Sub
HideCursor:
    Slide1.Shapes("WaitPlease").Visible = msoFalse
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

Sub ReorganizeITaskIcons()
    Dim Shp2 As Shape
    Dim ShpL As Shape
    Dim AppID As String
    Dim ASplit() As String
    Dim IDX As Integer
    For IDX = Slide1.Shapes.Count To 1 Step -1
        Set Shp2 = Slide1.Shapes(IDX)
        If InStr(1, Shp2.Name, "ITaskIcon:") Then
            ASplit = Split(Shp2.Name, ":")
            AppID = ASplit(1)
            Set ShpL = Slide1.Shapes("TaskIcon:" & AppID)
            Shp2.Top = ShpL.Top + 3
            
            Shp2.Left = CInt(ShpL.Left) + 4
        End If
    Next IDX
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
    On Error Resume Next
    
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
                If InStr(1, Shp.Name, "TaskIcon:") = 1 And Shp.Left > 0 And Shp.Left < ActivePresentation.SlideShowWindow.Width Then
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
        
        If ShapeExists(Slide25, "App" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & ":Icon") Then
            Slide25.Shapes("App" & Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text & ":Icon").Copy
        Else
            Slide29.Shapes("AppDefault:Icon").Copy
        End If
            With Slide1.Shapes.Paste
                .Left = Slide1.Shapes("TaskIcon:" & AppID).Left + 4
                .Top = Slide1.Shapes("TaskIcon:" & AppID).Top + 3
                .Height = Slide1.Shapes("TaskIcon:" & AppID).Height - 8
                .Width = .Height
                .Visible = msoTrue
                .Name = "ITaskIcon:" & AppID
                For Each Shp In .GroupItems
                    If Shp.TextFrame.TextRange.Text <> "" And Shp.Name <> "Background" Then
                        If Shp.Fill.Transparency = 1 Then
                            Dim TargetLeft As Integer
                            TargetLeft = .Left
                            Shp.Width = Shp.Width + 10
                            Shp.Left = Shp.Left - 5
                            Do While Shp.Left < TargetLeft
                                Shp.Left = Shp.Left + 0.5
                                Shp.Width = Shp.Width - 1
                                Shp.Top = Shp.Top + 0.1
                                Shp.TextFrame.TextRange.Font.Size = Shp.TextFrame.TextRange.Font.Size - 1
                            Loop
                            Do While .Top + Shp.Top + Shp.Height = .Top + .Height
                                Shp.TextFrame.TextRange.Font.Size = Shp.TextFrame.TextRange.Font.Size - 0.1
                            Loop
                        Else
                            Shp.TextFrame.TextRange.Text = ""
                        End If
                    End If
                    Shp.TextFrame.TextRange.Font.Size = Shp.TextFrame.TextRange.Font.Size / 4
                    Shp.ActionSettings(ppMouseClick).Run = "MinimizeRestore"
                Next Shp
                .ZOrder msoBringToFront
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
    ' Check for custom file associations
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
        ' Apply visual style
        If FileStreamsExist("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm") Then
            GetFileRef("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm").Copy
        Else
            GetFileRef("/Defaults/Themes/Default.thm").Copy
        End If
        With CurrentSld.Shapes.Paste
            .GroupItems("Handle").Delete
            .GroupItems("Button").Delete
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
            If InStr(1, Shp.Name, "Grp") = 1 Then
                Dim SubShp As Shape
                For Each SubShp In Shp.GroupItems
                    If InStr(1, SubShp.Name, "Button") Then
                        CopyFillFormat GetTheme().GroupItems("Button"), SubShp
                    End If
                Next SubShp
            End If
            If InStr(1, Shp.Name, "Button") Then
                CopyFillFormat GetTheme().GroupItems("Button"), Shp
            End If
            Shp.Name = SplitName(0) & ":" & AppID
            If InStr(Shp.Name, "AXTextBox") Then ApplyTbAttribs Shp
            Shapes = Shapes & Shp.Name & ","
        End If
    Next
    
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
    For I = ActivePresentation.Slides(CurrentSlide).timeLine.MainSequence.Count To 1 Step -1
        Set eff = Slide1.timeLine.MainSequence(I)
        eff.Delete
    Next I
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
        For I = ActivePresentation.Slides(CurrentSlide).timeLine.MainSequence.Count To 1 Step -1
            Set eff = Slide1.timeLine.MainSequence(I)
            eff.Delete
        Next I
        ActivePresentation.Slides(CurrentSlide).Shapes("RegularApp:" & AppID).Visible = msoFalse
        ActivePresentation.SlideShowWindow.Activate
        ActivePresentation.Slides(CurrentSlide).Shapes("RegularApp:" & AppID).Visible = msoTrue
        ActivePresentation.SlideShowWindow.Activate
        ActivePresentation.SlideShowWindow.View.GotoSlide (CurrentSlide)
    End If
End Sub

Public Function IsInArray(stringToBeFound As String, arr() As String) As Boolean
    Dim I
    For I = LBound(arr) To UBound(arr)
        If arr(I) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next I
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
                If ShapeExists(ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition), "ITaskIcon:" & AppID) Then
                    ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.CurrentShowPosition).Shapes("ITaskIcon:" & AppID).Delete
                End If
                If TaskIcon = 373 Then
                    TaskIcon = 3
                ElseIf TaskIcon = 527 Then
                    TaskIcon = 4
                ElseIf TaskIcon = 219 Then
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
        For I = Slide1.timeLine.MainSequence.Count To 1 Step -1
            Set eff = Slide1.timeLine.MainSequence(I)
            eff.Delete
        Next I
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
            If TaskIcon = 373 Then
                TaskIcon = 3
            ElseIf TaskIcon = 527 Then
                TaskIcon = 4
            ElseIf TaskIcon = 219 Then
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
            If TaskIcon = 373 Then
                TaskIcon = 3
            ElseIf TaskIcon = 527 Then
                TaskIcon = 4
            ElseIf TaskIcon = 219 Then
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
    If TaskIcon = 373 Then
        TaskIcon = 3
    ElseIf TaskIcon = 527 Then
        TaskIcon = 4
    ElseIf TaskIcon = 219 Then
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
        If HasModals(GetAppID(Shp), True) Then Exit Sub
        If Slide1.AxComboBox.Visible Then
            Slide1.AxComboBox.Visible = False
        End If
    End If
    Dim LoopState As Boolean
    SplitZ = Split(Shp.Name, ":")
    AppID = SplitZ(1)
    LogData "Starting move event for window with ID " & AppID
    For I = Slide1.timeLine.MainSequence.Count To 1 Step -1
        Set oeff = Slide1.timeLine.MainSequence(I)
        oeff.Delete
    Next I
    
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
            
            ' * Check bounds *
            Dim TB As Shape
            Dim RA As Shape
            
            Set TB = ActivePresentation.Slides(CurrentSlide).Shapes("Taskbar")
            Set RA = ActivePresentation.Slides(CurrentSlide).Shapes("RegularApp:" & AppID)
            
            Dim TBT As Single
            Dim RAT As Single
            Dim RAH As Single
            Dim RAL As Single
            Dim RAW As Single
            Dim SW As Single
            Dim SH As Single
            
            TBT = TB.Top
            RAT = RA.Top
            RAH = RA.Height
            RAL = RA.Left
            RAW = RA.Width
            SW = ActivePresentation.PageSetup.SlideWidth
            SH = ActivePresentation.PageSetup.SlideHeight
            
            If RAT - 4 > TBT - RAH Then RA.Top = TBT - RAH + 4
            If RAL + RAW > SW Then RA.Left = SW - RAW
            If RAT < 0 Then RA.Top = 0
            If RAL < 0 Then RA.Left = 0
            ' * End check bounds *
            
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
            LoopState = False
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
    If AAX Then
        If HasModals(GetAppID(Shp), True) Then Exit Sub
    End If
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
            If Slide1.Shapes("RegularApp:" & AppID).Height < 100 Then Slide1.Shapes("RegularApp:" & AppID).Height = 100
            If Slide1.Shapes("RegularApp:" & AppID).Width < 100 Then Slide1.Shapes("RegularApp:" & AppID).Width = 100
            If Slide1.Shapes("RegularApp:" & AppID).Width > ActivePresentation.PageSetup.SlideWidth Then Slide1.Shapes("RegularApp:" & AppID).Width = ActivePresentation.PageSetup.SlideWidth
            If Slide1.Shapes("RegularApp:" & AppID).Height > ActivePresentation.PageSetup.SlideHeight Then Slide1.Shapes("RegularApp:" & AppID).Height = ActivePresentation.PageSetup.SlideHeight
            
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
            dw = Slide1.Shapes("RegularApp:" & AppID).Width
            dh = Slide1.Shapes("RegularApp:" & AppID).Height
            Slide1.Shapes("ResizeEvent").TextFrame.TextRange.Text = "True"
        ' GrabArea not clicked, window was resized
        ElseIf LoopState = True And GetAsyncKeyState1 = False Then
            LoopState = False
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
                For I = Slide1.Shapes("RegularApp:" & AppID).GroupItems.Count To 0 Step -1
                    With .GroupItems(I)
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
                TryRunMacro GetAppName(Slide1.Shapes("RegularApp:" & AppID).GroupItems(1).Name), "SizeChanged", CInt(AppID)
                For I = Slide1.Shapes("RegularApp:" & AppID).GroupItems.Count To 0 Step -1
                    With .GroupItems(I)
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
    Dim HasModal As Boolean
    Dim IsModal As Boolean
    
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
    Dim I As Integer
    For I = 1 To 4
        Slide1.Shapes("WorkspaceCircle" & CStr(I)).Fill.Transparency = 0.5
    Next I
    If Shp.TextFrame.TextRange.Text = "Workspace 1" Then
        For Each Shp In Slide1.Shapes
            If InStr(Shp.Name, "RegularApp:") Or InStr(1, Shp.Name, "TaskIcon:") Then
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
        Slide1.Shapes("WorkspaceCircle2").Fill.Transparency = 0
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
        Slide1.Shapes("WorkspaceCircle3").Fill.Transparency = 0
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
        Slide1.Shapes("WorkspaceCircle4").Fill.Transparency = 0
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
        Slide1.Shapes("WorkspaceCircle1").Fill.Transparency = 0
        LogData "Switched to Workspace 1"
    End If
    Exit Sub
Crash:
    OSCrash "WORKSPACE_SWITCHER_FAILED", Err
End Sub

Sub ToLastWorkspace()
    Dim I As Integer
    For I = 1 To 3
        CheckSize Slide1.Shapes("SwitchWorkspace")
    Next I
End Sub

Sub ToNextWorkspace()
    CheckSize Slide1.Shapes("SwitchWorkspace")
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
    I = 1
    For Each oshp In Slide1.Shapes
        If InStr(oshp.Name, "RegularApp:") Then
            SplitZ = Split(oshp.Name, ":")
            AppID = SplitZ(1)
            With Slide1.Shapes("RegularApp:" & AppID)
                If .GroupItems.Count > 0 Then
                    AppNameSplit = Split(.GroupItems(1).Name, ":")
                    AppNameSplit2 = Split(AppNameSplit(0), "App")
                    AppName = AppNameSplit2(1)
                    If I < 11 Then
                        Dim hp As Boolean
                        hp = False
                        With Slide1.Shapes("RegularApp:" & TaskMgrID)
                            For Each GI In .GroupItems
                                If InStr(GI.Name, "Proc") And GI.TextFrame.TextRange.Text = "" And Not hp Then
                                    GI.TextFrame.TextRange.Text = AppName & " (PID: " & CStr(AppID) & ")"
                                    I = I + 1
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


Function GetAppName(ShapeName As String) As String
    AppNameSplit = Split(ShapeName, ":")
    AppNameSplit2 = Split(AppNameSplit(0), "App")
    GetAppName = AppNameSplit2(1)
End Function

Sub OnSlideShowPageChange(ByVal oSW As SlideShowWindow)
    On Error GoTo ExitSub
    PageChange oSW
    Exit Sub
ExitSub:
    Exit Sub
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
    Slide1.Shapes("WorkspaceCircle1").Fill.Transparency = 0
    Slide1.Shapes("WorkspaceCircle2").Fill.Transparency = 0.5
    Slide1.Shapes("WorkspaceCircle3").Fill.Transparency = 0.5
    Slide1.Shapes("WorkspaceCircle4").Fill.Transparency = 0.5
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
    GetBuildNo = "1001"
End Function

Sub CheckDelay()
    Dim Delay As Integer
    Delay = GetBootDelay
    If ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 1 Then
        Slide5.Shapes("BootloaderInfo").TextFrame.TextRange.Text = "Loader: LightBoot 3.0" & vbNewLine & "Date: 2024-12-08" & vbNewLine & "Boot delay: " & Delay & " second(s)" & vbNewLine & "Macros enabled"
    Else
        Slide5.Shapes("BootloaderInfo").TextFrame.TextRange.Text = "Loader: LightBoot 3.0" & vbNewLine & "Date: 2024-12-08" & vbNewLine & "Boot delay: " & Delay & " second(s)" & vbNewLine & "Macros disabled"
    End If
    For I = Slide5.timeLine.MainSequence.Count To 1 Step -1
        Dim eff As Effect
        Set eff = Slide5.timeLine.MainSequence(I)
        eff.Timing.TriggerDelayTime = Delay
    Next I
End Sub

Sub SetBootDelay(Secs As Integer)
    Slide5.Shapes("Bootdelay").TextFrame.TextRange.Text = CStr(Secs)
End Sub

Function GetBootDelay() As Integer
    GetBootDelay = CInt(Slide5.Shapes("Bootdelay").TextFrame.TextRange.Text)
End Function

' Apply diamond transition for old versions of PowerPoint
Sub CheckOfficeVer()
    Dim transEff As PpEntryEffect
    
    ' Office 2013 and earlier
    If CSng(Replace(Application.Version, ".", ",")) < 16# Then
        transEff = &HF22 ' Diamond glitter to right
    Else
        transEff = &HF72 ' Morph object
    End If
    Slide2.SlideShowTransition.EntryEffect = transEff
    Slide3.SlideShowTransition.EntryEffect = transEff
    Slide4.SlideShowTransition.EntryEffect = transEff
    Slide7.SlideShowTransition.EntryEffect = transEff
    Slide11.SlideShowTransition.EntryEffect = transEff
    Slide16.SlideShowTransition.EntryEffect = transEff
    Slide18.SlideShowTransition.EntryEffect = transEff
    Slide21.SlideShowTransition.EntryEffect = transEff
    Slide23.SlideShowTransition.EntryEffect = transEff
    Slide25.SlideShowTransition.EntryEffect = transEff
    Slide31.SlideShowTransition.EntryEffect = transEff
    Slide34.SlideShowTransition.EntryEffect = transEff
    Slide37.SlideShowTransition.EntryEffect = transEff
    Slide38.SlideShowTransition.EntryEffect = transEff
    Slide39.SlideShowTransition.EntryEffect = transEff
    
    ' Reset global variables
    Dim IDX As Integer
    Dim Shp As Shape
      
    For IDX = Slide21.Shapes.Count To 1 Step -1
        Set Shp = Slide21.Shapes(IDX)
        If Shp.Name <> "Label1" Then
            Shp.Delete
        End If
    Next IDX
End Sub

Sub PageChange(ByVal oSW As SlideShowWindow)
    If oSW.View.CurrentShowPosition = 6 And Slide7.Shapes("BootText").TextFrame.TextRange.Text = "It's now safe to close the presentation" Then
        Exit Sub
    End If
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
        CheckOfficeVer
        Slide7.Shapes("BootText").TextFrame.TextRange.Text = "Shutting down..."
        Dim BuildNo As String
        Slide3.Shapes("Wdym").Visible = msoFalse
        Slide3.Shapes("RecoveryModeBtn").Visible = msoFalse
        BuildNo = GetBuildNo
        CheckDelay
        Slide5.Shapes("MacroTest").TextFrame.TextRange.Text = "Macros enabled"
        If CheckVars("%ShowBuild%") <> "False" Then
            Dim BuildStr As String
            BuildStr = "Sunlight OS" + vbNewLine + "Build " + BuildNo + vbNewLine + "Test mode"
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
        For I = Slide3.timeLine.MainSequence.Count To 1 Step -1
            Set oeff = Slide3.timeLine.MainSequence(I)
            oeff.Delete
        Next I
        For I = Slide6.timeLine.MainSequence.Count To 1 Step -1
            Set oeff = Slide3.timeLine.MainSequence(I)
            oeff.Delete
        Next I
    ElseIf oSW.View.CurrentShowPosition = 4 Then
        If CheckVars("%Autoran%") = "False" Then
            UnsetVar "Autoran"
            Dim Autorunnable() As String
            Dim App As String
            Dim IDX As Integer
            Autorunnable = Split(GetSysConfig("Autorun"), ";")
            
            For IDX = 0 To UBound(Autorunnable) - 1 Step 1
                On Error Resume Next
                App = Autorunnable(IDX)
                With Slide1.Shapes
    
                    .AddShape(msoShapeRectangle, 0, 0, 0, 0).Name = "AutorunDummy1"
                
                    .AddShape(msoShapeRectangle, 0, 0, 0, 0).Name = "AutorunDummy2"
                
                    With .Range(Array("AutorunDummy1", "AutorunDummy2")).Group
                        .Visible = msoFalse
                        .Name = "AutorunDummy"
                    End With
                
                End With
                Application.Run "App" & App, Slide1.Shapes("AutorunDummy1")
                If ShapeExists(Slide1, "AutorunDummy") Then Slide1.Shapes("AutorunDummy").Delete
            Next IDX
        End If
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
            Slide7.Shapes("EndShowClickarea").Visible = msoFalse
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
            Slide7.Shapes("EndShowClickarea").Visible = msoTrue
            SavePresentation
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
    For I = UBound(UsersList) To 0 Step -1
        User = Replace(UsersList(I), vbNewLine, "")
        Slide1.Shapes("Username").TextFrame.TextRange.Text = User
        DeleteDir "/Users/" & User & "/"
    Next I
End Sub

Sub AddUser()
    If Slide17.UsernameFIeld.Text = "" Then
        AppMessage "Username cannot be empty", "Add user", "Error", False
        Exit Sub
    ElseIf InStr(1, Slide17.UsernameFIeld.Text, "/") Or InStr(1, Slide17.UsernameFIeld.Text, "*") Then
        AppMessage "Username contains disallowed characters", "Add user", "Error", False
        Exit Sub
    End If
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
    For I = 21 To 40 Step 1
        Slide15.Export Environ("TEMP") & "\Userpic.PNG", "PNG"
        SetFileContent "/Users/Test" & I & "/Password.txt", ""
        SetFileContent "/Users/Test" & I & "/Theme.txt", "0"
        SetFilePic "/Users/Test" & I & "/Background.png", Environ("TEMP") & "\Userpic.PNG"
    Next I
End Sub

Sub TestFixBackgrounds()
    Users = GetFiles("/Users/")
    UsersList = Split(Users, "/")
    For I = UBound(UsersList) To 0 Step -1
        User = Replace(UsersList(I), vbNewLine, "")
        Slide1.Shapes("Username").TextFrame.TextRange.Text = User
        DeleteFile "/Users/" & User & "/Background.png"
        If FileExists("/Users/" & User & "/Background.png") Then
            DeleteFile "/Users/" & User & "/Background.png"
        End If
        If FileExists("/Users/" & User & "/Background.pngBackground.png") Then
            DeleteFile "/Users/" & User & "/Background.pngBackground.png"
        End If
        CopyFile "/Defaults/Images/Background.png", "/Users/" & User & "/"
    Next I
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
    Slide2.Shapes("ShpMisc6ButtonAppSettings_").TextFrame.TextRange.Text = "Enable"
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
    For I = 0 To 3
        Dim Shp As Shape
        Set Shp = Slide1.Shapes("SwitchWorkspace")
        
        Slide1.Shapes("WorkspaceCircle" & CStr(I + 1)).Fill.Transparency = 0.5
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
            Slide1.Shapes("WorkspaceCircle2").Fill.Transparency = 0
        ElseIf Shp.TextFrame.TextRange.Text = "Workspace 2" Then
            For Each Shp In Slide1.Shapes
                If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                    If Shp.Name <> "RegularApp:-1" Then
                        Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                    End If
                End If
            Next Shp
            Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 3"
            Slide1.Shapes("WorkspaceCircle3").Fill.Transparency = 0
        ElseIf Shp.TextFrame.TextRange.Text = "Workspace 3" Then
            For Each Shp In Slide1.Shapes
                If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                    If Shp.Name <> "RegularApp:-1" Then
                        Shp.Left = Shp.Left - ActivePresentation.SlideShowWindow.Width
                    End If
                End If
            Next Shp
            Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 4"
            Slide1.Shapes("WorkspaceCircle4").Fill.Transparency = 0
        Else
            For Each Shp In Slide1.Shapes
                If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                    If Shp.Name <> "RegularApp:-1" Then
                        Shp.Left = Shp.Left + ActivePresentation.SlideShowWindow.Width * 3
                    End If
                End If
            Next Shp
            Slide1.Shapes("SwitchWorkspace").TextFrame.TextRange.Text = "Workspace 1"
            Slide1.Shapes("WorkspaceCircle1").Fill.Transparency = 0
        End If
        For Each Shp In Slide1.Shapes
            If InStr(Shp.Name, "RegularApp:") Or InStr(Shp.Name, "TaskIcon:") Then
                If Shp.Name <> "RegularApp:-1" Then
                    Shp.Delete
                End If
            End If
        Next Shp
    Next I
End Sub

Sub CopyFillFormat(Shp1 As Shape, Shp2 As Shape, Optional CopyAutoShape As Boolean = True)

    ' Remember Shp2 formatting
    Dim Fnt As Single
    Dim Fam As String
    Dim pAl As PpParagraphAlignment
    Dim mB As Single
    Dim mT As Single
    Dim mL As Single
    Dim mR As Single
    Dim trans As Single
    Dim trans2 As Single
    Dim fB As MsoTriState
    Dim fI As MsoTriState
    Dim fU As MsoTriState
    Dim fS As MsoTriState
    Dim fST As MsoTriState
    Fnt = Shp2.TextFrame.TextRange.Font.Size
    Fam = Shp2.TextFrame.TextRange.Font.Name
    pAl = Shp2.TextFrame.TextRange.ParagraphFormat.Alignment
    mB = Shp2.TextFrame.MarginBottom
    mL = Shp2.TextFrame.MarginLeft
    mT = Shp2.TextFrame.MarginTop
    mR = Shp2.TextFrame.MarginRight
    fB = Shp2.TextFrame.TextRange.Font.Bold
    fI = Shp2.TextFrame.TextRange.Font.Italic
    fU = Shp2.TextFrame.TextRange.Font.Underline
    fS = Shp2.TextFrame.TextRange.Font.Shadow
    fST = Shp2.TextFrame2.TextRange.Font.Strikethrough
    trans = Shp2.Fill.Transparency
    trans2 = Shp2.TextFrame2.TextRange.Font.Fill.Transparency
    
    ' Copy all formatting to another shape
    Shp1.PickUp
    Shp2.Apply
        
    ' Revert destination text formatting
    On Error GoTo SkipFormatting
    Shp2.Fill.Transparency = trans
    Shp2.TextFrame2.TextRange.Font.Fill.Transparency = trans2
    Shp2.TextFrame.TextRange.Font.Size = Fnt
    Shp2.TextFrame.TextRange.ParagraphFormat.Alignment = pAl
    Shp2.TextFrame.TextRange.Font.Name = Fam
    Shp2.TextFrame.MarginBottom = mB
    Shp2.TextFrame.MarginLeft = mL
    Shp2.TextFrame.MarginTop = mT
    Shp2.TextFrame.MarginRight = mR
    Shp2.TextFrame.TextRange.Font.Bold = fB
    Shp2.TextFrame.TextRange.Font.Italic = fI
    Shp2.TextFrame.TextRange.Font.Underline = fU
    Shp2.TextFrame.TextRange.Font.Shadow = fS
    Shp2.TextFrame2.TextRange.Font.Strikethrough = fST
    If CopyAutoShape Then
        Shp2.AutoShapeType = GetTheme().GroupItems("Button").AutoShapeType
    End If
SkipFormatting:
End Sub


Function HasModals(AppID As String, Optional Explicit As Boolean = False) As Boolean
    Dim SplitZ() As String
    Dim AID As String
    Dim AppNameSplit() As String
    Dim AppNameSplit2() As String
    Dim AppName As String
    Dim IsModal As Boolean
    Dim HasModal As Boolean
    Dim Shp3 As Shape
    
    HasModal = False
    IsModal = False

    AppNameSplit = Split(Slide1.Shapes("RegularApp:" & AppID).GroupItems(1).Name, ":")
    AppNameSplit2 = Split(AppNameSplit(0), "App")
    AppName = AppNameSplit2(1)
    
    If InStr(1, AppName, "Modal") Then IsModal = True
    If AppName = "Menu" Then IsModal = True
    If AppName = "InputBox" Then IsModal = True
    If AppName = "Message" Then IsModal = True
    For Each Shp3 In Slide1.Shapes
        If Shp3.Type = msoGroup Then
            If InStr(Shp3.Name, ":") And InStr(1, Shp3.Name, "ITaskIcon:") <> 1 Then
                SplitZ = Split(Shp3.Name, ":")
                AID = SplitZ(1)
                AppNameSplit = Split(Slide1.Shapes("RegularApp:" & AID).GroupItems(1).Name, ":")
                AppNameSplit2 = Split(AppNameSplit(0), "App")
                AppName = AppNameSplit2(1)
                If InStr(1, AppName, "Modal") Then HasModal = True
                If AppName = "Menu" Then HasModal = True
                If AppName = "InputBox" Then HasModal = True
                If AppName = "Message" Then HasModal = True
                If HasModal Then
                    Shp3.ZOrder msoBringToFront
                    GoTo ExitForModals
                End If
            End If
        End If
    Next Shp3
ExitForModals:
    If ((Not IsModal) And HasModal) Or (Explicit And HasModal) Then
        HasModals = True
        Slide1.Shapes("ResizeEvent").TextFrame.TextRange.Text = "N/A"
        Slide1.Shapes("MoveEvent").TextFrame.TextRange.Text = "N/A"
        Exit Function
    End If
    HasModals = False
End Function

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
        ' Avoid focusing other windows when a modal dialog is open
    End If
    Dim Shp As Shape
    Dim Shp2 As Shape
    For Each Shp In Slide1.Shapes
        If InStr(1, Shp.Name, "TaskIcon:") = 1 Then
            If InStr(1, Shp.Name, "TaskIcon:" & AppID) = 1 Then
                Shp.Fill.Transparency = 0.4
            Else
                Shp.Fill.Transparency = 0.8
            End If
        End If
        Dim ThemePath As String
        Dim FrameRef As Shape
        ThemePath = "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm"
        If Not FileStreamsExist(ThemePath) Then
            ThemePath = "/Defaults/Themes/Default.thm"
        End If
        Set FrameRef = GetFileRef(ThemePath)
        If InStr(Shp.Name, "RegularApp:") Then
            For Each Shp2 In Shp.GroupItems
                If InStr(Shp2.Name, "WindowFrameApp") Then
                    If InStr(Shp2.Name, AppID) Then
                        Shp2.Fill.Transparency = FrameRef.GroupItems("WindowFrame").Fill.Transparency
                    Else
                        Dim Tp As Double
                        If FrameRef.GroupItems("WindowFrame").Fill.Transparency = 0 Then
                            Tp = 0.4
                        Else
                            Tp = FrameRef.GroupItems("WindowFrame").Fill.Transparency * 2
                            If Tp = 2 Then
                                Tp = 1
                            ElseIf Tp > 1 Then
                                Tp = 0.9
                            End If
                        End If
                        Shp2.Fill.Transparency = Tp
                    End If
                End If
            Next Shp2
        End If
        Dim HasAx As Boolean
        Dim MultiLineTb As Boolean
        HasAx = False
        MultiLineTb = False
        If InStr(Shp.Name, "RegularApp:" & AppID) Then
            For x = 1 To Shp.GroupItems.Count
                With Shp.GroupItems(x)
                    If InStr(.Name, "AXTextBox2") Then
                        MultiLineTb = True
                    End If
                    If InStr(.Name, "AXTextBox") Then
                        ApplyTbAttribs Shp.GroupItems(x)
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
    If Not InStr(1, Shp.Name, "TaskIcon:") = 1 Then
        If InStr(1, Shp.ParentGroup.Name, "ITaskIcon:") = 1 Then
            MinimizeRestore Slide1.Shapes(Replace(Shp.ParentGroup.Name, "ITaskIcon:", "TaskIcon:"))
            Exit Sub
        End If
    End If
    Dim AppID As String
    Dim AppName As String
    AppID = GetAppID(Shp)
    If Slide1.Shapes("RegularApp:" & AppID).Visible = msoTrue Then
        If Shp.Fill.Transparency = 0.8 Then
            Slide1.Shapes("RegularApp:" & AppID).ZOrder msoBringToFront
            FocusWindow AppID
            AppName = GetAppName(Slide1.Shapes("RegularApp:" & AppID).GroupItems(1).Name)
            TryRunMacro AppName, "Focus", AppID
            UpdateTime
        Else
            LogData "Minimized window with ID " & AppID
            Slide1.Shapes("RegularApp:" & AppID).Visible = msoFalse
            Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.8
            If AAX Then
                Slide1.AxTextBox.Visible = False
            End If
            AppName = GetAppName(Slide1.Shapes("RegularApp:" & AppID).GroupItems(1).Name)
            TryRunMacro AppName, "Minimize", AppID
            ActivePresentation.SlideShowWindow.View.GotoSlide (4)
            UpdateTime
        End If
    Else
        LogData "Restoring window with ID " & AppID
        Slide1.Shapes("RegularApp:" & AppID).Visible = msoTrue
        Slide1.Shapes("RegularApp:" & AppID).ZOrder msoBringToFront
        FocusWindow AppID
        Slide1.Shapes("TaskIcon:" & AppID).Fill.Transparency = 0.4
        
        AppName = GetAppName(Slide1.Shapes("RegularApp:" & AppID).GroupItems(1).Name)
        TryRunMacro AppName, "Restore", AppID
        
        For x = 1 To Shp.GroupItems.Count
            With Slide1.Shapes("RegularApp:" & AppID).GroupItems(x)
                If InStr(.Name, "AXTextBox") Then
                    ApplyTbAttribs Shp.GroupItems(x)
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
    Dim Smp As Shape
    Set Smp = Slide1.Shapes("TaskbarButtonSample")
    If IDX = 1 Then
        MoveLeft Smp.Left
    ElseIf IDX = 2 Then
        MoveLeft Smp.Left + Smp.Width
    ElseIf IDX = 3 Then
        MoveLeft Smp.Left + 2 * Smp.Width
    ElseIf IDX = 4 Then
        MoveLeft Smp.Left + 3 * Smp.Width
    ElseIf IDX = 5 Then
        MoveLeft Smp.Left + 4 * Smp.Width
    End If
    ReorganizeITaskIcons
    LogData "Rearranged taskbar labels"
End Sub

Sub MoveLeft(Left As Integer)
    For Each Shp In Slide1.Shapes
        If InStr(1, Shp.Name, "TaskIcon:") = 1 And Shp.Left > Left And Shp.Left < ActivePresentation.SlideShowWindow.Width Then
            Shp.Left = Shp.Left - Shp.Width
        End If
    Next Shp
End Sub

Function CheckShape(ByVal Left As Integer)
    For Each Shp In Slide1.Shapes
        If InStr(1, Shp.Name, "TaskIcon:") = 1 Then
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
    For I = 0 To CInt(UBound(SplitShapes) - 1)
        CShape = SplitShapes(I)
        If Not IsInArray(CStr(CShape), ShapesX) Then
            ShapesX(I) = SplitShapes(I)
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
    For I = 0 To CInt(UBound(SplitShapes) - 1)
        CShape = SplitShapes(I)
        If Not IsInArray(CStr(CShape), ShapesX) Then
            ShapesX(I) = SplitShapes(I)
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
    x = Sld.Shapes(Replacable).Left
    y = Sld.Shapes(Replacable).Top
    W = Sld.Shapes(Replacable).Width
    h = Sld.Shapes(Replacable).Height
    T = Sld.Shapes(Replacable).TextFrame.TextRange.Text
    m = Sld.Shapes(Replacable).ActionSettings(ppMouseClick).Run
    MO = Sld.Shapes(Replacable).ActionSettings(ppMouseOver).Run
    SkinShapeLo Sld, Ref.Name, Replacable, Thm.GroupItems(control)
    With Sld.Shapes(Replacable)
        If control <> "WindowFrame" And control <> "WindowTitle" Then
            .Width = Thm.GroupItems(control).Width
            .Height = Thm.GroupItems(control).Height
        Else
            .Width = W
            .Height = h
        End If
        .Left = x
        .Top = y
        If control = "WindowTitle" Then
            .TextFrame.TextRange.Text = T
        End If
        .ActionSettings(ppMouseClick).Run = m
        .ActionSettings(ppMouseOver).Run = MO
    End With
SkipThisShape:
    Exit Sub
End Sub


' Deprecated, will break stuff
Sub ApplyTheme(ByVal Filename As String)
    'Slide2.Shapes("Icon4Part1AppMenu_").ZOrder msoSendToBack
    Dim Thm As Shape
    Dim W As Integer
    Dim h As Integer
    Dim x As Integer
    Dim y As Integer
    Set Thm = GetFileRef(Filename)
    Dim ShpArr() As Shape
    
    
    For I = Slide2.Shapes.Count To 1 Step -1
        If Slide2.Shapes(I).Type = msoGroup Then
            If InStr(1, Slide2.Shapes(I).Name, "App") = 1 Then
                If Slide2.Shapes(I).Name <> "AppMenu" Then
                    SkinShape Slide2.Shapes(I), "Close", Thm, Slide2
                    SkinShape Slide2.Shapes(I), "Minimize", Thm, Slide2
                    SkinShape Slide2.Shapes(I), "WindowTitle", Thm, Slide2
                    SkinShape Slide2.Shapes(I), "WindowFrame", Thm, Slide2
                    Slide2.Shapes(I).Visible = msoFalse
                End If
            End If
        End If
    Next I
End Sub

Sub RecoveryModeSaveAndShutdown()
    Unlight
    SavePresentation
    ActivePresentation.SlideShowWindow.View.Exit
End Sub