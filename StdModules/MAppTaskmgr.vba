' Task manager app

Sub AppTaskmgr(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Taskmgr"
    
    Dim x As Long
    Dim oshp As Shape
    Dim taskList As String
    taskList = ""
    For Each oshp In Slide1.Shapes
        If InStr(oshp.Name, "RegularApp:") Then
            SplitZ = Split(oshp.Name, ":")
            AppID = SplitZ(1)
            With Slide1.Shapes("RegularApp:" & AppID)
                If .GroupItems.Count > 0 Then
                    AppNameSplit = Split(.GroupItems(1).Name, ":")
                    AppNameSplit2 = Split(AppNameSplit(0), "App")
                    AppName = AppNameSplit2(1)
                    taskList = taskList & AppName & " (PID: " & CStr(AppID) & ")" & vbNewLine
                End If
            End With
        End If
    Next
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide1.Shapes("WindowTitleAppTaskmgr:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "Task list"
    UpdateTime
End Sub
