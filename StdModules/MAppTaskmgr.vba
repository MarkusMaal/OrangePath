' Task manager app

Sub AppTaskmgr(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Taskmgr"
    
    Dim X As Long
    Dim oShp As Shape
    Dim taskList As String
    taskList = ""
    For Each oShp In Slide1.Shapes
        If InStr(oShp.Name, "RegularApp:") Then
            SplitZ = Split(oShp.Name, ":")
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
End Sub

