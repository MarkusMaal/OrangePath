' Firmware upgrader code

Sub UpdateNow()
    UpdateTest
    Restart
End Sub

Sub UpdateSystem()
    If Slide12.Shapes("FirmwareSource").TextFrame.TextRange.Text <> "null" Then
        Slide12.Shapes("StatusText").TextFrame.TextRange.Text = "Updating system..."
        Slide12.Shapes("Notice").TextFrame.TextRange.Text = "Do not close the presentation!"
        ActivePresentation.SlideShowWindow.Activate
        Dim Presentation2 As Presentation
        Set Presentation2 = Presentations.Open(Filename:=Slide12.Shapes("FirmwareSource").TextFrame.TextRange.Text, ReadOnly:=msoTrue, WithWindow:=msoFalse)
        Presentation2.Application.Run "PatchParent"
        Presentation2.Close
        ActivePresentation.SlideShowWindow.Activate
    Else
        Slide12.Shapes("StatusText").TextFrame.TextRange.Text = "Missing update file"
        Slide12.Shapes("Notice").TextFrame.TextRange.Text = "In order to update, an update file must be selected first!"
        Pause (5)
        ActivePresentation.SlideShowWindow.Activate
    End If
End Sub

Sub UpdateTest()
    ' Create a backup copy
    With Application.ActivePresentation
        .SaveCopyAs "Sunlight.bak.ppsm"
    End With
    On Error Resume Next
    ' Open second presentation
    Dim sourcePresentation As Presentation
    Set sourcePresentation = Presentations.Open(Filename:=Slide12.Shapes("FirmwareSource").TextFrame.TextRange.Text, ReadOnly:=msoTrue, WithWindow:=msoFalse)
    Dim Shp As Shape
    ' Copy/Paste windows
    For Each Shp In Presentations(2).Slides(1).Shapes
        Shp.Copy
        Slide2.Shapes.Paste
        CopyModuleToAnotherPresentation sourcePresentation, "M" & Shp.Name
    Next Shp
    ' Copy/Paste shortcut icons
    For Each Shp In Presentations(2).Slides(2).Shapes
        Shp.Copy
        Slide25.Shapes.Paste
    Next Shp
    ' Close the second presentation
    sourcePresentation.Close
    ' Go back to the slide show window
    ActivePresentation.SlideShowWindow.Activate
End Sub

' thx ChatGPT
' modified that code a bit tough
Sub CopyModuleToAnotherPresentation(sourcePresentation As Presentation, moduleName As String)
    Dim targetPresentation As Presentation

    ' Open the source and target presentations
    Set targetPresentation = ActivePresentation

    ' Loop through the VBComponents in the source presentation
    For Each sourceVBComponent In sourcePresentation.VBProject.VBComponents
        ' Check if the current VBComponent is a module and has the specified name
        If sourceVBComponent.Name = moduleName Then
            ' Add a new module with the same code to the target presentation
            Set targetVBComponent = targetPresentation.VBProject.VBComponents.Add(sourceVBComponent.Type)
            targetVBComponent.CodeModule.AddFromString sourceVBComponent.CodeModule.Lines(1, sourceVBComponent.CodeModule.CountOfLines)
            ' Rename module to what it was in the source presentation
            ActivePresentation.VBProject.VBComponents("Module1").Name = moduleName
            Exit For ' Exit the loop once the module is found and copied
        End If
    Next sourceVBComponent

    ' Clean up
    Set targetPresentation = Nothing
End Sub

Function GetStdModuleType() As Integer
    For Each vbcomponent In ActivePresentation.VBProject.VBComponents
        GetStdModuleType = vbcomponent.Type
        Exit Function
    Next vbcomponent
End Function

Sub ClosePresentation2()
    Presentations(2).Close
End Sub

