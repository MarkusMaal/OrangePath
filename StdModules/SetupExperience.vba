Sub HideHourglass()
    Slide26.Shapes("Hourglass").Visible = msoFalse
    Slide26.Shapes("WaitText").Visible = msoFalse
End Sub

Sub ShowHourglass()
    Slide26.Shapes("Hourglass").Visible = msoTrue
    Slide26.Shapes("WaitText").Visible = msoTrue
End Sub

Sub SetupRadioUncheck(ByVal RadioGroup As String, ByVal Parent As Shape)
    Dim Shp As Shape
    For Each Shp In Parent.GroupItems
        Shp.Line.ForeColor.RGB = RGB(255, 255, 255)
    Next Shp
End Sub

Sub SetupRadioCheck(Shp As Shape)
    Dim RadioGroup As String
    Dim SplitName() As String
    SplitName = Split(Shp.Name, "/")
    RadioGroup = SplitName(1)
    SetupRadioUncheck RadioGroup, Shp.ParentGroup
    Shp.Line.ForeColor.RGB = RGB(0, 0, 0)
End Sub

Function IsRadioChecked(ByVal RadioGroup As String, ByVal Label As String)
    If Slide26.Shapes(Label + "/" + RadioGroup).Line.ForeColor.RGB = RGB(0, 0, 0) Then
        IsRadioChecked = msoTrue
    Else
        IsRadioChecked = msoFalse
    End If
End Function

Sub TestRadioMacro()
    SetupRadioCheck Slide26.Shapes("Upgrade/1")
End Sub

Sub SetupStep2()
    If Slide26.Shapes("SetupScreenWelcome").Visible = msoTrue Then
        If IsRadioChecked("1", "Migrate") Then
            AppMessage "Not Implemented", "Setup experience", "Error", False
        ElseIf IsRadioChecked("1", "Upgrade") Then
            AppMessage "Not Implemented", "Setup experience", "Error", False
        ElseIf IsRadioChecked("1", "FreshInstall") Then
            Slide26.Shapes("SetupScreenWelcome").Visible = msoFalse
            Slide26.Shapes("SetupScreenPartition").Visible = msoTrue
            Slide26.Shapes("Progress1").Visible = msoTrue
        End If
    End If
End Sub
