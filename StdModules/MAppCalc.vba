' Calculator app

Sub AppCalc(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Calc"
    Slide2.Shapes("AppCalc").Visible = msoTrue
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide1.Shapes("WindowTitleAppCalc:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "Calculator"
    Slide2.Shapes("AppCalc").Visible = msoFalse
    UpdateTime
End Sub

Sub CalcNumber(Shp As Shape)
    AppID = GetAppID(Shp.ParentGroup)
    CNum = Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text
    If Shp.TextFrame.TextRange.Text = "," And InStr(CNum, ",") Then Exit Sub
    If CNum = "0" Then
        Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text = Shp.TextFrame.TextRange.Text
    Else
        Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text = CNum & Shp.TextFrame.TextRange.Text
    End If
    
End Sub

Sub CalcClear(Shp As Shape)
    AppID = GetAppID(Shp.ParentGroup)
    Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text = "0"
End Sub

Sub CalcAdd(Shp As Shape)
    On Error GoTo Crash
    AppID = GetAppID(Shp.ParentGroup)
    Memory = Slide1.Shapes("Shape7AppCalc:" & AppID).TextFrame.TextRange.Text
    CurrentValue = CDbl(Memory) + CDbl(Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text)
    Slide1.Shapes("Shape7AppCalc:" & AppID).TextFrame.TextRange.Text = CStr(CurrentValue)
    Slide1.Shapes("Shape8AppCalc:" & AppID).TextFrame.TextRange.Text = "+"
    Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text = "0"
Done:
    Exit Sub
Crash:
    OSCrash "CALCULATION_ERROR", Err
End Sub

Sub CalcSub(Shp As Shape)
    On Error GoTo Crash
    AppID = GetAppID(Shp.ParentGroup)
    Memory = Slide1.Shapes("Shape7AppCalc:" & AppID).TextFrame.TextRange.Text
    If Memory <> 0 Then
        CurrentValue = CDbl(Memory) - CDbl(Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text)
    Else
        CurrentValue = CDbl(Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text)
    End If
    Slide1.Shapes("Shape7AppCalc:" & AppID).TextFrame.TextRange.Text = CStr(CurrentValue)
    Slide1.Shapes("Shape8AppCalc:" & AppID).TextFrame.TextRange.Text = "-"
    Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text = "0"
Done:
    Exit Sub
Crash:
    OSCrash "CALCULATION_ERROR", Err
End Sub

Sub CalcDiv(Shp As Shape)
    On Error GoTo Crash
    AppID = GetAppID(Shp.ParentGroup)
    Memory = Slide1.Shapes("Shape7AppCalc:" & AppID).TextFrame.TextRange.Text
    If Memory <> 0 Then
        CurrentValue = CDbl(Memory) / CDbl(Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text)
    Else
        CurrentValue = CDbl(Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text)
    End If
    Slide1.Shapes("Shape7AppCalc:" & AppID).TextFrame.TextRange.Text = CStr(CurrentValue)
    Slide1.Shapes("Shape8AppCalc:" & AppID).TextFrame.TextRange.Text = "/"
    Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text = "0"
Done:
    Exit Sub
Crash:
    OSCrash "CALCULATION_ERROR", Err
End Sub

Sub CalcMul(Shp As Shape)
    On Error GoTo Crash
    AppID = GetAppID(Shp.ParentGroup)
    Memory = Slide1.Shapes("Shape7AppCalc:" & AppID).TextFrame.TextRange.Text
    If Memory <> 0 Then
        CurrentValue = CDbl(Memory) * CDbl(Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text)
    Else
        CurrentValue = CDbl(Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text)
    End If
    Slide1.Shapes("Shape7AppCalc:" & AppID).TextFrame.TextRange.Text = CStr(CurrentValue)
    Slide1.Shapes("Shape8AppCalc:" & AppID).TextFrame.TextRange.Text = "*"
    Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text = "0"
Done:
    Exit Sub
Crash:
    OSCrash "CALCULATION_ERROR", Err
End Sub

Sub CalcEqu(Shp As Shape)
    On Error GoTo Crash
    AppID = GetAppID(Shp.ParentGroup)
    Memory = CDbl(Slide1.Shapes("Shape7AppCalc:" & AppID).TextFrame.TextRange.Text)
    CurrentValue = CDbl(Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text)
    Operation = Slide1.Shapes("Shape8AppCalc:" & AppID).TextFrame.TextRange.Text
    If Operation = "+" Then
        CurrentValue = CurrentValue + Memory
    ElseIf Operation = "-" Then
        CurrentValue = Memory - CurrentValue
    ElseIf Operation = "/" Then
        CurrentValue = Memory / CurrentValue
    ElseIf Operation = "*" Then
        CurrentValue = CurrentValue * Memory
    End If
    Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text = CurrentValue
    Slide1.Shapes("Shape7AppCalc:" & AppID).TextFrame.TextRange.Text = "0"
Done:
    Exit Sub
Crash:
    OSCrash "CALCULATION_ERROR", Err
End Sub

Sub CalcReset(Shp As Shape)
    AppID = GetAppID(Shp.ParentGroup)
    Memory = Slide1.Shapes("Shape7AppCalc:" & AppID).TextFrame.TextRange.Text
    Slide1.Shapes("Shape7AppCalc:" & AppID).TextFrame.TextRange.Text = "0"
    Slide1.Shapes("WindowAppCalc:" & AppID).TextFrame.TextRange.Text = "0"
    Operation = Slide1.Shapes("Shape8AppCalc:" & AppID).TextFrame.TextRange.Text = "+"
End Sub
