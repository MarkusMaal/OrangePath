' Components

Sub AppComponents(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Components"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide1.Shapes("WindowTitleAppComponents:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "UI components"
    UpdateTime
End Sub


Sub AssocComponents(Shp As Shape)
    ' Get full file path from shape
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Components"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide1.Shapes("WindowTitleAppComponents:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "UI components"
    AppComponentsSetLine Slide1.Shapes("AppID").TextFrame.TextRange.Text, 1, Filename
    AppComponentsDisplayShape Slide1.Shapes("RegularApp:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text)
    UpdateTime
    CloseWindow Slide1.Shapes("CloseAppComponents:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text)
    Slide1.Shapes("TaskIcon:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).Delete
    Slide1.Shapes("ITaskIcon:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).Delete
End Sub
Sub AssocIComponents(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocComponents Slide1.Shapes(ShapeName)
End Sub

Sub AppComponentsModalFiles(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    SetVar "Macro", "AppComponentsLoadFile"
    SetVar "AppID", AppID
    UnsetVar "Save"
    If Shp.TextFrame.TextRange.Text = "Save file" Then
        SetVar "Macro", "AppComponentsSaveFile"
        SetVar "Save", "True"
        
    End If
    AppModalFiles
End Sub

Sub AppComponentsSaveFile()
    AppMessage "Selected output: " & CheckVars("%InputValue%"), "Save test", "Info", True
End Sub

Sub AppComponentsSetLine(ByVal AppID As String, ByVal ID As Integer, ByVal Val As String)
    Dim CVal As String
    CVal = Slide1.Shapes("ValuesAppComponents:" & AppID).TextFrame.TextRange.Text
    Dim NVal As String
    NVal = ""
    Dim IDX As Integer
    IDX = 1
    Dim Line As Variant
    For Each Line In Split(Replace(vbCrLf, vbCr, CVal), vbCr)
        If IDX = ID Then
            NVal = NVal & Val
        Else
            NVal = NVal & Line
        End If
        If IDX <> 3 Then
            NVal = NVal & vbNewLine
        End If
        IDX = IDX + 1
    Next Line
    Slide1.Shapes("ValuesAppComponents:" & AppID).TextFrame.TextRange.Text = NVal
End Sub

Function AppComponentsGetLine(ByVal AppID As String, ByVal ID As Integer) As String
    Dim CVal As String
    Dim Line As Variant
    Dim IDX As Integer
    CVal = Slide1.Shapes("ValuesAppComponents:" & AppID).TextFrame.TextRange.Text
    IDX = 1
    For Each Line In Split(Replace(vbCrLf, vbCr, CVal), vbCr)
        If IDX = ID Then
            AppComponentsGetLine = Line
        End If
        IDX = IDX + 1
    Next Line
End Function

Sub AppComponentsLoadFile()
    Dim AppID As String
    AppID = CheckVars("%AppID%")
    AppComponentsSetLine AppID, 1, CheckVars("%InputValue%")
End Sub

Sub AppComponentsDispMessage(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    AppMessage AppComponentsGetLine(AppID, 3), AppComponentsGetLine(AppID, 2), Shp.TextFrame.TextRange.Text, True
End Sub

Sub Test1()
    AppComponentsDisplayShape Slide1.Shapes("DisplayShapeButtonAppComponents:5")
End Sub

Sub AppComponentsDispInputBox(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    SetVar "Macro", "AppComponentsConfirmInput"
    SetVar "AppID", AppID
    SetVar "MsgID", "2"
    If Shp.TextFrame.TextRange.Text = "Text" Then
        SetVar "MsgID", "3"
    End If
    AppInputBox "Enter value for '" & Shp.TextFrame.TextRange.Text & "'", "Login screen"
    
End Sub

Sub AppComponentsConfirmInput()
    Dim AppID As String
    Dim InputVal As String
    Dim MsgID As Integer
    AppID = CheckVars("%AppID%")
    InputVal = CheckVars("%InputValue%")
    MsgID = CInt(CheckVars("%MsgID%"))
    AppComponentsSetLine AppID, MsgID, InputVal
    
End Sub

Sub AppComponentsDisplayShape(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    GetFileRef(AppComponentsGetLine(AppID, 1)).Copy
    'On Error Resume Next
    With Slide27.Shapes.Paste
        .Name = "ExtractedShape"
        .Visible = msoTrue
        .Top = 0
        .Left = 0
        .Width = ActivePresentation.PageSetup.SlideWidth
        .Height = ActivePresentation.PageSetup.SlideHeight
        On Error GoTo SingleShape
        Dim Shp2 As Shape
        For Each Shp2 In .GroupItems
            Shp2.ActionSettings(ppMouseClick).Run = "AppComponentsCleanDisplay"
        Next Shp2
        GoTo ExitWith
SingleShape:
        .ActionSettings(ppMouseClick).Run = "AppComponentsCleanDisplay"
ExitWith:
    End With
    Slide27.Shapes("SlideShowWindow").ActionSettings(ppMouseClick).Run = "AppComponentsCleanDisplay"
    ActivePresentation.SlideShowWindow.View.GotoSlide 28
End Sub

Sub AppComponentsCleanDisplay()
    Slide27.Shapes("ExtractedShape").Delete
    Slide27.Shapes("SlideShowWindow").ActionSettings(ppMouseClick).Run = "AdvanceShow"
    ActivePresentation.SlideShowWindow.View.GotoSlide 4
End Sub