' 3D rotate app

Sub App3D(Shp As Shape)
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "3D"
    Slide2.Shapes("Shape11AppGuess_").TextFrame.TextRange.Text = CStr(Int(100 * Rnd))
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide1.Shapes("WindowTitleApp3D:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "3D Rotate"
    UpdateTime
End Sub

Sub RotXAdd(Shp As Shape)
    AppID = GetAppID(Shp)
    Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationX = Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationX + 10
End Sub

Sub RotXSub(Shp As Shape)
    AppID = GetAppID(Shp)
    Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationX = Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationX - 10
End Sub


Sub Assoc3D(Shp As Shape)
    ' Get full file path
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    ' Launch 3D rotate
    Dim Shape3D As Shape
    Set Shape3D = GetFileRef(Filename)
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "3D"
    Slide2.Shapes("Shape11AppGuess_").TextFrame.TextRange.Text = CStr(Int(100 * Rnd))
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    ' Get AppID of the new window
    AppID = Slide1.Shapes("AppID").TextFrame.TextRange.Text
    ' Load shape rotation
    Slide1.Shapes("WindowTitleApp3D:" & AppID).TextFrame.TextRange.Text = "Loading..."
    Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationZ = Shape3D.ThreeD.RotationZ
    Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationY = Shape3D.ThreeD.RotationY
    Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationX = Shape3D.ThreeD.RotationX
    ' Display filename on the task bar
    If Len(Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text) > 13 Then
        Slide1.Shapes("TaskIcon:" & AppID).TextFrame.TextRange.Text = "..." & Right(Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text, 13)
    Else
        Slide1.Shapes("TaskIcon:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    End If
    ' Display file path on the window title bar
    Slide1.Shapes("WindowTitleApp3D:" & AppID).TextFrame.TextRange.Text = Filename
    UpdateTime
End Sub

Sub AssocI3D(Shp As Shape)
    ' This function is used to find the label from an icon and then execute Assoc3D with the label as the parameter
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    Assoc3D Slide1.Shapes(ShapeName)
End Sub

Sub RotYAdd(Shp As Shape)
    AppID = GetAppID(Shp)
    Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationY = Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationY + 10
End Sub

Sub RotYSub(Shp As Shape)
    AppID = GetAppID(Shp)
    Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationY = Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationY - 10
End Sub


Sub RotZAdd(Shp As Shape)
    AppID = GetAppID(Shp)
    Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationZ = Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationZ + 10
End Sub

Sub RotZSub(Shp As Shape)
    AppID = GetAppID(Shp)
    Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationZ = Slide1.Shapes("Shape2App3D:" & AppID).ThreeD.RotationZ - 10
End Sub


