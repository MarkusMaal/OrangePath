' Test app (Generated from devCreateApp)

' This is executed when the application is launched
Sub AppTest(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Test"
    Slide2.Shapes("AppTest").Visible = msoTrue
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide2.Shapes("AppTest").Visible = msoFalse
End Sub

' This gets executed when a user clicks a file, which is associated with this application
Sub AssocTest(Shp As Shape)
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Test"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
End Sub

' This gets executed when a user clicks icon of a file, which is associated with this application
Sub AssocITest(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocTest Slide1.Shapes(ShapeName)
End Sub

Sub AppTestConfirm(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim InputValue As String
    InputValue = Slide1.Shapes("AXTextBox1AppTest:" & AppID).TextFrame.TextRange.Text
    Slide1.Shapes("NamePromptAppTest:" & AppID).TextFrame.TextRange.Text = "Hello, " & InputValue & "!"
    Slide1.Shapes("AXTextBox1AppTest:" & AppID).Delete
    Slide1.Shapes("ConfirmButtonAppTest:" & AppID).Delete
    Slide1.AxTextBox.Visible = False
End Sub

