' Picture view app

Sub AppPictureView(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "PictureView"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
End Sub


Sub AssocPictureView(Shp As Shape)
    On Error GoTo ReportIssue2
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    
    Dim Pic As Shape
    
    If InStr(1, Filename, "C:\") = 1 Then
        SetFilePic "/Temp/LocalDisk.pic", Filename
        Set Pic = GetFileRef("/Temp/LocalDisk.pic")
    Else
        Set Pic = GetFileRef(Filename)
    End If
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "PictureView"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Dim NewAppID As String
    NewAppID = Slide1.Shapes("AppID").TextFrame.TextRange.Text
    Slide1.Shapes("WindowAppPictureView:" & NewAppID).TextFrame.TextRange.Text = "Please wait, loading..."
    Slide1.Shapes("WindowTitleAppPictureView:" & NewAppID).TextFrame.TextRange.Text = Filename
    Slide1.Shapes("TaskIcon:" & NewAppID).TextFrame.TextRange.Text = Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    PasteToGroup Shp, Pic, "ImageAppPictureView:" & NewAppID, Slide1.Shapes("WindowAppPictureView:" & NewAppID).Left, Slide1.Shapes("WindowAppPictureView:" & NewAppID).Top, Slide1
    Slide1.Shapes("ImageAppPictureView:" & NewAppID).Height = Slide1.Shapes("WindowAppPictureView:" & NewAppID).Height
    Slide1.Shapes("ImageAppPictureView:" & NewAppID).Width = Slide1.Shapes("WindowAppPictureView:" & NewAppID).Width
    Slide1.Shapes("ImageAppPictureView:" & NewAppID).Visible = msoTrue
    Regroup NewAppID, Slide1
    Slide1.Shapes("WindowAppPictureView:" & NewAppID).TextFrame.TextRange.Text = ""
    Slide1.Shapes("WindowAppPictureView:" & NewAppID).Fill.Transparency = 1
    Exit Sub
ReportIssue2:
    Slide1.Shapes("RegularApp:" & NewAppID).Delete
    Slide1.Shapes("TaskIcon:" & NewAppID).Delete
    AppMessage Err.Description, "Error loading image", "Error", True
    Regroup AppID, Slide1
End Sub

Sub AssocIPictureView(Shp As Shape)
    On Error GoTo ReportIssue2
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocPictureView Slide1.Shapes(ShapeName)
    Exit Sub
ReportIssue2:
    AppMessage Err.Description, "Error loading image", "Error", True
End Sub


