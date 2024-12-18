' Picture view app

Sub AppPictureView(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "PictureView"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide1.Shapes("WindowTitleAppPictureView:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "Picture viewer"
    UpdateTime
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
    Slide2.Shapes("WindowAppPictureView_").TextFrame.TextRange.Text = "Generating preview..."
    CreateNewWindow
    Slide2.Shapes("WindowAppPictureView_").TextFrame.TextRange.Text = "No image file loaded"
    Dim NewAppID As String
    NewAppID = Slide1.Shapes("AppID").TextFrame.TextRange.Text
    Slide1.Shapes("WindowAppPictureView:" & NewAppID).TextFrame.TextRange.Text = "Generating preview..."
    Slide1.Shapes("WindowTitleAppPictureView:" & NewAppID).TextFrame.TextRange.Text = Filename
    Slide1.Shapes("TaskIcon:" & NewAppID).TextFrame.TextRange.Text = Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    ActivePresentation.SlideShowWindow.View.GotoSlide ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    ActivePresentation.SlideShowWindow.Activate
    PasteToGroup Shp, Pic, "ImageAppPictureView:" & NewAppID, Slide1.Shapes("WindowAppPictureView:" & NewAppID).Left, Slide1.Shapes("WindowAppPictureView:" & NewAppID).Top, Slide1
    With Slide1.Shapes("ImageAppPictureView:" & NewAppID)
        .Width = Slide1.Shapes("WindowAppPictureView:" & NewAppID).Width
        .Height = Slide1.Shapes("WindowAppPictureView:" & NewAppID).Height
        .Visible = msoTrue
        .ActionSettings(ppMouseClick).Action = ppActionRunMacro
        .ActionSettings(ppMouseClick).Run = "AppPictureViewFullscreen"
    End With
    Regroup NewAppID, Slide1
    Slide1.Shapes("WindowAppPictureView:" & NewAppID).TextFrame.TextRange.Text = ""
    Slide1.Shapes("WindowAppPictureView:" & NewAppID).Fill.Transparency = 1
    Exit Sub
ReportIssue2:
    Slide1.Shapes("RegularApp:" & NewAppID).Delete
    Slide1.Shapes("TaskIcon:" & NewAppID).Delete
    Slide1.Shapes("ITaskIcon:" & NewAppID).Delete
    AppMessage Err.Description, "Error loading image", "Error", True
    Regroup AppID, Slide1
    UpdateTime
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

Sub AppPictureViewFullscreen(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Shp.Copy
    With Slide27.Shapes.Paste
        .Left = 0
        .Top = 0
        .Width = Slide27.Shapes("SlideShowWindow").Width
        .Height = Slide27.Shapes("SlideShowWindow").Height
        .Name = "FullImage"
        .ActionSettings(ppMouseClick).Run = ""
    End With
    ActivePresentation.SlideShowWindow.View.GotoSlide 28
    AppGalleryShowControls
End Sub

' Deprecated
Sub AppPictureViewExitFullScreen(Shp As Shape)
    Shp.Delete
    ActivePresentation.SlideShowWindow.View.GotoSlide 4
End Sub