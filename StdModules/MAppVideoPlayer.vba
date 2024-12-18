' Video player app


Sub AppVideoPlayer(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "VideoPlayer"
    
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide1.Shapes("WindowTitleAppVideoPlayer:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "Video player"
    UpdateTime
End Sub


Sub AssocVideoPlayer(Shp As Shape)
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    Dim Video As Shape
    If InStr(1, Filename, "C:\") <> 1 Then
        Set Video = GetFileRef(Filename)
    Else
        LinkMovie Filename
        Set Video = GetFileRef("/Temp/Movie.mov")
    End If
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "VideoPlayer"
    Slide2.Shapes("WindowAppVideoPlayer_").TextFrame.TextRange.Text = "Please wait, loading..."
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide2.Shapes("WindowAppVideoPlayer_").TextFrame.TextRange.Text = "No video file loaded!"
    Dim NewAppID As String
    NewAppID = Slide1.Shapes("AppID").TextFrame.TextRange.Text
    Slide1.Shapes("WindowTitleAppVideoPlayer:" & NewAppID).TextFrame.TextRange.Text = Filename
    Slide1.Shapes("TaskIcon:" & NewAppID).TextFrame.TextRange.Text = Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    PasteToGroup Shp, Video, "VideoAppVideoPlayer:" & NewAppID, Slide1.Shapes("WindowAppVideoPlayer:" & NewAppID).Left, Slide1.Shapes("WindowAppVideoPlayer:" & NewAppID).Top, Slide1
    With Slide1.Shapes("VideoAppVideoPlayer:" & NewAppID)
        .Width = Slide1.Shapes("WindowAppVideoPlayer:" & NewAppID).Width
        .Height = Slide1.Shapes("WindowAppVideoPlayer:" & NewAppID).Height
        .Visible = msoTrue
    End With
    Regroup NewAppID, Slide1
    Slide1.Shapes("WindowAppVideoPlayer:" & NewAppID).Fill.Transparency = 1
    Slide1.Shapes("WindowAppVideoPlayer:" & NewAppID).TextFrame.TextRange.Text = ""
    UpdateTime
End Sub

Sub AssocIVideoPlayer(Shp As Shape)
    On Error GoTo ReportIssue2
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocVideoPlayer Slide1.Shapes(ShapeName)
    Exit Sub
ReportIssue2:
    MsgBox Err.Description
End Sub
