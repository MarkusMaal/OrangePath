' Video player app


Sub AppVideoPlayer(Shp As shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "VideoPlayer"
    
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
End Sub


Sub AssocVideoPlayer(Shp As shape)
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    Dim Video As shape
    If InStr(1, Filename, "C:\") <> 1 Then
        Set Video = GetFileRef(Filename)
    Else
        LinkMovie Filename
        Set Video = GetFileRef("/Temp/Movie.mov")
    End If
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "VideoPlayer"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Dim NewAppID As String
    NewAppID = Slide1.Shapes("AppID").TextFrame.TextRange.Text
    Slide1.Shapes("WindowAppVideoPlayer:" & NewAppID).TextFrame.TextRange.Text = "Please wait, loading..."
    Slide1.Shapes("Shape5AppVideoPlayer:" & NewAppID).TextFrame.TextRange.Text = Filename
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
End Sub

Sub AssocIVideoPlayer(Shp As shape)
    On Error GoTo ReportIssue2
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocVideoPlayer Slide1.Shapes(ShapeName)
    Exit Sub
ReportIssue2:
    MsgBox Err.Description
End Sub