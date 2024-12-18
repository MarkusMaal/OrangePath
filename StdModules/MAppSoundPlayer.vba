' SoundPlayer app (Generated from devCreateApp)
 
' This is executed when the application is launched
Sub AppSoundPlayer(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "SoundPlayer"
    Slide2.Shapes("AppSoundPlayer").Visible = msoTrue
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    
    Slide1.Shapes("WindowTitleAppSoundPlayer:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "Sound player"
    Slide2.Shapes("AppSoundPlayer").Visible = msoFalse
    UpdateTime
End Sub

Sub AssocSoundPlayer(Shp As Shape)
    On Error GoTo ReportIssue2
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    Dim Sound As Shape
    If InStr(1, Filename, "C:\") <> 1 Then
        Set Sound = GetFileRef(Filename)
    Else
        LinkMovie Filename
        Set Sound = GetFileRef("/Temp/Movie.mov")
    End If
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "SoundPlayer"
    Slide2.Shapes("WindowAppSoundPlayer_").TextFrame.TextRange.Text = "Please wait, loading..."
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide2.Shapes("WindowAppSoundPlayer_").TextFrame.TextRange.Text = "No file loaded!"
    Dim NewAppID As String
    NewAppID = Slide1.Shapes("AppID").TextFrame.TextRange.Text
    Slide1.Shapes("WindowTitleAppSoundPlayer:" & NewAppID).TextFrame.TextRange.Text = Filename
    Slide1.Shapes("TaskIcon:" & NewAppID).TextFrame.TextRange.Text = Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    PasteToGroup Shp, Sound, "SoundAppSoundPlayer:" & NewAppID, Slide1.Shapes("WindowAppSoundPlayer:" & NewAppID).Left, Slide1.Shapes("WindowAppSoundPlayer:" & NewAppID).Top, Slide1
    With Slide1.Shapes("SoundAppSoundPlayer:" & NewAppID)
        .Fill.Transparency = 1
        .Line.Transparency = 1
        .PictureFormat.Brightness = 0
        .Visible = msoTrue
        .LockAspectRatio = msoFalse
        .Width = Slide1.Shapes("WindowAppSoundPlayer:" & NewAppID).Width
        .Height = Slide1.Shapes("WindowAppSoundPlayer:" & NewAppID).Height
    End With
    Regroup NewAppID, Slide1
    Slide1.Shapes("WindowAppSoundPlayer:" & NewAppID).Fill.ForeColor.RGB = RGB(0, 0, 0)
    
    Slide1.Shapes("WindowAppSoundPlayer:" & NewAppID).TextFrame.TextRange.Text = ""
    Exit Sub
ReportIssue2:
    Slide1.Shapes("RegularApp:" & NewAppID).Delete
    Slide1.Shapes("TaskIcon:" & NewAppID).Delete
    AppMessage Err.Description, "Error loading audio", "Error", True
    Regroup AppID, Slide1
    UpdateTime
End Sub

Sub AssocISoundPlayer(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocSoundPlayer Slide1.Shapes(ShapeName)
End Sub
