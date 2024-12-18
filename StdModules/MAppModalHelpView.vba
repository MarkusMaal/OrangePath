' ModalHelpView app (Generated from devCreateApp)

' This is executed when the application is launched
Sub AppModalHelpView(Optional HelpFile As String = "")
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "ModalHelpView"
    Slide2.Shapes("AppModalHelpView").Visible = msoTrue
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    Slide2.Shapes("AppModalHelpView").Visible = msoFalse
    Dim NewAppID As String
    NewAppID = Slide1.Shapes("AppID").TextFrame.TextRange.Text
    If HelpFile <> "" Then
        Dim Ref As Shape
        Set Ref = GetFileRef(HelpFile)
        Ref.Copy
        With Slide1.Shapes.Paste
            Dim Shp As Shape
            .Left = Slide1.Shapes("WindowAppModalHelpView:" & NewAppID).Left
            .Top = Slide1.Shapes("WindowAppModalHelpView:" & NewAppID).Top
            .Width = Slide1.Shapes("WindowAppModalHelpView:" & NewAppID).Width
           .Height = Slide1.Shapes("WindowAppModalHelpView:" & NewAppID).Height
           .Visible = msoTrue
           Dim I As Integer
            I = 1
            For Each Shp In .GroupItems
                Shp.Name = "Content" & CStr(I) & "AppModalHelpView:" & NewAppID
                I = I + 1
            Next Shp
            .Ungroup
        End With
        Slide1.Shapes("RegularApp:" & NewAppID).Ungroup
        Regroup NewAppID, Slide1
    Else
        Slide1.Shapes("WindowAppModalHelpView:" & NewAppID).TextFrame.TextRange.Text = "No file loaded"
    End If
    Slide1.Shapes("WindowTitleAppModalHelpView:" & NewAppID).TextFrame.TextRange.Text = "Help viewer"
End Sub

' This gets executed when a user clicks a file, which is associated with this application
Sub AssocModalHelpView(Shp As Shape)
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    AppModalHelpView Filename
End Sub

' This gets executed when a user clicks icon of a file, which is associated with this application
Sub AssocIModalHelpView(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocModalHelpView Slide1.Shapes(ShapeName)
End Sub
