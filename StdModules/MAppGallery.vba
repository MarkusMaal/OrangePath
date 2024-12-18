' Gallery app (Generated from devCreateApp)

' This is executed when the application is launched
Sub AppGallery(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Gallery"
    Slide2.Shapes("AppGallery").Visible = msoTrue
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    AppGalleryFindPics Slide1.Shapes("AppID").TextFrame.TextRange.Text
    Slide1.Shapes("WindowTitleAppGallery:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "Gallery - Page 1"
    Slide2.Shapes("AppGallery").Visible = msoFalse
    Slide1.Shapes("ButtonBackArrowAppGallery:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).Fill.Transparency = 1
    Slide1.Shapes("ButtonBackArrowAppGallery:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame2.TextRange.Font.Fill.Transparency = 1
    Slide1.Shapes("ButtonBackArrowAppGallery:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).ActionSettings(ppMouseClick).Run = ""
End Sub

' This gets executed when a user clicks a file, which is associated with this application
Sub AssocGallery(Shp As Shape)
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Gallery"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
End Sub

' This gets executed when a user clicks icon of a file, which is associated with this application
Sub AssocIGallery(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocGallery Slide1.Shapes(ShapeName)
End Sub

Sub ClearPics(AppID As String)
    For I = 1 To 9
        Slide1.Shapes("Pic" & I & "AppGallery:" & AppID).ActionSettings(ppMouseClick).Run = ""
        Slide1.Shapes("Pic" & I & "AppGallery:" & AppID).Fill.Solid
        Slide1.Shapes("Pic" & I & "AppGallery:" & AppID).Fill.ForeColor = Slide2.Shapes("Pic" & I & "AppGallery_").Fill.ForeColor
    Next I
End Sub

Sub AppGalleryNextPage(Shp As Shape)
    Dim AppID As String
    Dim Page As Integer
    AppID = GetAppID(Shp)
    Page = CInt(Replace(Slide1.Shapes("WindowTitleAppGallery:" & AppID).TextFrame.TextRange.Text, "Gallery - Page ", ""))
    ClearPics AppID
    Page = Page + 1
    Slide1.Shapes("ButtonBackArrowAppGallery:" & AppID).Fill.Transparency = 0
    Slide1.Shapes("ButtonBackArrowAppGallery:" & AppID).TextFrame2.TextRange.Font.Fill.Transparency = 0
    Slide1.Shapes("ButtonBackArrowAppGallery:" & AppID).ActionSettings(ppMouseClick).Run = "AppGalleryPrevPage"
    AppGalleryFindPics AppID, CStr((Page - 1) * 9 + 1)
    Slide1.Shapes("WindowTitleAppGallery:" & AppID).TextFrame.TextRange.Text = "Gallery - Page " & CStr(Page)
End Sub

Sub AppGalleryPrevPage(Shp As Shape)
    Dim AppID As String
    Dim Page As Integer
    AppID = GetAppID(Shp)
    Page = CInt(Replace(Slide1.Shapes("WindowTitleAppGallery:" & AppID).TextFrame.TextRange.Text, "Gallery - Page ", ""))
    Page = Page - 1
    If Page = 1 Then
        Slide1.Shapes("ButtonBackArrowAppGallery:" & AppID).Fill.Transparency = 1
        Slide1.Shapes("ButtonBackArrowAppGallery:" & AppID).TextFrame2.TextRange.Font.Fill.Transparency = 1
        Slide1.Shapes("ButtonBackArrowAppGallery:" & AppID).ActionSettings(ppMouseClick).Run = ""
    Else
        Slide1.Shapes("ButtonBackArrowAppGallery:" & AppID).Fill.Transparency = 0
        Slide1.Shapes("ButtonBackArrowAppGallery:" & AppID).TextFrame2.TextRange.Font.Fill.Transparency = 0
        Slide1.Shapes("ButtonBackArrowAppGallery:" & AppID).ActionSettings(ppMouseClick).Run = "AppGalleryPrevPage"
    End If
    If Page < 1 Then
        Exit Sub
    End If
    ClearPics AppID
    AppGalleryFindPics AppID, CStr((Page - 1) * 9 + 1)
    Slide1.Shapes("WindowTitleAppGallery:" & AppID).TextFrame.TextRange.Text = "Gallery - Page " & CStr(Page)
End Sub

Sub AppGallerySizeChanged(AppID As String)
    Dim L1 As Shape
    Dim P1 As Shape
    Dim B1 As Shape
    Set L1 = Slide1.Shapes("TitleLabelAppGallery:" & AppID)
    Set P1 = Slide1.Shapes("Pic1AppGallery:" & AppID)
    Set B1 = Slide1.Shapes("ButtonBackArrowAppGallery:" & AppID)
    L1.Top = P1.Top - L1.Height - 2
    L1.TextFrame.TextRange.Font.Size = B1.Height
End Sub

Sub AppGalleryFindPics(AppID As String, Optional StartRange As Integer = 1)
    Dim Pics() As String
    Dim PicsList As String
    Dim Shp As Shape
    Dim EndRange As Integer
    Dim IDX As Integer
    Dim PicID As Integer
    PicID = 1
    EndRange = StartRange + 8
    For Each Shp In Slide10.Shapes
        Dim Tokens() As String
        Dim Extension As String
        Tokens = Split(Shp.Name, ".")
        Extension = FsAssoc(Tokens(UBound(Tokens)))
        If InStr(1, Shp.Name, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/") = 1 And Extension = "PictureView" Then
            If Shp.Name <> "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/UserPic.png" And Shp.Name <> "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Background.png" Then
                PicsList = PicsList & "\" & Shp.Name
            End If
        End If
    Next Shp
    Pics = Split(PicsList, "\")
    If EndRange > UBound(Pics) Then
        EndRange = UBound(Pics)
        Slide1.Shapes("ButtonNextArrowAppGallery:" & AppID).Fill.Transparency = 1
        Slide1.Shapes("ButtonNextArrowAppGallery:" & AppID).TextFrame2.TextRange.Font.Fill.Transparency = 1
        Slide1.Shapes("ButtonNextArrowAppGallery:" & AppID).ActionSettings(ppMouseClick).Run = ""
    Else
        Slide1.Shapes("ButtonNextArrowAppGallery:" & AppID).Fill.Transparency = 0
        Slide1.Shapes("ButtonNextArrowAppGallery:" & AppID).TextFrame2.TextRange.Font.Fill.Transparency = 0
        Slide1.Shapes("ButtonNextArrowAppGallery:" & AppID).ActionSettings(ppMouseClick).Run = "AppGalleryNextPage"
    End If
    For IDX = StartRange To EndRange Step 1
        PreparePic Pics(IDX)
        With Slide1.Shapes("Pic" & CStr(PicID) & "AppGallery:" & AppID)
            .Fill.UserPicture Environ("TEMP") & "\Userpic.PNG"
            .ActionSettings(ppMouseClick).Run = "AppGalleryShowPic"
        End With
        PicID = PicID + 1
    Next IDX
    If UBound(Pics) = 9 Then
        Slide1.Shapes("ButtonNextArrowAppGallery:" & AppID).Fill.Transparency = 1
        Slide1.Shapes("ButtonNextArrowAppGallery:" & AppID).TextFrame2.TextRange.Font.Fill.Transparency = 1
        Slide1.Shapes("ButtonNextArrowAppGallery:" & AppID).ActionSettings(ppMouseClick).Run = ""
    End If
End Sub

Function AppGalleryGetPic(ID As Integer) As String
    Dim Pics() As String
    Dim PicsList As String
    For Each Shp In Slide10.Shapes
        Dim Tokens() As String
        Dim Extension As String
        Tokens = Split(Shp.Name, ".")
        Extension = FsAssoc(Tokens(UBound(Tokens)))
        If InStr(1, Shp.Name, "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/") = 1 And Extension = "PictureView" Then
            If Shp.Name <> "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/UserPic.png" And Shp.Name <> "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Background.png" Then
                PicsList = PicsList & "\" & Shp.Name
            End If
        End If
    Next Shp
    Pics = Split(PicsList, "\")
    For IDX = 1 To UBound(Pics) Step 1
        If ID = IDX Then
            AppGalleryGetPic = Pics(IDX)
            Exit Function
        End If
    Next IDX
End Function

Sub AppGalleryShowPic(Shp As Shape)
    Dim AppID As String
    Dim PicFile As String
    AppID = GetAppID(Shp)
    Page = CInt(Replace(Slide1.Shapes("WindowTitleAppGallery:" & AppID).TextFrame.TextRange.Text, "Gallery - Page ", "")) - 1
    PicFile = AppGalleryGetPic(CInt(Replace(Replace(Shp.Name, "Pic", ""), "AppGallery:" & AppID, "")) + (Page * 9))
    GetFileRef(PicFile).Copy
    With Slide27.Shapes.Paste
        .Left = 0
        .Top = 0
        .Width = ActivePresentation.PageSetup.SlideWidth
        .Height = ActivePresentation.PageSetup.SlideHeight
        .Visible = msoTrue
        .Name = "FullImage"
    End With
    AppGalleryShowControls
    ActivePresentation.SlideShowWindow.View.GotoSlide 28
End Sub


Sub AppGalleryExitFs(Shp As Shape)
    For I = Slide27.Shapes.Count To 1 Step -1
        If Slide27.Shapes(I).Name <> "SlideShowWindow" Then
            Slide27.Shapes(I).Delete
        End If
    Next I
    ActivePresentation.SlideShowWindow.View.GotoSlide 4
End Sub

Sub AppGalleryHideControls(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide27.Shapes("FullImage").ActionSettings(ppMouseClick).Run = "AppGalleryShowControls"
End Sub

Sub AppGalleryShowControls()
    Dim Theme As String
    Theme = "/Defaults/Themes/Default.thm"
    If FileStreamsExist("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm") Then Theme = "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Theme.thm"
    GetFileRef("/Defaults/PicView.grp").Copy
    With Slide27.Shapes.Paste
        .Name = "WControls"
        .Visible = msoTrue
        CopyFillFormat GetFileRef(Theme).GroupItems("Button"), Slide27.Shapes("ButtonBack")
        CopyFillFormat GetFileRef(Theme).GroupItems("Button"), Slide27.Shapes("ButtonDesktopPic")
        CopyFillFormat GetFileRef(Theme).GroupItems("Button"), Slide27.Shapes("ButtonUserPic")
        CopyFillFormat GetFileRef(Theme).GroupItems("Button"), Slide27.Shapes("ButtonHideControls")
    End With
    Slide27.Shapes("FullImage").ActionSettings(ppMouseClick).Run = ""
End Sub

Sub AppGallerySetBackground(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide27.Export Environ("TEMP") & "/Userpic.PNG", "PNG"
    SetFilePic "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Background.png", Environ("TEMP") & "/Userpic.PNG"
    
    With Slide1.Background.Fill
    .UserPicture Environ("TEMP") & "/Userpic.PNG"
    End With
    
    With Slide4.Background.Fill
    .UserPicture Environ("TEMP") & "/Userpic.PNG"
    End With
    
    With Slide14.Background.Fill
    .UserPicture Environ("TEMP") & "/Userpic.PNG"
    End With
    
    With Slide16.Background.Fill
    .UserPicture Environ("TEMP") & "/Userpic.PNG"
    End With
    With Slide17.Background.Fill
    .UserPicture Environ("TEMP") & "/Userpic.PNG"
    End With
    Slide1.Shapes("BackgroundImg").Delete
    GetFileRef("/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Background.png").Copy
    With Slide1.Shapes.Paste
        .Name = "BackgroundImg"
        .Left = 0
        .Top = 0
        .Width = ActivePresentation.PageSetup.SlideWidth
        .Height = ActivePresentation.PageSetup.SlideHeight
        .ZOrder msoSendToBack
        .Visible = msoTrue
    End With
    AppGalleryShowControls
End Sub

Sub AppGallerySetUserPic(Shp As Shape)
    Shp.ParentGroup.Delete
    Slide27.Export Environ("TEMP") & "/Userpic.PNG", "PNG"
    SetFilePic "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/UserPic.png", Environ("TEMP") & "/Userpic.PNG"
    AppGalleryShowControls
End Sub