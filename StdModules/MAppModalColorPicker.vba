' ModalColorPicker app (Generated from devCreateApp)

' This is executed when the application is launched
Sub AppModalColorPicker()
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "ModalColorPicker"
    Slide2.Shapes("AppModalColorPicker").Visible = msoTrue
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
    SetVar "Hue", "1"
    SetVar "Sat", "1"
    SetVar "Lum", "1"
    Slide2.Shapes("AppModalColorPicker").Visible = msoFalse
    Slide1.Shapes("WindowTitleAppModalColorPicker:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).TextFrame.TextRange.Text = "Color picker"
End Sub

Sub AppModalColorPickerLumSatPicker(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim GetCursorPositionX1 As Single
    Dim GetCursorPositionY1 As Single
    Dim Hue As Integer
    Dim Sat As Integer
    Dim Lum As Integer
    
    GetCursorPositionX1 = GetCursorX - Shp.Left + 21
    GetCursorPositionY1 = GetCursorY - Shp.Top - 19
    
    
    Hue = CInt(CheckVars("%Hue%"))
    Sat = GetCursorPositionY1 / Shp.Height * 255 + 37
    Lum = GetCursorPositionX1 / Shp.Width * 255
    If Slide1.Shapes("GrayscaleCheckAppModalColorPicker:" & AppID).Fill.ForeColor.RGB = Slide1.ColorScheme.Colors(ppFill) Then
        Sat = 1
        Slide1.Shapes("SaturationPickerAppModalColorPicker:" & AppID).Visible = msoFalse
    Else
        Slide1.Shapes("SaturationPickerAppModalColorPicker:" & AppID).Visible = msoTrue
    End If
    
    Dim Result As Long
    Result = HSLtoRGB(Hue, Sat, Lum)
    SetVar "Hue", CStr(CInt(Hue))
    SetVar "Sat", CStr(CInt(Sat))
    SetVar "Lum", CStr(CInt(Lum))
    Slide1.Shapes("CurrentColorAppModalColorPicker:" & AppID).Fill.ForeColor.RGB = Result
End Sub

Sub AppModalColorPickerQuickColor(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    SetVar "InputValue", CStr(Shp.Fill.ForeColor.RGB)
    UnsetVar "Hue"
    UnsetVar "Sat"
    UnsetVar "Lum"
    Shp.ParentGroup.Delete
    Application.Run CheckVars("%Macro%")
    UnsetVar "Macro"
End Sub

Sub AppModalColorPickerHuePicker(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    Dim GetCursorPositionX1 As Single
    Dim GetCursorPositionY1 As Single
    Dim Hue As Integer
    Dim Sat As Integer
    Dim Lum As Integer
    
    GetCursorPositionX1 = GetCursorX - Shp.Left + 21
    GetCursorPositionY1 = GetCursorY - Shp.Top - 18
    
    Hue = GetCursorPositionY1 / Shp.Height * 270 + 10
    Sat = CInt(CheckVars("%Sat%"))
    Lum = CInt(CheckVars("%Lum%"))
    
    Dim Result As Long
    Dim FullHue As Long
    Result = HSLtoRGB(Hue, Sat, Lum)
    FullHue = HSLtoRGB(Hue, 255, 128)
    SetVar "Hue", CStr(CInt(Hue))
    SetVar "Sat", CStr(CInt(Sat))
    SetVar "Lum", CStr(CInt(Lum))
    Slide1.Shapes("SaturationPickerAppModalColorPicker:" & AppID).Fill.GradientStops(2).Color.RGB = FullHue
    Slide1.Shapes("CurrentColorAppModalColorPicker:" & AppID).Fill.ForeColor.RGB = Result
End Sub

Sub AppModalColorPickerOkClicked(Shp As Shape)
    Dim AppID As String
    AppID = GetAppID(Shp)
    SetVar "InputValue", CStr(Slide1.Shapes("CurrentColorAppModalColorPicker:" & AppID).Fill.ForeColor.RGB)
    UnsetVar "Hue"
    UnsetVar "Sat"
    UnsetVar "Lum"
    Shp.ParentGroup.Delete
    Application.Run CheckVars("%Macro%")
    UnsetVar "Macro"
End Sub

Function HSLtoRGB(Hue As Integer, Saturation As Integer, _
  Luminance As Integer) As Long
    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
    Dim C As Double
    Dim x As Double
    Dim m As Double
    Dim rfrac As Double
    Dim gfrac As Double
    Dim bfrac As Double
    Dim hangle As Double
    Dim hfrac As Double
    Dim sfrac As Double
    Dim lfrac As Double

    If (Saturation = 0) Then
        R = 255
        G = 255
        B = 255
    Else
        lfrac = Luminance / 255
        hangle = Hue / 255 * 360
        sfrac = Saturation / 255
        C = (1 - Abs(2 * lfrac - 1)) * sfrac
        hfrac = hangle / 60
        hfrac = hfrac - Int(hfrac / 2) * 2 'fmod calc
        x = (1 - Abs(hfrac - 1)) * C
        m = lfrac - C / 2
        Select Case hangle
            Case Is < 60
                rfrac = C
                gfrac = x
                bfrac = 0
            Case Is < 120
                rfrac = x
                gfrac = C
                bfrac = 0
            Case Is < 180
                rfrac = 0
                gfrac = C
                bfrac = x
            Case Is < 240
                rfrac = 0
                gfrac = x
                bfrac = C
            Case Is < 300
                rfrac = x
                gfrac = 0
                bfrac = C
            Case Else
                rfrac = C
                gfrac = 0
                bfrac = x
        End Select
        R = Round((rfrac + m) * 255)
        G = Round((gfrac + m) * 255)
        B = Round((bfrac + m) * 255)
    End If
    HSLtoRGB = RGB(R, G, B)
End Function