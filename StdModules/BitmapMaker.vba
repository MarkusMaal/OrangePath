Type typHEADER
     strType As String * 2  ' Signature of file = "BM"
     lngSize As Long        ' File size
     intRes1 As Integer     ' reserved = 0
     intRes2 As Integer     ' reserved = 0
     lngOffset As Long      ' offset to the bitmap data (bits)
 End Type
 
 Type typINFOHEADER
     lngSize As Long        ' Size
     lngWidth As Long       ' Height
     lngHeight As Long      ' Length
     intPlanes As Integer   ' Number of image planes in file
     intBits As Integer     ' Number of bits per pixel
     lngCompression As Long ' Compression type (set to zero)
     lngImageSize As Long   ' Image size (bytes, set to zero)
     lngxResolution As Long ' Device resolution (set to zero)
     lngyResolution As Long ' Device resolution (set to zero)
     lngColorCount As Long  ' Number of colors (set to zero for 24 bits)
     lngImportantColors As Long ' "Important" colors (set to zero)
 End Type
 
 Type typPIXEL
     bytB As Byte    ' Blue
     bytG As Byte    ' Green
     bytR As Byte    ' Red
 End Type
 
 Type typBITMAPFILE
     bmfh As typHEADER
     bmfi As typINFOHEADER
     bmbits() As Byte
 End Type
 
 Sub ExportBMP(ByVal Data As String)
     SplitData = Split(Data, ";")
     Dim bmpFile As typBITMAPFILE
     Dim lngRowSize As Long
     Dim lngPixelArraySize As Long
     Dim lngFileSize As Long
     Dim J, k, l, x As Integer
     Dim bytRed, bytGreen, bytBlue As Integer
     Dim lngRGBColoer() As Long
 
     Dim strBMP As String
 
     With bmpFile
 
         With .bmfh
             .strType = "BM"
             .lngSize = 0
             .intRes1 = 0
             .intRes2 = 0
             .lngOffset = 54
         End With
         With .bmfi
             .lngSize = 40
             .lngWidth = 21
             .lngHeight = 16
             .intPlanes = 1
             .intBits = 24
             .lngCompression = 0
             .lngImageSize = 0
             .lngxResolution = 0
             .lngyResolution = 0
             .lngColorCount = 0
             .lngImportantColors = 0
         End With
         lngRowSize = Round(.bmfi.intBits * .bmfi.lngWidth / 32) * 4
         lngPixelArraySize = lngRowSize * .bmfi.lngHeight
 
         ReDim .bmbits(lngPixelArraySize)
         ReDim lngRGBColor(21, 21)
         pos = UBound(SplitData) - 21
         For J = 0 To 15
             For x = 0 To 20
                 Pixel = SplitData(pos)
                 PixelRGB = CLng(Pixel)
                 R = PixelRGB Mod 256
                 G = PixelRGB \ 256 Mod 256
                 B = PixelRGB \ 65536 Mod 256
                 ' B
                 k = k + 1
                 .bmbits(k) = G
                 ' G
                 k = k + 1
                 .bmbits(k) = R
                 ' R
                 k = k + 1
                 .bmbits(k) = B
                 pos = pos + 1
             Next x
             pos = pos - 41
         Next J
         .bmfh.lngSize = 14 + 40 + lngPixelArraySize
      End With ' Defining bmpFile
     strBMP = "D:\Ajutised failid\export.bmp"
     Open strBMP For Binary Access Write As 1 Len = 1
         Put 1, 1, bmpFile.bmfh
         Put 1, , bmpFile.bmfi
         Put 1, , bmpFile.bmbits
     Close
 End Sub
