# ShapeFS functions

ShapeFS is a filesystem inside OrangePath OS. All of the filesystem code is contained within the `OpFileSystem` module.

## FileStreamsExist

* Parameters: Filename As String
* Returns: Boolean

This function checks if either specified folder and/or streams exist. Example usage:

```VB
    Dim folderExists As Boolean
    folderExists = FileStreamsExist("/Users/DemoAcc/")
```

## FileExists

* Parameters: Filename As String, Optional Stream As String = ""
* Returns: Boolean

This function checks if a file or specified stream of a file exists. Example usage:

```VB
    Dim fileExist As Boolean
    fileExist = FileExists("/Users/DemoAcc/Test.txt")
```

## NewFolder

* Parameters: Dirname As String

This function creates a new folder at a specified location. Dirname mustn't end with a forward slash.

Example usage:

```VB
    NewFolder "/Users/DemoAcc/NewFolder"
```

## GetFiles

* Paramaters: Dirname As String
* Returns: String

Allows you to get a directory listing of a folder specified. Returned value will contain a list of names for files and subfolders separated by `vbNewLine`.

**NOTE**: `Dirname` MUST end with a slash in this case

Example usage:
```VB
    Dim home As String
    home = GetFiles("/Users/DemoAcc/")
```

Example of a returned value:
```
Pictures/
Videos/
Background.png
Password.txt
Theme.txt
```

## GetFileContent

* Parameters: Filename As String, Optional Stream As String = ""
* Returns: String

Allows you to get plain-text content of a file stored in ShapeFS. In FS this is stored inside the shape's `.TextFrame.TextRange.Text`. Returns "*" if access to the file is denied or it doesn't exist.

Example usage:
```VB
    Dim Txt As String
    Txt = GetFileContent("/Users/DemoAcc/Theme.txt")
```

You can also get the text file content from host in the same way.
```VB
    Dim Txt As String
    Txt = GetFileContent("C:\Users\Admin\Example.txt")
```

## GetFileRef

* Parameters: Filename As String, Optional Stream As String = ""
* Returns: Shape

Gets the actual shape of the file in ShapeFS and returns it. Useful for getting files, which are actually shapes.

Example usage:
```VB
    Dim Shp As Shape
    Set Shp = GetFileRef("/Users/DemoAcc/NewFolder/Test.shp")
```

## SetFileContent

* Parameters: Filename As String, Content As String, Optional Stream As String = ""

Allows you to save text content to a file. If the access is denied or the file fails to save for whatever reason, an error message is displayed.

Example usage:
```VB
    Dim Txt As String
    Dim Pth As String
    Txt = "Hello, world!"
    Pth = "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Hello.txt"
    SetFileContent Pth, Txt
```

## DeleteFile

* Parameters: Filename As String, Optional Stream As String = ""

Deletes the file specified. Displays an error message if access is denied or some other error occurs.

Example usage:
```VB
    Dim Pth As String
    Pth = "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Hello.txt"
    DeleteFile Pth
```

## DeleteDir

* Parameters: Dirname As String

Deletes a directory specified. Displays an error message if access is denied or some other error occurs.

Example usage:
```VB
    Dim Pth As String
    Pth = "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/"
    DeleteDir Pth
```

## CopyFile

* Parameters: Source As String, Destination As String

Copies a file to specified destination path. Allows for importing certain files from host C: drive. If copying operation files for some reason, an error message is displayed.

Local copy example:
```VB
    Dim Src As String
    Dim Dst As String
    Src = "/Users/" &  Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Hello.txt"
    Dst = "/Users/" &  Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Hello2.txt"
    CopyFile Src, Dst
```

File import example:
```VB
    Dim Src As String
    Dim Dst As String
    Src = "C:\Users\Admin\Example.txt"
    Dst = "/Users/" &  Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Example.txt"
    CopyFile Src, Dst
```

## FsAssoc

* Parameters: Extension As String
* Returns: String

Allows you to get associated application from the file extension. Returns an empty string if no association is found.

Example usage:
```VB
    Dim Ext As String
    Ext = "txt"
    ' Returns app associated with the txt file, "Notes" by default
    AppName = FsAssoc(Ext)
```

## RenameFile

* Parameters: Filename As String, Newname As String, Optional Stream As String = ""

Allows you to rename files.

Example usage:
```VB
    Dim oldFile As String
    Dim newFile As String
    oldFile = "/Users/" &  Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Hello2.txt"
    newFile = "/Users/" &  Slide1.Shapes("Username").TextFrame.TextRange.Text & "/HelloWorld.txt"
    RenameFile oldFile, newFile
```

## LinkMovie

* Parameters: Filepath As String

This function temporarily copies a video file from host filesystem as `/Temp/Movie.mov`. Used by file manager to play a video file directly from host filesystem.

**NOTE**: Anything from `/Temp` directory will be deleted on next login.

Example usage:

```VB
    LinkMovie "C:\Users\Admin\Videos\Test.mp4"
```

## ImportMovie

* Paramaters: Filepath As String, OP_path As String

Imports entire video file from host filesystem to ShapeFS. Filepath is the path to the video file on host and OP_path is the ShapeFS path, where the video will be imported to. This is different from LinkMovie, because here, the video is actually saved inside the presentation.

**NOTE**: Storing a large amount of video files inside a PowerPoint presentation can lead to large filesizes, which may cause instability and inability to store the presentation in VCS.

Example usage:
```VB
    ImportMovie "C:\Users\Admin\Videos\Test.mp4", "/Users/DemoAcc/Videos/Video.mp4"
```

## SetFilePic

* Parameters: Filename As String, Tempfile As String, Optional Stream As String = ""

Imports an image file from host filesystem to ShapeFS. Tempfile is full path to host file you wish to import and filename is the location in ShapeFS where to import the image to.

Example usage:
```VB
    Dim ShpFsPth As String
    Dim HostPth As String 
    ShpFsPth = "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Background.png"
    HostPth = "C:\Users\Admin\Pictures\funny.png"
    SetFilePic ShpFsPth, HostPth
```

## PreparePic

* Parameters: Filename As String, Optional Stream As String = ""

Exports image file as "%Temp%\Userpic.PNG". Useful for setting slide backgrounds.

Example usage:
```VB
    PreparePic "/Defaults/Images/Background.png"
```

## WriteGroup

* Parameters: Filename As String, SShp As ShapeRange, SizeKey As String

This function allows you to write an array of shapes as a group to ShapeFS. The following parameters are expected:
- **Filename**: The filename you wish to write the group to
- **SShp**: A range of shapes you wish to write to this group (see example below)
- **SizeKey**: Name of the shape, which is a reference for dimensions of the group

Example:
```VB
    Sub SaveGroupExample(Shp As Shape)
        ' Get App ID
        Dim AppID As String
        AppID = GetAppID(Shp)
        ' Declarations
        Dim AppName As String
        AppName = "Example"
        Dim Shp2 As shape
        Dim Shapes As String
        Dim Filename As String
        Shapes = ""
        Filename = "/Users/" & Slide1.Shapes("Username").TextFrame.TextRange.Text & "/Example.grp"
        ' Go through each shape in the window
        For Each Shp2 In Slide1.Shapes("RegularApp:" & AppID).GroupItems()
            ' Check if the shape name starts with "MyGrp"
            If InStr(1, Shp2.Name, "MyGrp") = 1 Then
                Shapes = Shapes & Shp2.Name & ","
            End If
        Next Shp2
        
        ' Create the shape range
        SplitShapes = Split(Shapes, ",")
        UJ = CInt(UBound(SplitShapes))
        Dim ShapesX() As String
        
        ReDim ShapesX(UJ)
        For i = 0 To CInt(UBound(SplitShapes) - 1)
            CShape = SplitShapes(i)
            If Not IsInArray(CStr(CShape), ShapesX) Then
                ShapesX(i) = SplitShapes(i)
            End If
        Next

        ' Write group
        WriteGroup Filename, Slide1.Shapes.Range(ShapesX), "MyGrpBackdropApp" & AppName & ":" & AppID
    End Sub

```

## ReadGroup

* Parameters: Filename As String, TSld As Slide, OffsetX As Integer, OffsetY As Integer, Ref As Shape, SizeX As Integer, SizeY As Integer

This function allows you to read a group from ShapeFS and paste it to a window. The following parameters are expected:
- **Filename**: Path to the file, where the group is stored
- **TSld**: Target slide to paste the group to
- **OffsetX**: Horizontal offset, where to paste the group (from left edge of slide)
- **OffsetY**: Vertical offset, where to paste the group (from top edge of slide)
- **Ref**: Window shape, where to paste the group
- **SizeX** : Width of the group before pasting
- **SizeY** : Height of the group before pasting

Example:

```VB
    Sub ReadGroupExample(Shp As Shape)
        Dim AppID As Integer
        AppID = GetAppID(Shp)
        Dim OffX As Integer
        Dim OffY As Integer
        Dim SizeX As Integer
        Dim SizeY As Integer
        Dim AppName As String
        AppName = "Example"
        OffX = Slide1.Shapes("RefApp" & AppName & ":" & AppID).Left
        OffY = Slide1.Shapes("RefApp" & AppName & ":" & AppID).Top
        SizeX = Slide1.Shapes("RefApp" & AppName & ":" & AppID).Width
        SizeY = Slide1.Shapes("RefApp" & AppName & ":" & AppID).Height
        ReadGroup TextValue, Slide1, OffX, OffY, Slide1.Shapes("RegularApp:" & AppID).GroupItems(1), SizeX, SizeY
    End Sub
```