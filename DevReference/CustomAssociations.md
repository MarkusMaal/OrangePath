# Creating custom file associations

If your application saves a file in a specific file format, you may want to create an association for it, so that when the user clicks on the file with a specific extension, it'll be opened with your app.

## Association functions

When you use devCreateApp, the following functions will be created:

```VB
' This gets executed when a user clicks a file, which is associated with this application
Sub AssocHello(Shp As Shape)
    Dim Filename As String
    Dim AppID As String
    AppID = GetAppID(Shp)
    Filename = Slide1.Shapes("PathAppFiles:" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text
    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "Hello"
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    CreateNewWindow
End Sub

' This gets executed when a user clicks icon of a file, which is associated with this application
Sub AssocIHello(Shp As Shape)
    Dim ShapeName As String
    ShapeName = Replace(Shp.Name, "Icon", "Label")
    AssocHello Slide1.Shapes(ShapeName)
End Sub
```

The `Assoc<AppName>` allows you to do stuff when a file is opened with your app. Filename defines where the loaded file is located. Insert new code after the `CreateNewWindow` line. `AssocI<AppName>` is a subroutine you shouldn't delete, but also don't really have to worry about, since it's just dealing with redirecting icon click to label click.

## Creating file associations

You can create a custom function for checking and defining file assocations, which you could call from the app initialization subroutine.

```VB
    Sub AppHelloCheckAssoc()
        If GetSysConfig("IconHello") = "*" Then
            SaveSysConfig "IconHello", "/System/Icons/Hello.emf"
        End If
        If GetSysConfig("assochel") = "*" Then
            SaveSysConfig "assochel", "Hello"
        End If
    End Sub
```

The "Icon" part is optional, but allows you to define a custom icon for your file association.

To define a specific file extension, simply call SaveSysConfig with the key of "assoc<extension>" and value of "<AppName>".

## Creating the icon for file association

If you go to design slide without running any macros, you may notice there's an example shape for a file assocation:

![1cc819c04751be2595438db8c08d817b.png](./Pics/1cc819c04751be2595438db8c08d817b.png)

Copy and paste it, then ungroup the pasted shape. While desiging the icon, you can use whatever colors you want, but make sure that it stays within bounds.

![96b0910781aa014beaacef9d5dafac44.png](./Pics/96b0910781aa014beaacef9d5dafac44.png)

After designing the association icon, select all shapes for it and right click to save as picture. When choosing file format, make sure you select **Enhanced Metafile (*.emf)**.

After saving the image, you can delete the original shape. Next run **aaDevEditApp** macro and select your app. From there, you can import the .emf file and add it to your app. Make sure that the image is named "AssocIcon" before running **aaDevRefreshApp**. Now, when you run the app, the icon should be automatically saved as `/System/Icons/<AppName>.emf`.