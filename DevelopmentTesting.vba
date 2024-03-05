Sub devNameThatSlide()
  MsgBox "The actual name of the selected slide in code is " & ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.slideIndex).Name, vbInformation
End Sub

Function GetModuleContent(ByVal moduleName As String) As String
    Dim vbComp As Object
    
    ' Search for the module by name
    For Each vbComp In ActivePresentation.VBProject.VBComponents
        If vbComp.Type = vbext_ct_StdModule And vbComp.Name = moduleName Then
            ' Get the content of the module as a string
            GetModuleContent = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            Exit Function
        End If
    Next vbComp
    
    ' Return an empty string if the module is not found
    GetModuleContent = ""
End Function


Sub devSerializeSlide()
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim vbComp As Object
    
    ' Search for the module by name
    For Each vbComp In ActivePresentation.VBProject.VBComponents
        
        If vbComp.Type = vbext_ct_Document Then
            moduleName = vbComp.Name
            Dim Fileout As Object
            Set Fileout = fso.CreateTextFile("C:\Users\marku\Documents\OrangePath_VBA\DocModules\" & moduleName & ".vba", True, True)
            Fileout.Write vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            Fileout.Close
        End If
    Next vbComp
    
    MsgBox "DocModules exported to plaintext successfully", vbInformation
End Sub

Sub devSerializeApp()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim vbComp As Object
    
    ' Search for the module by name
    For Each vbComp In ActivePresentation.VBProject.VBComponents
        
        If vbComp.Type = vbext_ct_StdModule Then
            moduleName = vbComp.Name
            Dim Fileout As Object
            Set Fileout = fso.CreateTextFile("C:\Users\marku\Documents\OrangePath_VBA\" & moduleName & ".vba", True, True)
            Fileout.Write GetModuleContent(moduleName)
            Fileout.Close
        End If
    Next vbComp
    
    MsgBox "Modules exported to plaintext successfully", vbInformation
End Sub

' Generate a new OP application automatically
Sub devCreateApp()
    Dim Name As String
    Name = InputBox("Enter the name of your app" & vbCrLf & vbCrLf & "Disallowed characters: spaces, underscores, colons" & vbCrLf & vbCrLf & "It is strongly recommended to use PascalCase when specifying an application name.")
    If InStr(Name, " ") Then
        MsgBox "Application name cannot contain spaces.", vbCritical, "Application builder"
        Exit Sub
    ElseIf InStr(Name, "_") Then
        MsgBox "Application name cannot contain underscore characters.", vbCritical, "Application builder"
        Exit Sub
    ElseIf InStr(Name, ":") Then
        MsgBox "Application name cannot contain colons.", vbCritical, "Application builder"
        Exit Sub
    ElseIf InStr(Name, "App") Then
        MsgBox "Application name cannot contain the word ""App"".", vbCritical, "Application builder"
        Exit Sub
    End If
    Dim FriendlyName As String
    Dim Access As String
    Access = "Everyone"
    FriendlyName = InputBox("Enter a user friendly name for your app. This can contain spaces and will appear in the application menu and taskbar icon.")
    
    Dim result As VbMsgBoxResult
    result = MsgBox("Do you want this application to be visible for everyone, even guest accounts?", vbQuestion + vbYesNoCancel, "Access control")
    If result = vbCancel Then
        MsgBox "Cancelled, no changes made", vbInformation, "Application builder"
        Exit Sub
    End If
    If result = vbNo Then
        Access = "Administrators"
    End If
    
    Dim NewApp As shape
    Set NewApp = GetFileRef("/Defaults/Themes/Default.thm")
    NewApp.Copy
    With Slide2.Shapes.Paste
        .Name = "App" & Name
        .GroupItems("WindowTitle").TextFrame.TextRange.Text = FriendlyName
        .GroupItems("Window").TextFrame.TextRange.Text = ""
    End With
    For i = 1 To Slide2.Shapes("App" & Name).GroupItems.Count Step 1
        With Slide2.Shapes("App" & Name).GroupItems(i)
            .Name = .Name & "App" & Name & "_"
            .Visible = msoTrue
        End With
    Next i
    Slide24.Shapes("AppDefault:Properties").Copy
    With Slide25.Shapes.Paste
        .Name = "App" & Name & ":Properties"
        .TextFrame.TextRange.Text = Access & ":" & FriendlyName
    End With
    Slide29.Shapes("AppDefault:Icon").Copy
    With Slide25.Shapes.Paste
        .Name = "App" & Name & ":Icon"
    End With
    For i = 1 To Slide25.Shapes("App" & Name & ":Icon").GroupItems.Count
        Slide25.Shapes("App" & Name & ":Icon").GroupItems(i).ActionSettings(ppMouseClick).Run = "App" & Name
    Next i
    result = MsgBox("Would you like me to automatically create a VBA module, which includes some boilerplate code required for the application to laucnh?" & vbCrLf & "Note: Selecting ""No"" here will leave you with an application that cannot be launched", vbQuestion + vbYesNo, "VBA modules")
    If result = vbYes Then
         Dim newModule As Object
        Dim moduleName As String
        
        ' Define the desired module name
        moduleName = "MApp" & Name
        
        ' Add a new module to the PowerPoint VBA project
        Set newModule = ActivePresentation.VBProject.VBComponents.Add(vbext_ct_StdModule)
        
        ' Rename the newly added module
        newModule.Name = moduleName
        
        ' Generates sample code (spooky)
        newModule.CodeModule.AddFromString "' " & Name & " app (Generated from devCreateApp)" & vbCrLf & vbCrLf & _
                                          "' This is executed when the application is launched" & vbCrLf & _
                                          "Sub App" & Name & "(Shp As Shape)" & vbCrLf & _
                                          "    Shp.ParentGroup.Delete" & vbCrLf & _
                                          "    Slide1.Shapes(""AppCreatingEvent"").TextFrame.TextRange.Text = """ & Name & """" & vbCrLf & _
                                          "    Slide2.Shapes(""App" & Name & """).Visible = msoTrue" & vbCrLf & _
                                          "    ActivePresentation.SlideShowWindow.View.GotoSlide (4)" & vbCrLf & _
                                          "    CreateNewWindow" & vbCrLf & _
                                          "    Slide2.Shapes(""App" & Name & """).Visible = msoFalse" & vbCrLf & _
                                          "End Sub" & vbCrLf & vbCrLf & _
                                          "' This gets executed when a user clicks a file, which is associated with this application" & vbCrLf & _
                                          "Sub Assoc" & Name & "(Shp As Shape)" & vbCrLf & _
                                          "    Dim Filename As String" & vbCrLf & _
                                          "    Dim AppID As String" & vbCrLf & _
                                          "    AppID = GetAppID(Shp)" & vbCrLf & _
                                          "    Filename = Slide1.Shapes(""PathAppFiles:"" & AppID).TextFrame.TextRange.Text & Slide1.Shapes(Shp.Name).TextFrame.TextRange.Text" & vbCrLf & _
                                          "    Slide1.Shapes(""AppCreatingEvent"").TextFrame.TextRange.Text = """ & Name & """" & vbCrLf & _
                                          "    ActivePresentation.SlideShowWindow.View.GotoSlide (4)" & vbCrLf & "    CreateNewWindow" & vbCrLf & _
                                          "End Sub" & vbCrLf & vbCrLf & _
                                          "' This gets executed when a user clicks icon of a file, which is associated with this application" & vbCrLf & _
                                          "Sub AssocI" & Name & "(Shp As Shape)" & vbCrLf & _
                                          "    Dim ShapeName As String" & vbCrLf & _
                                          "    ShapeName = Replace(Shp.Name, ""Icon"", ""Label"")" & vbCrLf & _
                                          "    Assoc" & Name & " Slide1.Shapes(ShapeName)" & vbCrLf & _
                                          "End Sub" & vbCrLf
    End If
    MsgBox "Created " & FriendlyName & " (" & Name & ") successfully. This application should now be accessible in the application menu with a default icon. If you want to modify the application icon, run the devDesignIcon macro.", vbInformation, "Application builder"
End Sub

Sub devDesignIcon()
    On Error GoTo NotExist
    Dim Name As String
    Name = InputBox("Enter the name of the app you wish to edit the icon for")
    Slide24.Select
    For i = Slide24.Shapes.Count To 1 Step -1
        If Slide24.Shapes(i).Name <> "DesignSlideLabel" Then
            Slide24.Shapes(i).Visible = msoFalse
        End If
    Next i
    Slide25.Shapes("App" & Name & ":Icon").Copy
    With Slide24.Shapes.Paste
        .Ungroup
    End With
    MsgBox "Please run the devDeployIcon macro after you finish designing the icon (OrangePath OS might not operate correctly if you don't run that macro). Note that the icon must fit within the rounded rectangle shape and no shape should go outside of it. Keep shapes separated, do not group them. Action settings are applied automatically during the devDeployIcon macro.", vbInformation, "Icon design"
    Exit Sub
NotExist:
    MsgBox "The specified application doesn't exist", vbCritical, "Icon designer"
End Sub

Sub devDeployIcon()
    On Error GoTo DeployErrd
    Slide29.Select
    If Slide24.Shapes("FileIcon_*").Visible = msoTrue Then
        MsgBox "Not designing an icon", vbCritical, "Icon design"
        Exit Sub
    End If
    Dim Name As String
    Name = InputBox("Enter the application name you wish to deploy this icon for")
    
    Dim Sld As slide
    Set Sld = Slide24
    Dim Shp2 As shape
    Dim Shapes As String
    Shapes = ""
    ShapeName = "App" & Name & ":Icon"
    For Each Shp2 In Sld.Shapes()
        If Shp2.Visible = msoTrue And Shp2.Name <> "DesignSlideLabel" Then
            'If InStr(Shp2.Name, "AXTextBox") Then ApplyTbAttribs Shp2
            
            Shapes = Shapes & Shp2.Name & ","
        End If
    Next Shp2
    SplitShapes = Split(Shapes, ",")
    UJ = CInt(UBound(SplitShapes))
    Dim ShapesX() As String
    
    ReDim ShapesX(UJ)
    For i = CInt(UBound(SplitShapes)) To 0 Step -1
        CShape = SplitShapes(i)
        ShapesX(i) = SplitShapes(i)
    Next
    With Sld.Shapes.Range(ShapesX).Group
        .Name = ShapeName
    End With
    
    Slide24.Shapes("App" & Name & ":Icon").Copy
    Slide25.Shapes("App" & Name & ":Icon").Delete
    With Slide25.Shapes.Paste
        .Visible = msoFalse
        For j = 1 To .GroupItems.Count Step 1
            .GroupItems(j).Visible = msoTrue
            .GroupItems(j).ActionSettings(ppMouseClick).Run = "App" & Name
        Next j
    End With
    Slide24.Shapes("App" & Name & ":Icon").Delete
    For i = Slide24.Shapes.Count To 1 Step -1
        Slide24.Shapes(i).Visible = msoTrue
    Next i
    MsgBox "Icon deployed successfully", vbInformation, "Icon design"
    Exit Sub
DeployErr:
    MsgBox "Something went wrong while trying to deploy the icon. Please try again!", vbCritical, "Icon design"
End Sub

Sub devRemoveApp()
    Dim Name As String
    Name = InputBox("Enter the name of the app you wish to delete" & vbCrLf & "Warning: You cannot undo this action!")
    
    
    Dim vbComp As Object
    Dim deleted As Boolean
    deleted = False
    ' Loop through each module in the VBA project
    For Each vbComp In ActivePresentation.VBProject.VBComponents
        ' Check if the current component is a module and its name matches the one to delete
        If vbComp.Type = vbext_ct_StdModule And vbComp.Name = "MApp" & Name Then
            ' Remove the module
            ActivePresentation.VBProject.VBComponents.Remove vbComp
            deleted = True
            Exit For ' Exit the loop once the module is deleted
        End If
    Next vbComp
    If deleted = False Then
        MsgBox "This application does not exist. No changes were made!", vbCritical
        Exit Sub
    End If
    Slide25.Shapes("App" & Name & ":Properties").Delete
    Slide25.Shapes("App" & Name & ":Icon").Delete
    Slide2.Shapes("App" & Name).Delete
    MsgBox "Application removed!", vbInformation
End Sub

Sub GetTaskbarLefts()
    For Each Shp In Slide1.Shapes
        If InStr(Shp.Name, "TaskIcon:") Then
            MsgBox CStr(CInt(Shp.Left)) & " " & CStr(CInt(Shp.Width))
        End If
    Next Shp
End Sub

Sub CleanDesktop()
    Dim i As Integer
    For i = Slide1.Shapes.Count To 1 Step -1
        If InStr(1, Slide1.Shapes(i).Name, "FileLabel") = 1 Then
            Slide1.Shapes(i).Delete
        ElseIf InStr(1, Slide1.Shapes(i).Name, "FileIcon") = 1 Then
            Slide1.Shapes(i).Delete
        End If
    Next i
End Sub

Sub devAddMacroActions()
    Dim inputShape As String
    inputShape = InputBox("Please enter the name of the App you want to add the action to (e.g. Calc)")
    For Each Shp In Slide25.Shapes("App" & inputShape & ":Icon").GroupItems
         With Shp.ActionSettings(ppMouseClick)
            .Run = "App" & inputShape
         End With
    Next Shp
    MsgBox "Macro added to App" & inputShape & ":Icon group", vbInformation, "Action automation script"
End Sub

Sub devListApps()
    ' A macro for displaying a list of installed applications
    Dim AppList As String
    AppList = ""
    For Each Shp In Slide25.Shapes
        If InStr(Shp.Name, ":Properties") Then
            FullName = Replace(Shp.Name, ":Properties", "")
            Props = Split(Shp.TextFrame.TextRange.Text, ":")
            Access = Props(0)
            FriendlyName = Props(1)
            PackageName = Right(FullName, Len(FullName) - 3)
            AppList = AppList & FriendlyName & " (" & Access & ")" & vbNewLine
        End If
    Next Shp
    MsgBox AppList
End Sub

Sub DTest()
    AppID = "2"
    MsgBox Slide1.AxTextBox.Text
    Args = Split(Slide1.AxTextBox.Text, " ")
    If UBound(Args) < 2 Then
        Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "The syntax of the command is incorrect."
        Exit Sub
    End If
    PID = Args(2)
    Tp = Args(1)
    If Not ShapeExists(Slide1, "RegularApp:" & PID) Then
        Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Invalid process ID"
        Exit Sub
    End If
    If Tp = "0" Then
        Slide5.Shapes("Bootlogo").ThreeD.RotationZ = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationZ
        Slide5.Shapes("Bootlogo").ThreeD.RotationY = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationY
        Slide5.Shapes("Bootlogo").ThreeD.RotationX = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationX
        Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully"
        Exit Sub
    ElseIf Tp = "1" Then
        Slide3.Shapes("Bootlogo").ThreeD.RotationZ = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationZ
        Slide3.Shapes("Bootlogo").ThreeD.RotationY = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationY
        Slide3.Shapes("Bootlogo").ThreeD.RotationX = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationX
        Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully"
        Exit Sub
    ElseIf Tp = "2" Then
        Slide7.Shapes("Bootlogo").ThreeD.RotationZ = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationZ
        Slide7.Shapes("Bootlogo").ThreeD.RotationY = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationY
        Slide7.Shapes("Bootlogo").ThreeD.RotationX = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationX
        Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully"
        Exit Sub
    End If
End Sub