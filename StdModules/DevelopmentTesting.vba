' Functions starting with zz are internal and should not be invoked manually

Sub aaDevNameThatSlide()
  MsgBox "The actual name of the selected slide in code is " & ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex).Name, vbInformation
End Sub

Function GetModuleContent(ByVal moduleName As String) As String
    Dim vbComp As Object
    
    ' Search for the module by name
    For Each vbComp In ActivePresentation.VBProject.VBComponents
        If vbComp.Name = moduleName Then
            ' Get the content of the module as a string
            GetModuleContent = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            Exit Function
        End If
    Next vbComp
    
    ' Return an empty string if the module is not found
    GetModuleContent = ""
End Function

Sub aaDevExport()
    Set dlgOpen = Application.FileDialog(Type:=msoFileDialogFolderPicker)
    Dim strResult As String
    strResult = ""
    With dlgOpen
        .Title = "Select folder, where the VCS repository is stored"
        .AllowMultiSelect = False
        If .Show = True Then
            strResult = .SelectedItems(1)
        End If
    End With
    If strResult = "" Then
        MsgBox "No path specified", vbCritical, "VCS export"
        Exit Sub
    End If
    zzSerializeSlide strResult
    zzSerializeApp strResult
    zzSerializeForms strResult
    ActivePresentation.SaveCopyAs strResult & "\Sunlight.pptx", ppSaveAsOpenXMLPresentation
    MsgBox "Exported presentation for VCS", vbInformation, "VCS export"
End Sub

Sub aaDevCheckDuplicates()
    Dim Shp As Shape
    Dim ShapesList As String
    ShapesList = ""
    For Each Shp In Slide24.Shapes
        If Not InStr(ShapesList, Shp.Name) Then
            ShapesList = ShapesList & Shp.Name & ","
        Else
            MsgBox "Duplicate shape: " & Shp.Name, vbExclamation, "Duplicate detector"
            Exit Sub
        End If
    Next Shp
    MsgBox "No duplicates detected", vbInformation, "Duplicate detector"
End Sub


Sub zzSerializeSlide(Dirname As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim vbComp As Object
    
    ' Search for the module by name
    For Each vbComp In ActivePresentation.VBProject.VBComponents
        
        If vbComp.Type = 100 Then
            moduleName = vbComp.Name
            Dim Fileout As Object
            Set Fileout = fso.CreateTextFile(Dirname & "\DocModules\" & moduleName & ".vba", True, True)
            Fileout.Write vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            Fileout.Close
        End If
    Next vbComp
End Sub

Sub zzSerializeApp(Dirname As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim vbComp As Object
    
    ' Search for the module by name
    For Each vbComp In ActivePresentation.VBProject.VBComponents
        If vbComp.Type = 1 Then
            moduleName = vbComp.Name
            Dim Fileout As Object
            Set Fileout = fso.CreateTextFile(Dirname & "\StdModules\" & moduleName & ".vba", True, True)
            Fileout.Write GetModuleContent(moduleName)
            Fileout.Close
        End If
    Next vbComp
End Sub

Sub zzSerializeForms(Dirname As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim vbComp As Object
    
    ' Search for the module by name
    For Each vbComp In ActivePresentation.VBProject.VBComponents
        Dim Fileout As Object
        If vbComp.Type = 3 Then
            moduleName = vbComp.Name
            Set Fileout = fso.CreateTextFile(Dirname & "\FormModules\" & moduleName & ".vba", True, True)
            Fileout.Write vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            Fileout.Close
        End If
    Next vbComp
End Sub

Sub aaDevRemoveApp()
    SetVar "Macro", "zzRemoveApp"
    With SelectApp
        Dim I As Integer
        For I = Slide2.Shapes.Count To 1 Step -1
            If InStr(Slide2.Shapes(I).Name, "App") Then
                .AppList.AddItem (Replace(Slide2.Shapes(I).Name, "App", ""))
            End If
        Next I
        .Show
    End With
End Sub

Function DetectDesignApp() As String
    Dim Shp As Shape
    For Each Shp In Slide24.Shapes
        If InStr(Shp.Name, "App") And Right(Shp.Name, 1) = "_" Then
            NameSplit = Split(Shp.Name, "App")
            Name = NameSplit(1)
            DetectDesignApp = Left(Name, Len(Name) - 1)
        End If
    Next Shp
End Function

Sub aaDevRefreshApp()
    Dim AppName As String
    AppName = DetectDesignApp
    If AppName = vbNullString Then
        MsgBox "Not designing an app right now", vbExclamation, "Refresh app"
        Exit Sub
    End If
    Dim Suffix As String
    Suffix = "App" & AppName & "_"
    Dim Shapes As String
    Dim Shp As Shape
    Shapes = ""
    For Each Shp In Slide24.Shapes
        If Right(Shp.Name, Len(Suffix)) <> Suffix And Shp.Visible = msoTrue And Shp.Name <> "DesignSlideLabel" Then
            Shp.Name = Shp.Name & Suffix
        End If
    Next Shp
    GetFileRef("/Defaults/Themes/Default.thm").GroupItems("Handle").Copy
    If ShapeExists(Slide24, "WindowApp" & AppName & "_") Then
        With Slide24.Shapes.Paste
            .Name = "HandleApp" & AppName & "_"
            .Left = Slide24.Shapes("WindowApp" & AppName & "_").Left + Slide24.Shapes("WindowApp" & AppName & "_").Width
            .Top = Slide24.Shapes("WindowApp" & AppName & "_").Top + Slide24.Shapes("WindowApp" & AppName & "_").Height
            .ActionSettings(ppMouseOver).Run = "ResizingWindow"
            .Visible = msoTrue
        End With
    End If
    For Each Shp In Slide24.Shapes
        If Shp.Visible = msoTrue And Shp.Name <> "DesignSlideLabel" Then
            Shapes = Shapes & Shp.Name & ","
        End If
    Next Shp
    
    ' Create the shape range
    SplitShapes = Split(Shapes, ",")
    UJ = CInt(UBound(SplitShapes))
    Dim ShapesX() As String
    
    ReDim ShapesX(UJ)
    For I = 0 To CInt(UBound(SplitShapes) - 1)
        CShape = SplitShapes(I)
        If Not IsInArray(CStr(CShape), ShapesX) Then
            ShapesX(I) = SplitShapes(I)
        End If
    Next
    
    With Slide24.Shapes.Range(ShapesX).Group
        .Name = "App" & AppName
    End With
    Slide2.Shapes("App" & AppName).Delete
    Slide24.Shapes("App" & AppName).Copy
    With Slide2.Shapes.Paste
        .Name = "App" & AppName
        .Visible = msoFalse
    End With
    Slide24.Shapes("App" & AppName).Delete
    For Each Shp In Slide24.Shapes
        Shp.Visible = msoTrue
    Next Shp
End Sub

Sub aaDevEditApp()
    SetVar "Macro", "zzEditApp"
    With SelectApp
        Dim I As Integer
        For I = Slide2.Shapes.Count To 1 Step -1
            If InStr(Slide2.Shapes(I).Name, "App") Then
                .AppList.AddItem (Replace(Slide2.Shapes(I).Name, "App", ""))
            End If
        Next I
        .Show
    End With
End Sub

Sub zzEditApp(AppName As String)
    On Error GoTo EndEdit
    Slide2.Shapes("App" & AppName).Copy
    Dim I As Integer
    For I = Slide24.Shapes.Count To 1 Step -1
        If Slide24.Shapes(I).Name <> "DesignSlideLabel" Then
            Slide24.Shapes(I).Visible = msoFalse
        End If
    Next I
    With Slide24.Shapes.Paste
        .Visible = msoTrue
        .Ungroup
    End With
    Dim Shp As Shape
    For Each Shp In Slide24.Shapes
        If InStr(1, Shp.Name, "HandleApp" & AppName) = 1 Then
            Shp.Delete
        End If
    Next Shp
    UnsetVar "Macro"
EndEdit:
    Slide24.Select
    Exit Sub
End Sub

Sub aaDevCreateApp()
    With devCreateAppDialog
        .Show
    End With
End Sub

Sub UninstallMAppEmpty()
    For Each vbcomponent In ActivePresentation.VBProject.VBComponents
        If vbcomponent.Name = "Slide1" Then
            vbcomponent.Name = "Slide1"
            Exit Sub
        End If
    Next vbcomponent
End Sub

' Generate a new OP application automatically
Sub zzCreateApp()
    Dim Name As String
    Name = CheckVars("%Name%")
    UnsetVar "Name"
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
    FriendlyName = CheckVars("%FriendlyName%")
    Access = CheckVars("%Access%")
    UnsetVar "Access"
    UnsetVar "FriendlyName"
    Dim NewApp As Shape
    Set NewApp = GetFileRef("/Defaults/Themes/Default.thm")
    NewApp.Copy
    With Slide2.Shapes.Paste
        .Name = "App" & Name
        .GroupItems("WindowTitle").Delete
        .GroupItems("WindowFrame").Delete
        .GroupItems("Close").Delete
        .GroupItems("Minimize").Delete
        .GroupItems("Button").Delete
        .GroupItems("Window").TextFrame.TextRange.Text = ""
    End With
    For I = 1 To Slide2.Shapes("App" & Name).GroupItems.Count Step 1
        With Slide2.Shapes("App" & Name).GroupItems(I)
            .Name = .Name & "App" & Name & "_"
            .Visible = msoTrue
        End With
    Next I
    If CheckVars("%Shortcuts%") = "True" Then
        Slide29.Shapes("AppDefault:Properties").Copy
        With Slide25.Shapes.Paste
            .Name = "App" & Name & ":Properties"
            .TextFrame.TextRange.Text = Access & ":" & FriendlyName
        End With
        Slide29.Shapes("AppDefault:Icon").Copy
        With Slide25.Shapes.Paste
            .Name = "App" & Name & ":Icon"
        End With
        For I = 1 To Slide25.Shapes("App" & Name & ":Icon").GroupItems.Count
            Slide25.Shapes("App" & Name & ":Icon").GroupItems(I).ActionSettings(ppMouseClick).Run = "App" & Name
        Next I
    End If
    If CheckVars("%GenModule%") = "True" Then
        Dim newModule As Object
        Dim moduleName As String
        
        ' Define the desired module name
        moduleName = "MApp" & Name
        
        ' Add a new module to the PowerPoint VBA project
        Set newModule = ActivePresentation.VBProject.VBComponents.Add(vbext_ct_StdModule)
        
        ' Rename the newly added module
        newModule.Name = moduleName
        If CheckVars("%GenCode%") = "True" Then
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
    End If
    UnsetVar "GenModule"
    UnsetVar "GenCode"
    MsgBox "Created " & FriendlyName & " (" & Name & ") successfully. This application should now be accessible in the application menu with a default icon. If you want to modify the application icon, run the devDesignIcon macro.", vbInformation, "Application builder"
End Sub
Sub aaDevDesignIcon()
    Dim Name As String
    SetVar "Macro", "zzDesignIcon"
    With SelectApp
        Dim I As Integer
        For I = Slide2.Shapes.Count To 1 Step -1
            If InStr(Slide2.Shapes(I).Name, "App") Then
                .AppList.AddItem (Replace(Slide2.Shapes(I).Name, "App", ""))
            End If
        Next I
        .Show
    End With
End Sub

Sub zzDesignIcon(AppName As String)
    On Error GoTo NotExist
    Name = AppName
    Slide24.Select
    For I = Slide24.Shapes.Count To 1 Step -1
        If Slide24.Shapes(I).Name <> "DesignSlideLabel" Then
            Slide24.Shapes(I).Visible = msoFalse
        End If
    Next I
    Slide25.Shapes("App" & Name & ":Icon").Copy
    With Slide24.Shapes.Paste
        .Ungroup
    End With
    MsgBox "Please run the devDeployIcon macro after you finish designing the icon (Sunlight OS might not operate correctly if you don't run that macro). Note that the icon must fit within the rounded rectangle shape and no shape should go outside of it. Keep shapes separated, do not group them. Action settings are applied automatically during the devDeployIcon macro.", vbInformation, "Icon design"
    Exit Sub
NotExist:
    MsgBox "The specified application doesn't exist", vbCritical, "Icon designer"
End Sub

Sub aaDevDeployIcon()
    Slide29.Select
    If Slide24.Shapes("ButtonSample").Visible = msoTrue Then
        MsgBox "Not designing an icon", vbCritical, "Icon design"
        Exit Sub
    End If
    SetVar "Macro", "zzDeployIcon"
    With SelectApp
        Dim I As Integer
        For I = Slide2.Shapes.Count To 1 Step -1
            If InStr(Slide2.Shapes(I).Name, "App") Then
                .AppList.AddItem (Replace(Slide2.Shapes(I).Name, "App", ""))
            End If
        Next I
        .Show
    End With
End Sub

Sub zzDeployIcon(Name As String)
    On Error GoTo DeployErr
    Dim Sld As Slide
    Set Sld = Slide24
    Dim Shp2 As Shape
    Dim Shapes As String
    Shapes = ""
    ShapeName = "App" & Name & ":Icon"
    For Each Shp2 In Sld.Shapes()
        If Shp2.Visible = msoTrue And Shp2.Name <> "DesignSlideLabel" Then
            If InStr(Shp2.Name, "AXTextBox") Then ApplyTbAttribs Shp2
            
            Shapes = Shapes & Shp2.Name & ","
        End If
    Next Shp2
    SplitShapes = Split(Shapes, ",")
    UJ = CInt(UBound(SplitShapes))
    Dim ShapesX() As String
    
    ReDim ShapesX(UJ)
    For I = CInt(UBound(SplitShapes)) To 0 Step -1
        CShape = SplitShapes(I)
        ShapesX(I) = SplitShapes(I)
    Next
    With Sld.Shapes.Range(ShapesX).Group
        .Name = ShapeName
    End With
    
    Slide24.Shapes("App" & Name & ":Icon").Copy
    Slide25.Shapes("App" & Name & ":Icon").Delete
    With Slide25.Shapes.Paste
        .Visible = msoFalse
        For J = 1 To .GroupItems.Count Step 1
            .GroupItems(J).Visible = msoTrue
            .GroupItems(J).ActionSettings(ppMouseClick).Run = "App" & Name
        Next J
    End With
    Slide24.Shapes("App" & Name & ":Icon").Delete
    For I = Slide24.Shapes.Count To 1 Step -1
        Slide24.Shapes(I).Visible = msoTrue
    Next I
    MsgBox "Icon deployed successfully", vbInformation, "Icon design"
    Exit Sub
DeployErr:
    MsgBox "Something went wrong while trying to deploy the icon. Please try again!", vbCritical, "Icon design"
End Sub

Sub zzRemoveApp(Name As String)
    Dim vbComp As Object
    Dim deleted As Boolean
    UnsetVar "Macro"
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
    Dim I As Integer
    For I = Slide1.Shapes.Count To 1 Step -1
        If InStr(1, Slide1.Shapes(I).Name, "FileLabel") = 1 Then
            Slide1.Shapes(I).Delete
        ElseIf InStr(1, Slide1.Shapes(I).Name, "FileIcon") = 1 Then
            Slide1.Shapes(I).Delete
        End If
    Next I
End Sub

Sub aaDevAddMacroActions()
    Dim inputShape As String
    inputShape = InputBox("Please enter the name of the App you want to add the action to (e.g. Calc)")
    For Each Shp In Slide25.Shapes("App" & inputShape & ":Icon").GroupItems
         With Shp.ActionSettings(ppMouseClick)
            .Run = "App" & inputShape
         End With
    Next Shp
    MsgBox "Macro added to App" & inputShape & ":Icon group", vbInformation, "Action automation script"
End Sub

Sub aaDevListApps()
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

