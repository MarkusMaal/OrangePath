Private Sub AxComboBox_Change()
    If CheckVars("%Macro%") <> "%Macro%" Then
        Application.Run CheckVars("%Macro%")
    End If
End Sub

Private Sub AxTextBox_Change()
    Dim SubShp As Shape
    For Each Shp In Slide1.Shapes
        If Shp.Type = msoGroup Then
            For Each SubShp In Shp.GroupItems
                If SubShp.Left = AxTextBox.Left And SubShp.Top = AxTextBox.Top And SubShp.Name <> "AxTextBox" Then
                    SetTextBoxVal SubShp
                End If
            Next SubShp
        End If
    Next Shp
End Sub

Private Sub AxTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If AxTextBox.ForeColor = RGB(254, 254, 254) Then
        On Error GoTo Crash
        If KeyCode.Value = 13 Then
            Dim AppID As String
            AppID = "-1"
            For Each Shp In Slide1.Shapes
                If Shp.Type = msoGroup Then
                    For Each SubShp In Shp.GroupItems
                        If SubShp.Left = AxTextBox.Left And SubShp.Top = AxTextBox.Top And SubShp.Name <> "AxTextBox" Then
                            SplitZ = Split(SubShp.Name, ":")
                            AppID = SplitZ(1)
                        End If
                    Next SubShp
                End If
            Next Shp
            If AppID = "-1" Then Exit Sub
            If InStr(1, AxTextBox.Text, "launch ") = 1 Then
                AppName = Replace(AxTextBox.Text, "launch ", "")
                If InStr(1, AppName, "InputBox") = 1 Then
                    Args = Split(Replace(AppName, "InputBox ", ""), " ")
                    AppInputBox CheckVars(Args(0)), CheckVars(Args(1)), True
                ElseIf InStr(1, AppName, "Message") = 1 Then
                    Args = Split(Replace(AppName, "Message ", ""), " ")
                    AppMessage CheckVars(Args(0)), CheckVars(Args(1)), CheckVars(Args(2)), True
                ElseIf AppName = "Taskmgr" Then
                    AppTaskmgr Slide1.Shapes("RegularApp:" & AppID)
                ElseIf AppName = "Guess" Then
                    AppGuess Slide1.Shapes("RegularApp:" & AppID)
                ElseIf AppName = "Settings" Then
                    AppSettings Slide1.Shapes("RegularApp:" & AppID)
                Else
                    Slide1.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = AppName
                    ActivePresentation.SlideShowWindow.View.GotoSlide 4
                    CreateNewWindow
                End If
            ElseIf AxTextBox.Text = "clear" Then
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = ""
            ElseIf AxTextBox.Text = "fullscreen" Then
                Slide1.Shapes("WindowFrameAppShell:" & AppID).Delete
                Slide1.Shapes("WindowTitleAppShell:" & AppID).Delete
                Slide1.Shapes("HandleAppShell:" & AppID).Delete
                Slide1.Shapes("CloseAppShell:" & AppID).Delete
                Slide1.Shapes("MinimizeAppShell:" & AppID).Delete
                Slide1.Shapes("OutputAppShell:" & AppID).ActionSettings(ppMouseClick).Run = "MinimizeRestore"
                With Slide1.Shapes("RegularApp:" & AppID)
                    .Left = 0
                    .Top = 0
                    .Width = ActivePresentation.PageSetup.SlideWidth
                    .Height = ActivePresentation.PageSetup.SlideHeight
                    AppShellSizeChanged AppID, True
                End With
            ElseIf InStr(1, AxTextBox.Text, "print ") = 1 Then
                Message = CheckVars(Replace(AxTextBox.Text, "print ", ""))
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & Message
            ElseIf AxTextBox.Text = "applist" Then
                Dim HasNewLine As Boolean
                HasNewLine = False
                For Each Shp In Slide2.Shapes
                    If Shp.Type = msoGroup Then
                        If Not HasNewLine Then
                            Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & Replace(Shp.Name, "App", "") & " "
                            HasNewLine = True
                        Else
                            Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & Replace(Shp.Name, "App", "") & " "
                        End If
                    End If
                Next Shp
            ElseIf AxTextBox.Text = "proclist" Then
                For Each Shp In Slide1.Shapes
                    If Shp.Type = msoGroup Then
                        If InStr(Shp.Name, ":") And InStr(1, Shp.Name, "ITaskIcon:") <> 1 Then
                            SplitZ = Split(Shp.Name, ":")
                            AID = SplitZ(1)
                            AppNameSplit = Split(Slide1.Shapes("RegularApp:" & AID).GroupItems(1).Name, ":")
                            AppNameSplit2 = Split(AppNameSplit(0), "App")
                            AppName = AppNameSplit2(1)
                            Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & AID & ": " & AppName
                        End If
                    End If
                Next Shp
            ElseIf InStr(1, AxTextBox.Text, "killapp ") = 1 Then
                AppID1 = Replace(AxTextBox.Text, "killapp ", "")
                Result = 1
                For Each Shp In Slide1.Shapes
                    If Shp.Type = msoGroup Then
                        If Shp.Name = "RegularApp:" & AppID1 Then
                            Result = 0
                        End If
                    End If
                Next Shp
                If Result = 1 Then
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Process not found"
                ElseIf Result = 0 Then
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully"
                    CloseWindow Slide1.Shapes("RegularApp:" & AppID1).GroupItems(1)
                End If
            ElseIf InStr(1, AxTextBox.Text, "color ") = 1 Then
                ColorID = Replace(AxTextBox.Text, "color ", "")
                If Len(ColorID) = 1 Then
                    ColorID = "0" & ColorID
                End If
                BG = UCase(Left(ColorID, 1))
                FG = UCase(Right(ColorID, 1))
                BRGB = RGB(0, 0, 0)
                FRGB = RGB(192, 192, 192)
                If BG = "1" Then BRGB = RGB(0, 0, 128)
                If BG = "2" Then BRGB = RGB(0, 128, 0)
                If BG = "3" Then BRGB = RGB(0, 128, 128)
                If BG = "4" Then BRGB = RGB(128, 0, 0)
                If BG = "5" Then BRGB = RGB(128, 0, 128)
                If BG = "6" Then BRGB = RGB(128, 128, 0)
                If BG = "7" Then BRGB = RGB(192, 192, 192)
                If BG = "8" Then BRGB = RGB(128, 128, 128)
                If BG = "9" Then BRGB = RGB(0, 0, 256)
                If BG = "A" Then BRGB = RGB(0, 255, 0)
                If BG = "B" Then BRGB = RGB(0, 255, 255)
                If BG = "C" Then BRGB = RGB(255, 0, 0)
                If BG = "D" Then BRGB = RGB(255, 0, 255)
                If BG = "E" Then BRGB = RGB(255, 255, 0)
                If BG = "F" Then BRGB = RGB(255, 255, 255)
                If FG = "0" Then FRGB = RGB(0, 0, 0)
                If FG = "1" Then FRGB = RGB(0, 0, 128)
                If FG = "2" Then FRGB = RGB(0, 128, 0)
                If FG = "3" Then FRGB = RGB(0, 128, 128)
                If FG = "4" Then FRGB = RGB(128, 0, 0)
                If FG = "5" Then FRGB = RGB(128, 0, 128)
                If FG = "6" Then FRGB = RGB(128, 128, 0)
                If FG = "8" Then FRGB = RGB(128, 128, 128)
                If FG = "9" Then FRGB = RGB(0, 0, 256)
                If FG = "A" Then FRGB = RGB(0, 255, 0)
                If FG = "B" Then FRGB = RGB(0, 255, 255)
                If FG = "C" Then FRGB = RGB(255, 0, 0)
                If FG = "D" Then FRGB = RGB(255, 0, 255)
                If FG = "E" Then FRGB = RGB(255, 255, 0)
                If FG = "F" Then FRGB = RGB(255, 255, 255)
                Slide1.Shapes("OutputAppShell:" & AppID).Fill.ForeColor.RGB = BRGB
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Font.Color.RGB = FRGB
            ElseIf AxTextBox.Text = "whoami" Then
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & Slide1.Shapes("Username").TextFrame.TextRange.Text
            ElseIf AxTextBox.Text = "time" Then
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & Time
            ElseIf AxTextBox.Text = "exit" Then
                Slide1.Shapes("RegularApp:" & AppID).Delete
                If AAX Then Slide1.AxTextBox.Visible = False
                Dim TaskIcon As Integer
                TaskIcon = CInt(Slide1.Shapes("TaskIcon:" & AppID).Left)
                Slide1.Shapes("TaskIcon:" & AppID).Delete
                Slide1.Shapes("ITaskIcon:" & AppID).Delete
                If TaskIcon = 373 Then
                    TaskIcon = 3
                ElseIf TaskIcon = 527 Then
                    TaskIcon = 4
                ElseIf TaskIcon = 219 Then
                    TaskIcon = 2
                ElseIf TaskIcon = 65 Then
                    TaskIcon = 1
                Else
                    TaskIcon = 5
                End If
                ReorganizeTaskIcons TaskIcon
                Exit Sub
            ElseIf AxTextBox.Text = "pm install" Then
                UpdateTest
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Package install OK"
                Exit Sub
            ElseIf InStr(1, AxTextBox.Text, "pm uninstall ") = 1 Then
                Critical = Array("calc", "1", "3d", "menu", "taskmgr", "shell", "message", "settings", "notes", "guess", "videoplayer", "paint", "gallery", "components", "help", "modalcolorpicker", "modalfiles", "modalhelpview", "soundplayer", "words")
                Package = Replace(AxTextBox.Text, "pm uninstall ", "")
                For Each CriticalPackage In Critical
                    If LCase(Package) = CriticalPackage And CheckVars("%override%") <> "I realize that by doing this, I might permanently destroy Sunlight OS" Then
                        Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "This is a system component, which cannot be uninstalled. To override this failsafe, set the override variable to equal 'I realize that by doing this, I might permanently destroy Sunlight OS'"
                        Exit Sub
                    End If
                Next CriticalPackage
                'Delete window
                Slide2.Shapes("App" & Package).Delete
                'Delete shortcut icons
                Slide25.Shapes("App" & Package & ":Icon").Delete
                Slide25.Shapes("App" & Package & ":Properties").Delete
                'Delete VBA module (this is why we don't want to mess with system components)
                ActivePresentation.VBProject.VBComponents.Remove ActivePresentation.VBProject.VBComponents("MApp" & Package)
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Package uninstall OK"
            ElseIf AxTextBox.Text = "reboot" Then
                Restart
                Exit Sub
            ElseIf AxTextBox.Text = "reboot recovery" Then
                RestartRecovery
                Exit Sub
            ElseIf AxTextBox.Text = "hibernate" Then
                Hibernate
                Exit Sub
            ElseIf AxTextBox.Text = "shutdown" Then
                ActivePresentation.SlideShowWindow.View.GotoSlide 5
            ElseIf AxTextBox.Text = "setfactoryconfig" Then
                HardReset
                Restart
                Exit Sub
            ElseIf InStr(1, AxTextBox.Text, "dir ") = 1 Then
                Directory = Right(AxTextBox.Text, Len(AxTextBox.Text) - 4)
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "-------------------------------" & vbNewLine & "Directory listing of " & Directory & vbNewLine & "-------------------------------" & vbNewLine & GetFiles(Directory) & vbNewLine
                Slide1.AxTextBox.Text = ""
                Exit Sub
            ElseIf InStr(1, AxTextBox.Text, "type ") = 1 Then
                Filename = Right(AxTextBox.Text, Len(AxTextBox.Text) - 5)
                Text = GetFileContent(Filename)
                If Text <> "*" Then
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & Text & vbNewLine
                Else
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Access denied" & vbNewLine
                End If
                Slide1.AxTextBox.Text = ""
                Exit Sub
            ElseIf InStr(1, AxTextBox.Text, "deltree ") = 1 Then
                Directory = Right(AxTextBox.Text, Len(AxTextBox.Text) - 8)
                DeleteDir Directory
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully" & vbNewLine
                Slide1.AxTextBox.Text = ""
                Exit Sub
            ElseIf InStr(1, AxTextBox.Text, "del ") = 1 Then
                Filename = Right(AxTextBox.Text, Len(AxTextBox.Text) - 4)
                DeleteFile Filename
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully" & vbNewLine
                Slide1.AxTextBox.Text = ""
                Exit Sub
            ElseIf InStr(1, AxTextBox.Text, "open ") = 1 Then
                Filename = Right(AxTextBox.Text, Len(AxTextBox.Text) - 5)
                With Slide1.Shapes.AddShape(msoShapeRectangle, 0, 0, 0, 0)
                    .Name = "PathAppFiles:" & AppID
                    .TextFrame.TextRange.Text = Filename
                    .Visible = msoFalse
                End With
                With Slide1.Shapes.AddShape(msoShapeRectangle, 0, 0, 0, 0)
                    .Name = "SelectModeCheckAppFiles:" & AppID
                    .Fill.ForeColor.RGB = Slide1.ColorScheme.Colors(ppBackground).RGB
                    .Visible = msoFalse
                End With
                Dim Assoc As String
                Assoc = GetAssoc(Filename, AppID)
                If Assoc <> "" Then
                    Application.Run "Assoc" & Assoc, Slide1.Shapes("RegularApp:" & AppID).GroupItems(1)
                Else
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "There are no associations for this file format." & vbNewLine
                End If
                Slide1.Shapes("PathAppFiles:" & AppID).Delete
                Slide1.Shapes("SelectModeCheckAppFiles:" & AppID).Delete
                Slide1.AxTextBox.Text = ""
                Exit Sub
            ElseIf AxTextBox.Text = "logout" Then
                Logout
                Exit Sub
            ElseIf InStr(1, AxTextBox.Text, "login ") = 1 Then
                UP = Split(AxTextBox.Text, " ")
                Slide13.UsernameFIeld.Text = UP(1)
                Slide13.PasswordField.Text = UP(2)
                Login
                Exit Sub
            ElseIf InStr(1, AxTextBox.Text, "getbootscreen ") = 1 Then
                Args = Split(AxTextBox.Text, " ")
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
                    With Slide1.Shapes("Shape2App3D:" & PID)
                        .ThreeD.RotationZ = Slide5.Shapes("Bootlogo").ThreeD.RotationZ
                        .ThreeD.RotationY = Slide5.Shapes("Bootlogo").ThreeD.RotationY
                        .ThreeD.RotationX = Slide5.Shapes("Bootlogo").ThreeD.RotationX
                    End With
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully"
                    Exit Sub
                ElseIf Tp = "1" Then
                    With Slide1.Shapes("Shape2App3D:" & PID)
                        .ThreeD.RotationZ = Slide3.Shapes("Bootlogo").ThreeD.RotationZ
                        .ThreeD.RotationY = Slide3.Shapes("Bootlogo").ThreeD.RotationY
                        .ThreeD.RotationX = Slide3.Shapes("Bootlogo").ThreeD.RotationX
                    End With
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully"
                    Exit Sub
                ElseIf Tp = "2" Then
                    With Slide1.Shapes("Shape2App3D:" & PID)
                        .ThreeD.RotationZ = Slide2.Shapes("Bootlogo").ThreeD.RotationZ
                        .ThreeD.RotationY = Slide2.Shapes("Bootlogo").ThreeD.RotationY
                        .ThreeD.RotationX = Slide2.Shapes("Bootlogo").ThreeD.RotationX
                    End With
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully"
                    Exit Sub
                End If
            ElseIf InStr(1, AxTextBox.Text, "setbootscreen ") = 1 Then
                Dim AxTxt As String
                AxTxt = AxTextBox.Text
                Args = Split(AxTxt, " ")
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
                    Slide7.Shapes("Bootlogo").ThreeD.RotationZ = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationZ
                    Slide7.Shapes("Bootlogo").ThreeD.RotationY = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationY
                    Slide7.Shapes("Bootlogo").ThreeD.RotationX = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationX
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully"
                    Exit Sub
                ElseIf Tp = "2" Then
                    Slide2.Shapes("Bootlogo").ThreeD.RotationZ = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationZ
                    Slide2.Shapes("Bootlogo").ThreeD.RotationY = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationY
                    Slide2.Shapes("Bootlogo").ThreeD.RotationX = Slide1.Shapes("Shape2App3D:" & PID).ThreeD.RotationX
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully"
                    Exit Sub
                End If
            ElseIf InStr(1, AxTextBox.Text, "getconfig ") = 1 Then
                Configs = Split(AxTextBox.Text, " ")
                Key = Configs(1)
                Value = GetFileContent("/System/Settings.cnf", Key)
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & Value
            ElseIf AxTextBox.Text = "set" Then
                Dim Shp2 As Shape
                For Each Shp2 In Slide21.Shapes
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & Shp2.Name & "=" & Shp2.TextFrame.TextRange.Text
                Next Shp2
                AxTextBox.Text = ""
                Exit Sub
            ElseIf InStr(1, AxTextBox.Text, "unset ") = 1 Then
                Key = Replace(AxTextBox.Text, "unset ", "")
                UnsetVar Key
                Exit Sub
            ElseIf InStr(1, AxTextBox.Text, "set ") = 1 Then
                ASplit = Split(Replace(AxTextBox.Text, "set ", ""), "=")
                Key = ASplit(0)
                Value = ASplit(1)
                SetVar Key, Value
            ElseIf AxTextBox.Text = "crash" Then
                OSCrash "MANUALLY_INITIATED_CRASH", Err
            ElseIf InStr(1, AxTextBox.Text, "setuconfig ") = 1 Then
                Configs = Split(Replace(AxTextBox.Text, "setuconfig ", ""), "=")
                Key = Configs(0)
                Value = Configs(1)
                Username = Slide1.Shapes("Username").TextFrame.TextRange.Text
                SetFileContent "/Users/" & Username & "/" & Key & ".txt", Value
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully."
            ElseIf InStr(1, AxTextBox.Text, "setconfig ") = 1 Then
                Configs = Split(Replace(AxTextBox.Text, "setconfig ", ""), "=")
                Key = Configs(0)
                Value = Configs(1)
                SetFileContent "/System/Settings.cnf", Value, Key
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully."
            ElseIf InStr(1, AxTextBox.Text, "delconfig ") = 1 Then
                Configs = Split(AxTextBox.Text, " ")
                Key = Configs(1)
                If FileExists("/System/Settings.cnf", Key) Then
                    DeleteFile "/System/Settings.cnf", Key
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully."
                Else
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Key not found."
                End If
            ElseIf InStr(1, AxTextBox.Text, "deluconfig ") = 1 Then
                Configs = Split(AxTextBox.Text, " ")
                Key = Configs(1)
                Username = Slide1.Shapes("Username").TextFrame.TextRange.Text
                If FileExists("/Users/" & Username & "/" & Key & ".txt") Then
                    DeleteFile "/Users/" & Username & "/" & Key & ".txt"
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully."
                Else
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Key not found for this user."
                End If
            ElseIf InStr(1, AxTextBox.Text, "sleep ") = 1 Then
                Args = Split(AxTextBox.Text, " ")
                SlpTime = CInt(Args(1))
                Pause (SlpTime)
                Exit Sub
            ElseIf InStr(1, AxTextBox.Text, "setbootdelay ") = 1 Then
                Args = Split(AxTextBox.Text, " ")
                SetBootDelay CInt(Args(1))
                AxTextBox.Text = ""
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command completed successfully."
                Exit Sub
            ElseIf AxTextBox.Text = "getbootdelay" Then
                AxTextBox.Text = ""
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Boot delay: " & GetBootDelay & " second(s)"
                Exit Sub
            ElseIf InStr(1, AxTextBox.Text, "getuconfig ") = 1 Then
                Configs = Split(AxTextBox.Text, " ")
                Key = Configs(1)
                Username = Slide1.Shapes("Username").TextFrame.TextRange.Text
                Value = GetFileContent("/Users/" & Username & "/" & Key & ".txt")
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & Value
            ElseIf InStr(1, AxTextBox.Text, "title ") = 1 Then
                Dim TitleText As String
                TitleText = CheckVars(Replace(AxTextBox.Text, "title ", ""))
                If ShapeExists(Slide1, "WindowTitleAppShell:" & AppID) Then
                    Slide1.Shapes("WindowTitleAppShell:" & AppID).TextFrame.TextRange.Text = TitleText
                End If
                Slide1.Shapes("TaskIcon:" & AppID).TextFrame.TextRange.Text = TitleText
                CompensateText Slide1.Shapes("TaskIcon:" & AppID), TitleText
            ElseIf AxTextBox.Text = "help" Then
                HelpMsg = "launch [AppName], launch Message [text] [title] [Info/Error/Exclamation], clear, applist, proclist, killapp [PID], print [message], title [message], color [00-FF], whoami, time, reboot [recovery], hibernate, shutdown, logout, exit, login [username] [password], getconfig [Key], getuconfig [Key], setconfig [Key]=[Value], setuconfig [Key]=[Value], setfactoryconfig, delconfig [key], deluconfig [key], sleep [n], set [key]=[value], unset [key], getbootscreen [0/1/2] [PID], setbootscreen [0/1/2] [PID], dir [directory], type [file path], del [filename], deltree [directory], open [filename], fullscreen, setbootdelay [secs], getbootdelay"
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & HelpMsg
            Else
                Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Bad command"
            End If
            AxTextBox.Text = ""
            If ShapeExists(Slide1, "OutputAppShell:" & AppID) Then
                If Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.BoundHeight > Slide1.Shapes("OutputAppShell:" & AppID).Height Then
                    TextSplit = Split(Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text, Chr(13))
                    FirstItem = TextSplit(1)
                    Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = ""
                    For I = 2 To UBound(TextSplit)
                        Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & Replace(Replace(TextSplit(I), Chr(10), ""), Chr(13), "")
                    Next I
                End If
            End If
Done:
            Exit Sub
Crash:
            Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text = Slide1.Shapes("OutputAppShell:" & AppID).TextFrame.TextRange.Text & vbNewLine & "Command failed. " & Err.Description
        End If
    End If
End Sub