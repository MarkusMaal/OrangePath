Private Sub BootText_Click()
    If Slide7.BootText.Caption = "It's now safe to close the presentation" Then
        ActivePresentation.SlideShowWindow.View.Exit
    End If
End Sub