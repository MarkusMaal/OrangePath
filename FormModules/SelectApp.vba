Private Sub CommandButton1_Click()
    Application.Run CheckVars("%Macro%"), AppList.Value
    Unload Me
End Sub