Private Sub CommandButton1_Click()
    SetVar "Name", AppNameTb.Value
    SetVar "FriendlyName", FriendlyNameTb.Value
    If PublicOption.Value = True Then
        SetVar "Access", "Everyone"
    Else
        SetVar "Access", "Administrators"
    End If
    If GenModuleCheckbox.Value = True Then
        SetVar "GenModule", "True"
    Else
        SetVar "GenModule", "False"
    End If
    If GenCodeCheckbox.Value = True Then
        SetVar "GenCode", "True"
    Else
        SetVar "GenCode", "False"
    End If
    If CreateShortcutsCheck.Value = True Then
        SetVar "Shortcuts", "True"
    Else
        SetVar "Shortcuts", "False"
    End If
    zzCreateApp
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub