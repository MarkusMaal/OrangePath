Private Sub CommandButton1_Click()
    SetVar "Name", AppNameTb.Value
    SetVar "FriendlyName", FriendlyNameTb.Value
    If PublicOption.Value = True Then
        SetVar "Access", "Everyone"
    Else
        SetVar "Access", "Administrators"
    End If
    If GenModuleCheckBox.Value = True Then
        SetVar "GenModule", "True"
    Else
        SetVar "GenModule", "False"
    End If
    If GenCodeCheckBox.Value = True Then
        SetVar "GenCode", "True"
    Else
        SetVar "GenCode", "False"
    End If
    zzCreateApp
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub