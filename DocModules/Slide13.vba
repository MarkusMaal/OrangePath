Private Sub AxTextBox_Change()
    Dim SubShp As Shape
    For Each Shp In Slide13.Shapes
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
    If KeyCode.Value = 13 Then
        SetVar "InputValue", AxTextBox.Text
        Slide13.Shapes("RegularApp:" & Slide1.Shapes("AppID").TextFrame.TextRange.Text).Delete
        If CheckVars("%Macro%") <> "" And CheckVars("%Macro%") <> "%Macro%" Then
            Application.Run CheckVars("%Macro%"), Shp
        End If
        UnsetVar "Macro"
    End If
End Sub

Private Sub PasswordField_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub UsernameFIeld_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
End Sub
