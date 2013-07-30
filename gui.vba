Private Sub CommandButton1_Click()
'    process (GREECE)
    process (3)
    'SaveTurkey.Enabled = True
End Sub

Private Sub CommandButton2_Click()
'    process (TURKEY)
    process (1)
    'SaveTurkey.Enabled = True
End Sub

Private Sub CommandButton3_Click()
'    process (ITALY)
    process (2)
End Sub

Private Sub CommandButton4_Click()
    ClearStatementData
    'SaveTurkey.Enabled = False
End Sub

Private Sub SaveTurkey_Click()
    SaveTurkeySub
End Sub

