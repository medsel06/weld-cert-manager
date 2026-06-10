Option Explicit

Private Sub UserForm_Initialize()
    FormStil.Uygula Me
End Sub

Private Sub cmdPersonel_Click()
    frmPersonelListe.Show vbModal
End Sub

Private Sub cmdDonem_Click()
    frmDonem.Show vbModal
End Sub

Private Sub cmdOzet_Click()
    frmPersonelOzet.Show vbModal
End Sub

Private Sub cmdCikis_Click()
    If MsgBox("Kaydedilip çıkılsın mı?", vbYesNo + vbQuestion, "Çıkış") = vbYes Then
        UYGULAMA_CIKIS
    End If
End Sub
