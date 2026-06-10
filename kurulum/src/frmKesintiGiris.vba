Option Explicit

Public OnPersonel As String
Public OnDosya As String

Private Sub UserForm_Initialize()
    FormStil.Uygula Me
    UI_PersonelComboYukle cboPersonel
    UI_AyComboYukle cboAy
    txtYil.Value = CStr(Year(Date))
    cboAy.Value = UI_AyAdi(Date)
    If Len(OnPersonel) > 0 Then
        cboPersonel.Value = OnPersonel
        If Len(OnDosya) > 0 Then cboDosya.Value = OnDosya
    End If
End Sub

Private Sub cboPersonel_Change()
    UI_DosyaComboYukle cboDosya, cboPersonel.Value
End Sub

Private Sub cmdKaydet_Click()
    Dim sonuc As String
    sonuc = UI_ManuelKesintiKaydet(cboPersonel.Value, cboDosya.Value, txtYil.Value, _
                                   cboAy.Value, txtTutar.Value, txtAciklama.Value)
    If Left$(sonuc, 3) = "OK|" Then
        MsgBox Mid$(sonuc, 4), vbInformation, "Manuel kesinti"
        Unload Me
    Else
        MsgBox sonuc, vbExclamation, "Eksik veya hatalı bilgi"
    End If
End Sub

Private Sub cmdKapat_Click()
    Unload Me
End Sub
