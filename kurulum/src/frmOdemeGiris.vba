Option Explicit

Public OnPersonel As String
Public OnDosya As String

Private Sub UserForm_Initialize()
    FormStil.Uygula Me
    UI_PersonelComboYukle cboPersonel
    txtTarih.Value = Format$(Date, "dd.mm.yyyy")
    If Len(OnPersonel) > 0 Then
        cboPersonel.Value = OnPersonel
        If Len(OnDosya) > 0 Then cboDosya.Value = OnDosya
    End If
End Sub

Private Sub cboPersonel_Change()
    UI_DosyaComboYukle cboDosya, cboPersonel.Value
    BekleyeniGoster
End Sub

Private Sub cboDosya_Change()
    BekleyeniGoster
End Sub

Private Sub BekleyeniGoster()
    If Len(Trim$(cboPersonel.Value)) > 0 And Len(Trim$(cboDosya.Value)) > 0 Then
        lblBekleyenBilgi.Caption = "Bu dosyada bekleyen: " & _
            Format$(UI_DosyaBekleyen(cboPersonel.Value, cboDosya.Value), "#,##0.00") & " TL"
    Else
        lblBekleyenBilgi.Caption = ""
    End If
End Sub

Private Sub txtTarih_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim secilen As Variant
    If IsDate(txtTarih.Value) Then
        secilen = frmTarihSec.TarihSec(CDate(txtTarih.Value))
    Else
        secilen = frmTarihSec.TarihSec(Date)
    End If
    If IsDate(secilen) Then txtTarih.Value = Format$(CDate(secilen), "dd.mm.yyyy")
End Sub

Private Sub cmdKaydet_Click()
    Dim sonuc As String
    sonuc = UI_OdemeKaydet(cboPersonel.Value, cboDosya.Value, txtTarih.Value, _
                           txtTutar.Value, txtDekont.Value, txtAciklama.Value)
    If Left$(sonuc, 3) = "OK|" Then
        MsgBox Mid$(sonuc, 4), vbInformation, "Ödeme"
        Unload Me
    Else
        MsgBox sonuc, vbExclamation, "Eksik veya hatalı bilgi"
    End If
End Sub

Private Sub cmdKapat_Click()
    Unload Me
End Sub
