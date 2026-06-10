Option Explicit

Public DuzenleSatir As Long

Private Sub UserForm_Initialize()
    FormStil.Uygula Me
    UI_PersonelComboYukle cboPersonel
    cboKesintiSekli.Clear
    cboKesintiSekli.AddItem "Oran"
    cboKesintiSekli.AddItem "Sabit Tutar"
    cboKesintiSekli.ListIndex = 0
    txtDosyaTarihi.Value = Format$(Date, "dd.mm.yyyy")
    txtPay.Value = "1"
    txtPayda.Value = "4"

    If DuzenleSatir > 0 Then
        DuzenlemeyiYukle
    ElseIf Len(UI_SeciliPersonel) > 0 Then
        cboPersonel.Value = UI_SeciliPersonel
        UI_SeciliPersonel = ""
    End If
    KesintiSekliniAyarla
End Sub

Private Sub DuzenlemeyiYukle()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("İCRA DOSYALARI")
    lblTitle.Caption = "İCRA DOSYASI  -  DÜZENLE"
    cboPersonel.Value = ws.Cells(DuzenleSatir, 1).Text
    cboPersonel.Enabled = False
    txtDosyaTarihi.Value = ws.Cells(DuzenleSatir, 4).Text
    txtDosyaNo.Value = ws.Cells(DuzenleSatir, 5).Text
    txtOncelik.Value = ws.Cells(DuzenleSatir, 6).Text
    txtDaire.Value = ws.Cells(DuzenleSatir, 7).Text
    txtAlacakli.Value = ws.Cells(DuzenleSatir, 8).Text
    txtIcraTutar.Value = ws.Cells(DuzenleSatir, 9).Text
    cboKesintiSekli.Value = IIf(StrComp(Trim$(CStr(ws.Cells(DuzenleSatir, 10).Value)), "Oran", vbTextCompare) = 0, "Oran", "Sabit Tutar")
    txtPay.Value = ws.Cells(DuzenleSatir, 11).Text
    txtPayda.Value = ws.Cells(DuzenleSatir, 12).Text
    txtSabit.Value = ws.Cells(DuzenleSatir, 13).Text
    txtOncekiTahsil.Value = ws.Cells(DuzenleSatir, 15).Text
    txtIban.Value = ws.Cells(DuzenleSatir, 20).Text
    txtAciklama.Value = ws.Cells(DuzenleSatir, 21).Text
End Sub

Private Sub txtDosyaTarihi_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim secilen As Variant
    If IsDate(txtDosyaTarihi.Value) Then
        secilen = frmTarihSec.TarihSec(CDate(txtDosyaTarihi.Value))
    Else
        secilen = frmTarihSec.TarihSec(Date)
    End If
    If IsDate(secilen) Then txtDosyaTarihi.Value = Format$(CDate(secilen), "dd.mm.yyyy")
End Sub

Private Sub cboPersonel_Change()
    lblPersonelBilgi.Caption = UI_PersonelBilgi(cboPersonel.Value)
    If DuzenleSatir = 0 And Len(cboPersonel.Value) > 0 Then
        txtOncelik.Value = CStr(UI_SonrakiOncelik(cboPersonel.Value))
    End If
End Sub

Private Sub cboKesintiSekli_Change()
    KesintiSekliniAyarla
End Sub

Private Sub KesintiSekliniAyarla()
    Dim oranMi As Boolean
    oranMi = (StrComp(cboKesintiSekli.Value, "Oran", vbTextCompare) = 0)
    txtPay.Enabled = oranMi
    txtPayda.Enabled = oranMi
    txtSabit.Enabled = Not oranMi
    If oranMi Then
        If Len(txtPay.Value) = 0 Then txtPay.Value = "1"
        If Len(txtPayda.Value) = 0 Then txtPayda.Value = "4"
        txtSabit.Value = ""
    Else
        txtPay.Value = ""
        txtPayda.Value = ""
    End If
End Sub

Private Sub cmdKaydet_Click()
    Dim sonuc As String
    If DuzenleSatir > 0 Then
        sonuc = UI_IcraGuncelle(DuzenleSatir, cboPersonel.Value, txtDosyaTarihi.Value, _
                                txtDosyaNo.Value, txtOncelik.Value, txtDaire.Value, _
                                txtAlacakli.Value, txtIcraTutar.Value, cboKesintiSekli.Value, _
                                txtPay.Value, txtPayda.Value, txtSabit.Value, _
                                txtOncekiTahsil.Value, txtIban.Value, txtAciklama.Value)
    Else
        sonuc = UI_IcraKaydet(cboPersonel.Value, txtDosyaTarihi.Value, txtDosyaNo.Value, _
                              txtOncelik.Value, txtDaire.Value, txtAlacakli.Value, _
                              txtIcraTutar.Value, cboKesintiSekli.Value, txtPay.Value, _
                              txtPayda.Value, txtSabit.Value, txtOncekiTahsil.Value, _
                              txtIban.Value, txtAciklama.Value)
    End If
    If Left$(sonuc, 3) = "OK|" Then
        MsgBox Mid$(sonuc, 4), vbInformation, "İcra dosyası"
        Unload Me
    Else
        MsgBox sonuc, vbExclamation, "Eksik veya hatalı bilgi"
    End If
End Sub

Private Sub cmdTemizle_Click()
    If DuzenleSatir > 0 Then
        DuzenlemeyiYukle
    Else
        txtDosyaNo.Value = ""
        txtDaire.Value = ""
        txtAlacakli.Value = ""
        txtIcraTutar.Value = ""
        txtOncekiTahsil.Value = ""
        txtIban.Value = ""
        txtAciklama.Value = ""
        txtDosyaTarihi.Value = Format$(Date, "dd.mm.yyyy")
        If Len(cboPersonel.Value) > 0 Then txtOncelik.Value = CStr(UI_SonrakiOncelik(cboPersonel.Value))
    End If
End Sub

Private Sub cmdKapat_Click()
    Unload Me
End Sub
