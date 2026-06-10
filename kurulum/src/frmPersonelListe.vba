Option Explicit

Private seciliSatir As Long

Private Sub UserForm_Initialize()
    FormStil.Uygula Me
    ListeyiYukle ""
    Temizle
End Sub

Private Sub txtAra_Change()
    ListeyiYukle txtAra.Value
End Sub

Private Sub ListeyiYukle(ByVal filtre As String)
    Dim ws As Worksheet
    Dim r As Long
    Dim ix As Long
    Dim ad As String
    filtre = LCase$(Trim$(filtre))
    Set ws = ThisWorkbook.Worksheets("PERSONEL")

    lstPersonel.Clear
    For r = 2 To 300
        ad = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(ad) > 0 Then
            If Len(filtre) = 0 Or InStr(1, LCase$(ad & " " & ws.Cells(r, 2).Text), filtre, vbTextCompare) > 0 Then
                lstPersonel.AddItem ad
                ix = lstPersonel.ListCount - 1
                lstPersonel.List(ix, 1) = ws.Cells(r, 2).Text
                lstPersonel.List(ix, 2) = Format$(ws.Cells(r, 3).Value, "#,##0.00")
                lstPersonel.List(ix, 3) = ws.Cells(r, 4).Text
                lstPersonel.List(ix, 4) = ws.Cells(r, 5).Text
                lstPersonel.List(ix, 5) = CStr(r)
            End If
        End If
    Next r
    lblSayac.Caption = lstPersonel.ListCount & " personel"
End Sub

Private Sub lstPersonel_Click()
    If lstPersonel.ListIndex < 0 Then Exit Sub
    seciliSatir = CLng(lstPersonel.List(lstPersonel.ListIndex, 5))
    txtAd.Value = lstPersonel.List(lstPersonel.ListIndex, 0)
    txtTC.Value = lstPersonel.List(lstPersonel.ListIndex, 1)
    txtNetMaas.Value = lstPersonel.List(lstPersonel.ListIndex, 2)
    txtIseGiris.Value = lstPersonel.List(lstPersonel.ListIndex, 3)
    txtAciklama.Value = lstPersonel.List(lstPersonel.ListIndex, 4)
End Sub

Private Sub lstPersonel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    KartAc
End Sub

Private Sub cmdKart_Click()
    KartAc
End Sub

Private Sub KartAc()
    Dim ad As String
    If lstPersonel.ListIndex >= 0 Then
        ad = lstPersonel.List(lstPersonel.ListIndex, 0)
    Else
        ad = Trim$(txtAd.Value)
    End If
    If Len(ad) = 0 Then
        MsgBox "Önce listeden personel seçin.", vbExclamation, "İcra kartı"
        Exit Sub
    End If
    frmPersonelKart.PersonelAd = ad
    frmPersonelKart.Show vbModal
    ListeyiYukle txtAra.Value
End Sub

Private Sub txtIseGiris_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim secilen As Variant
    If IsDate(txtIseGiris.Value) Then
        secilen = frmTarihSec.TarihSec(CDate(txtIseGiris.Value))
    Else
        secilen = frmTarihSec.TarihSec(Date)
    End If
    If IsDate(secilen) Then txtIseGiris.Value = Format$(CDate(secilen), "dd.mm.yyyy")
End Sub

Private Sub cmdYeni_Click()
    Temizle
End Sub

Private Sub cmdKaydet_Click()
    Dim sonuc As String
    If seciliSatir = 0 Then
        sonuc = UI_PersonelKaydet(txtAd.Value, txtTC.Value, txtNetMaas.Value, txtIseGiris.Value, txtAciklama.Value)
    Else
        sonuc = UI_PersonelGuncelle(seciliSatir, txtAd.Value, txtTC.Value, txtNetMaas.Value, txtIseGiris.Value, txtAciklama.Value)
    End If
    If Left$(sonuc, 3) = "OK|" Then
        MsgBox Mid$(sonuc, 4), vbInformation, "Personel"
        ListeyiYukle txtAra.Value
        Temizle
    Else
        MsgBox sonuc, vbExclamation, "Personel"
    End If
End Sub

Private Sub cmdSil_Click()
    Dim sonuc As String
    If seciliSatir = 0 Then
        MsgBox "Silmek için listeden personel seçin.", vbExclamation, "Personel"
        Exit Sub
    End If
    sonuc = UI_PersonelSil(seciliSatir)
    If Left$(sonuc, 3) = "OK|" Then
        MsgBox Mid$(sonuc, 4), vbInformation, "Personel"
        ListeyiYukle txtAra.Value
        Temizle
    Else
        MsgBox sonuc, vbExclamation, "Personel"
    End If
End Sub

Private Sub cmdKapat_Click()
    Unload Me
End Sub

Private Sub Temizle()
    seciliSatir = 0
    txtAd.Value = ""
    txtTC.Value = ""
    txtNetMaas.Value = ""
    txtIseGiris.Value = Format$(Date, "dd.mm.yyyy")
    txtAciklama.Value = ""
    If lstPersonel.ListIndex >= 0 Then lstPersonel.ListIndex = -1
    txtAd.SetFocus
End Sub
