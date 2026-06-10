Option Explicit

Public PersonelAd As String

Private Sub UserForm_Initialize()
    FormStil.Uygula Me
    Yenile
End Sub

Public Sub Yenile()
    Application.Calculate
    KisiYukle
    DosyalariYukle
End Sub

Private Sub KisiYukle()
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Worksheets("PERSONEL")
    lblTitle.Caption = "İCRA KARTI  -  " & PersonelAd
    lblKisi.Caption = ""
    For r = 2 To 300
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), PersonelAd, vbTextCompare) = 0 Then
            lblKisi.Caption = "TC: " & ws.Cells(r, 2).Text & "    Net Maaş: " & _
                              Format$(Sayi(ws.Cells(r, 3).Value), "#,##0.00") & " TL"
            Exit For
        End If
    Next r
End Sub

Private Sub DosyalariYukle()
    Dim ws As Worksheet
    Dim r As Long
    Dim i As Long
    Dim j As Long
    Dim ix As Long
    Dim n As Long
    Dim sat(1 To 64) As Long
    Dim onc(1 To 64) As Double
    Dim seciliDosya As String
    Dim toplamKalan As Double

    If lstDosyalar.ListIndex >= 0 Then seciliDosya = lstDosyalar.List(lstDosyalar.ListIndex, 1)
    Set ws = ThisWorkbook.Worksheets("İCRA DOSYALARI")

    For r = 2 To 300
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), PersonelAd, vbTextCompare) = 0 Then
            If Len(Trim$(CStr(ws.Cells(r, 5).Value))) > 0 Then
                If n < 64 Then
                    n = n + 1
                    sat(n) = r
                    onc(n) = Sayi(ws.Cells(r, 6).Value)
                End If
            End If
        End If
    Next r

    Dim tL As Long
    Dim tD As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If onc(j) < onc(i) Then
                tL = sat(i): sat(i) = sat(j): sat(j) = tL
                tD = onc(i): onc(i) = onc(j): onc(j) = tD
            End If
        Next j
    Next i

    lstDosyalar.Clear
    For i = 1 To n
        r = sat(i)
        toplamKalan = toplamKalan + Sayi(ws.Cells(r, 18).Value)
        lstDosyalar.AddItem ws.Cells(r, 6).Text
        ix = lstDosyalar.ListCount - 1
        lstDosyalar.List(ix, 1) = ws.Cells(r, 5).Text
        lstDosyalar.List(ix, 2) = ws.Cells(r, 4).Text
        lstDosyalar.List(ix, 3) = ws.Cells(r, 8).Text
        lstDosyalar.List(ix, 4) = ws.Cells(r, 9).Text
        lstDosyalar.List(ix, 5) = ws.Cells(r, 18).Text
        lstDosyalar.List(ix, 6) = ws.Cells(r, 24).Text
        lstDosyalar.List(ix, 7) = ws.Cells(r, 19).Text
        lstDosyalar.List(ix, 8) = CStr(r)
    Next i
    lblToplamlar.Caption = n & " dosya    Toplam kalan: " & Format$(toplamKalan, "#,##0.00") & " TL"

    DetayTemizle
    If Len(seciliDosya) > 0 Then
        For i = 0 To lstDosyalar.ListCount - 1
            If StrComp(lstDosyalar.List(i, 1), seciliDosya, vbTextCompare) = 0 Then
                lstDosyalar.ListIndex = i
                Exit For
            End If
        Next i
    ElseIf lstDosyalar.ListCount > 0 Then
        lstDosyalar.ListIndex = 0
    End If
End Sub

Private Sub lstDosyalar_Click()
    If lstDosyalar.ListIndex < 0 Then Exit Sub
    DetayYukle CLng(lstDosyalar.List(lstDosyalar.ListIndex, 8))
End Sub

Private Sub DetayYukle(ByVal r As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("İCRA DOSYALARI")

    lblDosyaBilgi.Caption = "Daire: " & ws.Cells(r, 7).Text & "    Alacaklı: " & ws.Cells(r, 8).Text & _
                            "    IBAN: " & ws.Cells(r, 20).Text & "    Kesinti: " & KuralMetni(ws, r)
    lblKesilen.Caption = "Bordrodan Kesilen: " & Format$(Sayi(ws.Cells(r, 16).Value), "#,##0.00")
    lblOdenen.Caption = "Daireye Ödenen: " & Format$(Sayi(ws.Cells(r, 23).Value), "#,##0.00")
    lblBekleyen.Caption = "Ödenmesi Bekleyen: " & Format$(Sayi(ws.Cells(r, 24).Value), "#,##0.00")
    GecmisYukle ws.Cells(r, 1).Text, ws.Cells(r, 5).Text
End Sub

Private Function KuralMetni(ByVal ws As Worksheet, ByVal r As Long) As String
    If StrComp(Trim$(CStr(ws.Cells(r, 10).Value)), "Oran", vbTextCompare) = 0 Then
        KuralMetni = "Oran " & ws.Cells(r, 11).Text & "/" & ws.Cells(r, 12).Text
    Else
        KuralMetni = "Sabit " & Format$(Sayi(ws.Cells(r, 13).Value), "#,##0.00") & " TL/ay"
    End If
End Function

Private Sub GecmisYukle(ByVal personel As String, ByVal dosyaNo As String)
    Dim ws As Worksheet
    Dim r As Long
    Dim ix As Long

    lstKesintiler.Clear
    Set ws = ThisWorkbook.Worksheets("AYLIK KESİNTİLER")
    For r = 2 To 5000
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), Trim$(personel), vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(ws.Cells(r, 2).Value)), Trim$(dosyaNo), vbTextCompare) = 0 Then
                lstKesintiler.AddItem ws.Cells(r, 3).Text
                ix = lstKesintiler.ListCount - 1
                lstKesintiler.List(ix, 1) = ws.Cells(r, 4).Text
                lstKesintiler.List(ix, 2) = ws.Cells(r, 5).Text
                lstKesintiler.List(ix, 3) = ws.Cells(r, 6).Text
            End If
        End If
    Next r

    lstOdemeler.Clear
    Set ws = ThisWorkbook.Worksheets("İCRA ÖDEMELERİ")
    For r = 2 To 5000
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), Trim$(personel), vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(ws.Cells(r, 2).Value)), Trim$(dosyaNo), vbTextCompare) = 0 Then
                lstOdemeler.AddItem ws.Cells(r, 3).Text
                ix = lstOdemeler.ListCount - 1
                lstOdemeler.List(ix, 1) = ws.Cells(r, 4).Text
                lstOdemeler.List(ix, 2) = ws.Cells(r, 5).Text
                lstOdemeler.List(ix, 3) = ws.Cells(r, 6).Text
                lstOdemeler.List(ix, 4) = CStr(r)
            End If
        End If
    Next r
End Sub

Private Sub lstOdemeler_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstOdemeler.ListIndex < 0 Then Exit Sub
    Dim r As Long
    Dim sonuc As String
    r = CLng(lstOdemeler.List(lstOdemeler.ListIndex, 4))
    If MsgBox("Seçili ödeme kaydı silinsin mi?" & vbCrLf & _
              lstOdemeler.List(lstOdemeler.ListIndex, 0) & "  " & _
              lstOdemeler.List(lstOdemeler.ListIndex, 1) & " TL", _
              vbYesNo + vbQuestion, "Ödeme sil") <> vbYes Then Exit Sub
    sonuc = UI_OdemeSil(r)
    If Left$(sonuc, 3) = "OK|" Then
        Yenile
    Else
        MsgBox sonuc, vbExclamation, "Ödeme sil"
    End If
End Sub

Private Sub DetayTemizle()
    lblDosyaBilgi.Caption = ""
    lblKesilen.Caption = ""
    lblOdenen.Caption = ""
    lblBekleyen.Caption = ""
    lstKesintiler.Clear
    lstOdemeler.Clear
End Sub

Private Function SeciliSatir() As Long
    If lstDosyalar.ListIndex >= 0 Then
        SeciliSatir = CLng(lstDosyalar.List(lstDosyalar.ListIndex, 8))
    End If
End Function

Private Function SeciliDosyaNo() As String
    If lstDosyalar.ListIndex >= 0 Then
        SeciliDosyaNo = lstDosyalar.List(lstDosyalar.ListIndex, 1)
    End If
End Function

Private Sub cmdYeniDosya_Click()
    UI_SeciliPersonel = PersonelAd
    frmIcraGiris.Show vbModal
    Yenile
End Sub

Private Sub cmdDuzenle_Click()
    If SeciliSatir() = 0 Then
        MsgBox "Önce listeden dosya seçin.", vbExclamation, "Düzenle"
        Exit Sub
    End If
    frmIcraGiris.DuzenleSatir = SeciliSatir()
    frmIcraGiris.Show vbModal
    Yenile
End Sub

Private Sub cmdDosyaSil_Click()
    If SeciliSatir() = 0 Then
        MsgBox "Önce listeden dosya seçin.", vbExclamation, "Sil"
        Exit Sub
    End If
    If MsgBox("'" & SeciliDosyaNo() & "' dosyası ve bağlı TÜM kesinti/ödeme kayıtları silinecek. Emin misiniz?", _
              vbYesNo + vbExclamation, "Dosya sil") <> vbYes Then Exit Sub
    Dim sonuc As String
    sonuc = UI_IcraSil(SeciliSatir())
    If Left$(sonuc, 3) = "OK|" Then
        MsgBox Mid$(sonuc, 4), vbInformation, "Dosya sil"
        Yenile
    Else
        MsgBox sonuc, vbExclamation, "Dosya sil"
    End If
End Sub

Private Sub cmdKesintiEkle_Click()
    If SeciliSatir() = 0 Then
        MsgBox "Önce listeden dosya seçin.", vbExclamation, "Kesinti"
        Exit Sub
    End If
    frmKesintiGiris.OnPersonel = PersonelAd
    frmKesintiGiris.OnDosya = SeciliDosyaNo()
    frmKesintiGiris.Show vbModal
    Yenile
End Sub

Private Sub cmdOdemeEkle_Click()
    If SeciliSatir() = 0 Then
        MsgBox "Önce listeden dosya seçin.", vbExclamation, "Ödeme"
        Exit Sub
    End If
    frmOdemeGiris.OnPersonel = PersonelAd
    frmOdemeGiris.OnDosya = SeciliDosyaNo()
    frmOdemeGiris.Show vbModal
    Yenile
End Sub

Private Sub cmdYenile_Click()
    Yenile
End Sub

Private Sub cmdKapat_Click()
    Unload Me
End Sub

Private Function Sayi(ByVal v As Variant) As Double
    If IsNumeric(v) Then Sayi = CDbl(v)
End Function
