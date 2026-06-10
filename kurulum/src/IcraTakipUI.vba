Option Explicit

Public UI_SeciliPersonel As String

Private Const UI_SAYFA_PERSONEL As String = "PERSONEL"
Private Const UI_SAYFA_ICRA As String = "İCRA DOSYALARI"
Private Const UI_SAYFA_KES As String = "AYLIK KESİNTİLER"
Private Const UI_SAYFA_PANEL As String = "PANEL"
Private Const UI_SAYFA_OZET As String = "PERSONEL ÖZET"
Private Const UI_PERSONEL_SON As Long = 300
Private Const UI_ICRA_SON As Long = 300
Private Const UI_KES_SON As Long = 5000
Private Const UI_SAYFA_ODEME As String = "İCRA ÖDEMELERİ"
Private Const UI_ODEME_SON As Long = 5000

Public Sub ANA_MENU_AC()
    UI_PROGRAM_MODU_AC
    frmAnaMenu.Show vbModal
    UYGULAMA_NORMAL_GORUNUM
End Sub

Public Sub UYGULAMA_BASLAT()
    On Error Resume Next
    UI_PROGRAM_MODU_AC
    On Error GoTo 0
    frmAnaMenu.Show vbModal
    ' Menü kapandıysa (X ile dahi) Excel görünmez kalmasın:
    UYGULAMA_NORMAL_GORUNUM
End Sub

Public Sub UYGULAMA_CIKIS()
    On Error Resume Next
    ThisWorkbook.Save
    UYGULAMA_NORMAL_GORUNUM
    Application.DisplayAlerts = False
    ThisWorkbook.Close SaveChanges:=True
    If Application.Workbooks.Count = 0 Then Application.Quit
End Sub

Public Sub SAYFALARI_GIZLE()
    UI_PROGRAM_MODU_AC
    MsgBox "Program modu açıldı. Veri sayfaları gizlendi.", vbInformation, "Program modu"
End Sub

Public Sub SAYFALARI_GOSTER()
    UYGULAMA_NORMAL_GORUNUM
    MsgBox "Bakım görünümü açıldı. Sayfalar tekrar görünür durumda.", vbInformation, "Bakım görünümü"
End Sub

Public Sub PERSONEL_FORM_AC()
    frmPersonelGiris.Show vbModal
End Sub

Public Sub ICRA_FORM_AC()
    frmIcraGiris.Show vbModal
End Sub

Public Sub MANUEL_KESINTI_FORM_AC()
    frmKesintiGiris.Show vbModal
End Sub

Public Sub PROGRAM_BUTONLARI_KUR()
    Dim ws As Worksheet
    Dim b As Button
    Dim sol As Double
    Dim ust As Double
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_PANEL)

    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0

    ws.Cells.ClearFormats
    ws.Range("A1:H24").Interior.Color = RGB(247, 249, 252)
    ws.Range("B2").Value = "PERSONEL İCRA TAKİP"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 18
    ws.Range("B5").Value = "İşlemler ana menü ve modal formlardan yürütülür."
    ws.Range("B5").Font.Italic = True
    ws.Range("B5").Font.Color = RGB(90, 90, 90)

    ws.Activate
    MsgBox "Program arka planı kuruldu.", vbInformation, "Kurulum"
End Sub

Public Sub UI_PROGRAM_MODU_AC()
    On Error Resume Next
    Dim ws As Worksheet

    ThisWorkbook.Worksheets(UI_SAYFA_PANEL).Visible = xlSheetVisible
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, UI_SAYFA_PANEL, vbTextCompare) = 0 Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ThisWorkbook.Worksheets(UI_SAYFA_PANEL).Activate
    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    Application.DisplayScrollBars = False
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayWorkbookTabs = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.WindowState = xlMaximized
    End If
    Application.Visible = False
End Sub

Public Sub UYGULAMA_NORMAL_GORUNUM()
    On Error Resume Next
    Dim ws As Worksheet

    Application.Visible = True
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws

    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    Application.DisplayScrollBars = True
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayWorkbookTabs = True
        ActiveWindow.DisplayHeadings = True
        ActiveWindow.DisplayGridlines = True
    End If
    ThisWorkbook.Worksheets(UI_SAYFA_PANEL).Activate
End Sub

Public Function UI_PersonelKaydet(ByVal adSoyad As String, ByVal tcKimlik As String, _
                                  ByVal netMaasMetin As String, ByVal iseGirisMetin As String, _
                                  ByVal aciklama As String) As String
    On Error GoTo Hata
    Dim ws As Worksheet
    Dim satir As Long
    Dim netMaas As Double
    Dim tc As String
    Dim girisVar As Boolean
    Dim girisTarihi As Date

    adSoyad = Trim$(adSoyad)
    tc = UI_SadeceRakam(tcKimlik)

    If Len(adSoyad) = 0 Then
        UI_PersonelKaydet = "Personel adı soyadı zorunludur."
        Exit Function
    End If
    If Len(tc) <> 11 Then
        UI_PersonelKaydet = "T.C. Kimlik No 11 rakam olmalıdır."
        Exit Function
    End If
    If Not UI_TutarOku(netMaasMetin, netMaas, True) Or netMaas <= 0 Then
        UI_PersonelKaydet = "Net maaş geçerli bir tutar olmalıdır."
        Exit Function
    End If
    If Len(Trim$(iseGirisMetin)) > 0 Then
        If Not IsDate(iseGirisMetin) Then
            UI_PersonelKaydet = "İşe giriş tarihi geçerli değil. Örn: 10.06.2026"
            Exit Function
        End If
        girisVar = True
        girisTarihi = CDate(iseGirisMetin)
    End If

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_PERSONEL)
    If UI_PersonelTcVar(tc) Then
        UI_PersonelKaydet = "Bu T.C. Kimlik No ile personel zaten kayıtlı."
        Exit Function
    End If
    If UI_PersonelAdVar(adSoyad) Then
        If MsgBox("Aynı isimde bir personel var. Yine de kaydedilsin mi?", _
                  vbYesNo + vbQuestion, "Benzer personel") <> vbYes Then
            UI_PersonelKaydet = "Kayıt iptal edildi."
            Exit Function
        End If
    End If

    satir = UI_SonrakiBosSatir(ws, 1, 2, UI_PERSONEL_SON)
    If satir = 0 Then
        UI_PersonelKaydet = "PERSONEL sayfasında boş satır kalmadı."
        Exit Function
    End If

    If satir > 2 Then
        ws.Rows(satir - 1).Copy
        ws.Rows(satir).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If

    ws.Cells(satir, 1).Value = adSoyad
    ws.Cells(satir, 2).Value = tc
    ws.Cells(satir, 3).Value = netMaas
    If girisVar Then ws.Cells(satir, 4).Value = girisTarihi Else ws.Cells(satir, 4).ClearContents
    ws.Cells(satir, 5).Value = Trim$(aciklama)
    ws.Cells(satir, 3).NumberFormat = "#,##0.00"
    ws.Cells(satir, 4).NumberFormat = "dd.mm.yyyy"

    UI_PersonelOzetGuncelle adSoyad
    Application.Calculate

    UI_SeciliPersonel = adSoyad
    UI_PersonelKaydet = "OK|Personel kaydedildi: " & adSoyad
    Exit Function
Hata:
    UI_PersonelKaydet = "Personel kaydedilemedi: " & Err.Description
End Function

Public Function UI_PersonelGuncelle(ByVal satir As Long, ByVal adSoyad As String, ByVal tcKimlik As String, _
                                    ByVal netMaasMetin As String, ByVal iseGirisMetin As String, _
                                    ByVal aciklama As String) As String
    On Error GoTo Hata
    Dim ws As Worksheet
    Dim eskiAd As String
    Dim tc As String
    Dim netMaas As Double
    Dim girisVar As Boolean
    Dim girisTarihi As Date
    Dim r As Long

    If satir < 2 Or satir > UI_PERSONEL_SON Then
        UI_PersonelGuncelle = "Güncellenecek personel seçilmelidir."
        Exit Function
    End If

    adSoyad = Trim$(adSoyad)
    tc = UI_SadeceRakam(tcKimlik)
    If Len(adSoyad) = 0 Then
        UI_PersonelGuncelle = "Personel adı soyadı zorunludur."
        Exit Function
    End If
    If Len(tc) <> 11 Then
        UI_PersonelGuncelle = "T.C. Kimlik No 11 rakam olmalıdır."
        Exit Function
    End If
    If Not UI_TutarOku(netMaasMetin, netMaas, True) Or netMaas <= 0 Then
        UI_PersonelGuncelle = "Net maaş geçerli bir tutar olmalıdır."
        Exit Function
    End If
    If Len(Trim$(iseGirisMetin)) > 0 Then
        If Not IsDate(iseGirisMetin) Then
            UI_PersonelGuncelle = "İşe giriş tarihi geçerli değil. Örn: 10.06.2026"
            Exit Function
        End If
        girisVar = True
        girisTarihi = CDate(iseGirisMetin)
    End If

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_PERSONEL)
    eskiAd = Trim$(CStr(ws.Cells(satir, 1).Value))
    If Len(eskiAd) = 0 Then
        UI_PersonelGuncelle = "Seçili satırda personel kaydı yok."
        Exit Function
    End If

    For r = 2 To UI_PERSONEL_SON
        If r <> satir Then
            If Trim$(CStr(ws.Cells(r, 2).Value)) = tc Then
                UI_PersonelGuncelle = "Bu T.C. Kimlik No başka personelde kayıtlı."
                Exit Function
            End If
        End If
    Next r

    ws.Cells(satir, 1).Value = adSoyad
    ws.Cells(satir, 2).Value = tc
    ws.Cells(satir, 3).Value = netMaas
    If girisVar Then ws.Cells(satir, 4).Value = girisTarihi Else ws.Cells(satir, 4).ClearContents
    ws.Cells(satir, 5).Value = Trim$(aciklama)
    ws.Cells(satir, 3).NumberFormat = "#,##0.00"
    ws.Cells(satir, 4).NumberFormat = "dd.mm.yyyy"

    If StrComp(eskiAd, adSoyad, vbTextCompare) <> 0 Then
        UI_PersonelAdiniDegistir eskiAd, adSoyad
    End If
    UI_PersonelOzetGuncelle adSoyad
    Application.Calculate
    UI_PersonelGuncelle = "OK|Personel güncellendi: " & adSoyad
    Exit Function
Hata:
    UI_PersonelGuncelle = "Personel güncellenemedi: " & Err.Description
End Function

Public Function UI_PersonelSil(ByVal satir As Long) As String
    On Error GoTo Hata
    Dim wsP As Worksheet
    Dim ad As String
    Dim iliskili As Long

    If satir < 2 Or satir > UI_PERSONEL_SON Then
        UI_PersonelSil = "Silinecek personel seçilmelidir."
        Exit Function
    End If
    Set wsP = ThisWorkbook.Worksheets(UI_SAYFA_PERSONEL)
    ad = Trim$(CStr(wsP.Cells(satir, 1).Value))
    If Len(ad) = 0 Then
        UI_PersonelSil = "Seçili satırda personel kaydı yok."
        Exit Function
    End If

    iliskili = UI_PersonelIliskiliKayitSayisi(ad)
    If iliskili > 0 Then
        If MsgBox(ad & " için " & iliskili & " ilişkili icra/kesinti kaydı var." & vbCrLf & _
                  "Personel ve ilişkili kayıtlar silinsin mi?", _
                  vbYesNo + vbCritical, "Personel sil") <> vbYes Then
            UI_PersonelSil = "Silme iptal edildi."
            Exit Function
        End If
        UI_PersonelIliskiliKayitlariSil ad
    Else
        If MsgBox(ad & " silinsin mi?", vbYesNo + vbQuestion, "Personel sil") <> vbYes Then
            UI_PersonelSil = "Silme iptal edildi."
            Exit Function
        End If
    End If

    wsP.Rows(satir).Delete
    UI_PersonelOzettenSil ad
    Application.Calculate
    UI_PersonelSil = "OK|Personel silindi: " & ad
    Exit Function
Hata:
    UI_PersonelSil = "Personel silinemedi: " & Err.Description
End Function

Public Function UI_IcraKaydet(ByVal personel As String, ByVal dosyaTarihiMetin As String, _
                              ByVal dosyaNo As String, ByVal oncelikMetin As String, _
                              ByVal daire As String, ByVal alacakli As String, _
                              ByVal icraTutarMetin As String, ByVal kesintiSekli As String, _
                              ByVal payMetin As String, ByVal paydaMetin As String, _
                              ByVal sabitMetin As String, ByVal oncekiTahsilMetin As String, _
                              ByVal iban As String, ByVal aciklama As String) As String
    On Error GoTo Hata
    Dim ws As Worksheet
    Dim satir As Long
    Dim dosyaTarihi As Date
    Dim icraTutar As Double
    Dim oncelik As Long
    Dim pay As Double
    Dim payda As Double
    Dim sabit As Double
    Dim oncekiTahsil As Double

    personel = Trim$(personel)
    dosyaNo = Trim$(dosyaNo)
    kesintiSekli = Trim$(kesintiSekli)

    If Len(personel) = 0 Then
        UI_IcraKaydet = "Personel seçilmelidir."
        Exit Function
    End If
    If Not UI_PersonelAdVar(personel) Then
        UI_IcraKaydet = "Seçilen personel PERSONEL sayfasında bulunamadı."
        Exit Function
    End If
    If Len(dosyaNo) = 0 Then
        UI_IcraKaydet = "Dosya numarası zorunludur."
        Exit Function
    End If
    If UI_IcraDosyaVar(personel, dosyaNo) Then
        UI_IcraKaydet = "Bu personel için aynı dosya numarası zaten kayıtlı."
        Exit Function
    End If
    If Len(Trim$(dosyaTarihiMetin)) = 0 Or Not IsDate(dosyaTarihiMetin) Then
        UI_IcraKaydet = "Dosya tarihi geçerli değil. Örn: 10.06.2026"
        Exit Function
    End If
    dosyaTarihi = CDate(dosyaTarihiMetin)
    If Len(Trim$(oncelikMetin)) = 0 Then
        oncelik = UI_SonrakiOncelik(personel)
    ElseIf IsNumeric(oncelikMetin) Then
        oncelik = CLng(oncelikMetin)
    Else
        UI_IcraKaydet = "Öncelik sırası sayı olmalıdır."
        Exit Function
    End If
    If oncelik <= 0 Then
        UI_IcraKaydet = "Öncelik sırası 1 veya daha büyük olmalıdır."
        Exit Function
    End If
    If Not UI_TutarOku(icraTutarMetin, icraTutar, True) Or icraTutar <= 0 Then
        UI_IcraKaydet = "İcra tutarı geçerli bir tutar olmalıdır."
        Exit Function
    End If
    If StrComp(kesintiSekli, "Oran", vbTextCompare) = 0 Then
        If Not UI_TutarOku(payMetin, pay, True) Or pay <= 0 Then
            UI_IcraKaydet = "Oran - Pay geçerli olmalıdır."
            Exit Function
        End If
        If Not UI_TutarOku(paydaMetin, payda, True) Or payda <= 0 Then
            UI_IcraKaydet = "Oran - Payda geçerli olmalıdır."
            Exit Function
        End If
        sabit = 0
    ElseIf StrComp(kesintiSekli, "Sabit Tutar", vbTextCompare) = 0 Then
        If Not UI_TutarOku(sabitMetin, sabit, True) Or sabit <= 0 Then
            UI_IcraKaydet = "Sabit aylık kesinti geçerli bir tutar olmalıdır."
            Exit Function
        End If
        pay = 0
        payda = 0
    Else
        UI_IcraKaydet = "Kesinti şekli Oran veya Sabit Tutar olmalıdır."
        Exit Function
    End If
    If Not UI_TutarOku(oncekiTahsilMetin, oncekiTahsil, False) Then
        UI_IcraKaydet = "Önceki firmada tahsil geçerli bir tutar olmalıdır."
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ICRA)
    satir = UI_SonrakiBosSatir(ws, 5, 2, UI_ICRA_SON)
    If satir = 0 Then
        UI_IcraKaydet = "İCRA DOSYALARI sayfasında boş satır kalmadı."
        Exit Function
    End If

    If satir > 2 Then
        ws.Rows(satir - 1).Copy
        ws.Rows(satir).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If

    ws.Cells(satir, 1).Value = personel
    ws.Cells(satir, 4).Value = dosyaTarihi
    ws.Cells(satir, 5).Value = dosyaNo
    ws.Cells(satir, 6).Value = oncelik
    ws.Cells(satir, 7).Value = Trim$(daire)
    ws.Cells(satir, 8).Value = Trim$(alacakli)
    ws.Cells(satir, 9).Value = icraTutar
    ws.Cells(satir, 10).Value = IIf(StrComp(kesintiSekli, "Oran", vbTextCompare) = 0, "Oran", "Sabit Tutar")
    If pay > 0 Then ws.Cells(satir, 11).Value = pay Else ws.Cells(satir, 11).ClearContents
    If payda > 0 Then ws.Cells(satir, 12).Value = payda Else ws.Cells(satir, 12).ClearContents
    If sabit > 0 Then ws.Cells(satir, 13).Value = sabit Else ws.Cells(satir, 13).ClearContents
    ws.Cells(satir, 15).Value = oncekiTahsil
    ws.Cells(satir, 20).Value = Trim$(iban)
    ws.Cells(satir, 21).Value = Trim$(aciklama)
    UI_IcraFormulleriniYaz ws, satir
    ws.Range(ws.Cells(satir, 4), ws.Cells(satir, 4)).NumberFormat = "dd.mm.yyyy"
    ws.Range(ws.Cells(satir, 9), ws.Cells(satir, 18)).NumberFormat = "#,##0.00"
    Application.Calculate

    UI_IcraKaydet = "OK|İcra dosyası kaydedildi: " & personel & " - " & dosyaNo
    Exit Function
Hata:
    UI_IcraKaydet = "İcra dosyası kaydedilemedi: " & Err.Description
End Function

Public Function UI_ManuelKesintiKaydet(ByVal personel As String, ByVal dosyaNo As String, _
                                        ByVal yilMetin As String, ByVal ay As String, _
                                        ByVal tutarMetin As String, ByVal aciklama As String) As String
    On Error GoTo Hata
    Dim ws As Worksheet
    Dim satir As Long
    Dim yil As Long
    Dim tutar As Double

    personel = Trim$(personel)
    dosyaNo = Trim$(dosyaNo)
    ay = Trim$(ay)

    If Len(personel) = 0 Then
        UI_ManuelKesintiKaydet = "Personel seçilmelidir."
        Exit Function
    End If
    If Len(dosyaNo) = 0 Then
        UI_ManuelKesintiKaydet = "Dosya numarası seçilmelidir."
        Exit Function
    End If
    If Not IsNumeric(yilMetin) Then
        UI_ManuelKesintiKaydet = "Yıl sayı olmalıdır."
        Exit Function
    End If
    yil = CLng(yilMetin)
    If yil < 2000 Or yil > 2100 Then
        UI_ManuelKesintiKaydet = "Yıl 2000 - 2100 arasında olmalıdır."
        Exit Function
    End If
    If Not UI_AyGecerli(ay) Then
        UI_ManuelKesintiKaydet = "Geçerli bir ay seçilmelidir."
        Exit Function
    End If
    If Not UI_TutarOku(tutarMetin, tutar, True) Or tutar <= 0 Then
        UI_ManuelKesintiKaydet = "Kesinti tutarı geçerli bir tutar olmalıdır."
        Exit Function
    End If
    If UI_KesintiVar(personel, dosyaNo, yil, ay) Then
        If MsgBox("Bu personel/dosya için aynı dönemde kesinti var. Yine de eklensin mi?", _
                  vbYesNo + vbQuestion, "Mükerrer kontrol") <> vbYes Then
            UI_ManuelKesintiKaydet = "Kayıt iptal edildi."
            Exit Function
        End If
    End If

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_KES)
    satir = UI_SonrakiBosSatir(ws, 1, 2, UI_KES_SON)
    If satir = 0 Then
        UI_ManuelKesintiKaydet = "AYLIK KESİNTİLER sayfasında boş satır kalmadı."
        Exit Function
    End If
    If satir > 2 Then
        ws.Rows(satir - 1).Copy
        ws.Rows(satir).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If

    ws.Cells(satir, 1).Value = personel
    ws.Cells(satir, 2).Value = dosyaNo
    ws.Cells(satir, 3).Value = yil
    ws.Cells(satir, 4).Value = ay
    ws.Cells(satir, 5).Value = tutar
    ws.Cells(satir, 6).Value = "MANUEL (" & Format$(Now, "dd.mm.yyyy hh:nn") & ")" & _
                               IIf(Len(Trim$(aciklama)) > 0, " - " & Trim$(aciklama), "")
    ws.Cells(satir, 5).NumberFormat = "#,##0.00"
    Application.Calculate

    UI_ManuelKesintiKaydet = "OK|Manuel kesinti kaydedildi."
    Exit Function
Hata:
    UI_ManuelKesintiKaydet = "Kesinti kaydedilemedi: " & Err.Description
End Function

Public Sub UI_PersonelComboYukle(ByVal cbo As Object)
    Dim ws As Worksheet
    Dim r As Long
    cbo.Clear
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_PERSONEL)
    For r = 2 To UI_PERSONEL_SON
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0 Then
            cbo.AddItem CStr(ws.Cells(r, 1).Value)
        End If
    Next r
End Sub

Public Sub UI_AyComboYukle(ByVal cbo As Object)
    Dim aylar As Variant
    Dim i As Long
    aylar = Array("Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", _
                  "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık")
    cbo.Clear
    For i = LBound(aylar) To UBound(aylar)
        cbo.AddItem CStr(aylar(i))
    Next i
End Sub

Public Sub UI_DosyaComboYukle(ByVal cbo As Object, ByVal personel As String)
    Dim ws As Worksheet
    Dim r As Long
    cbo.Clear
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ICRA)
    For r = 2 To UI_ICRA_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), Trim$(personel), vbTextCompare) = 0 Then
            If Len(Trim$(CStr(ws.Cells(r, 5).Value))) > 0 Then cbo.AddItem CStr(ws.Cells(r, 5).Value)
        End If
    Next r
End Sub

Public Function UI_PersonelBilgi(ByVal personel As String) As String
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_PERSONEL)
    For r = 2 To UI_PERSONEL_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), Trim$(personel), vbTextCompare) = 0 Then
            UI_PersonelBilgi = "TC: " & ws.Cells(r, 2).Text & " | Net Maaş: " & _
                               Format$(UI_Sayi(ws.Cells(r, 3).Value), "#,##0.00") & " TL"
            Exit Function
        End If
    Next r
    UI_PersonelBilgi = ""
End Function

Public Function UI_SonrakiOncelik(ByVal personel As String) As Long
    Dim ws As Worksheet
    Dim r As Long
    Dim enBuyuk As Long
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ICRA)
    For r = 2 To UI_ICRA_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), Trim$(personel), vbTextCompare) = 0 Then
            If IsNumeric(ws.Cells(r, 6).Value) Then
                If CLng(ws.Cells(r, 6).Value) > enBuyuk Then enBuyuk = CLng(ws.Cells(r, 6).Value)
            End If
        End If
    Next r
    UI_SonrakiOncelik = enBuyuk + 1
End Function

Private Sub UI_IcraFormulleriniYaz(ByVal ws As Worksheet, ByVal satir As Long)
    ws.Cells(satir, 2).FormulaR1C1 = "=IF(RC[-1]="""","""",IFERROR(VLOOKUP(RC[-1],PERSONEL!R2C1:R300C3,2,0),""""))"
    ws.Cells(satir, 3).FormulaR1C1 = "=IF(RC[-2]="""","""",IFERROR(VLOOKUP(RC[-2],PERSONEL!R2C1:R300C3,3,0),""""))"
    ws.Cells(satir, 14).FormulaR1C1 = "=IF(RC5="""","""",IF(RC19<>""ÖDEMEDE"",0,MIN(RC18,IF(RC10=""Oran"",IFERROR(RC3*RC11/RC12,0),RC13))))"
    ws.Cells(satir, 16).FormulaR1C1 = "=IF(RC5="""","""",SUMIFS('AYLIK KESİNTİLER'!R2C5:R5000C5,'AYLIK KESİNTİLER'!R2C1:R5000C1,RC1,'AYLIK KESİNTİLER'!R2C2:R5000C2,RC5))"
    ws.Cells(satir, 17).FormulaR1C1 = "=IF(RC5="""","""",RC15+RC16)"
    ws.Cells(satir, 18).FormulaR1C1 = "=IF(RC5="""","""",RC9-RC17)"
    ws.Cells(satir, 19).FormulaR1C1 = "=IF(RC5="""","""",IF(RC18<=0,""KAPANDI"",IF(COUNTIFS(R2C1:R300C1,RC1,R2C6:R300C6,""<""&RC6,R2C18:R300C18,"">0"")=0,""ÖDEMEDE"",""SIRADA"")))"
    ws.Cells(satir, 22).FormulaR1C1 = "=IF(RC5="""","""",RC1&""|""&RC19)"
    ws.Cells(satir, 23).FormulaR1C1 = "=IF(RC5="""","""",SUMIFS('İCRA ÖDEMELERİ'!R2C4:R5000C4,'İCRA ÖDEMELERİ'!R2C1:R5000C1,RC1,'İCRA ÖDEMELERİ'!R2C2:R5000C2,RC5))"
    ws.Cells(satir, 24).FormulaR1C1 = "=IF(RC5="""","""",RC16-RC23)"
    ws.Range(ws.Cells(satir, 23), ws.Cells(satir, 24)).NumberFormat = "#,##0.00"
End Sub

Private Sub UI_PersonelOzetGuncelle(ByVal personel As String)
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_OZET)
    For r = 2 To 100
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), Trim$(personel), vbTextCompare) = 0 Then Exit Sub
    Next r
    For r = 2 To 100
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then
            ws.Cells(r, 1).Value = personel
            If r > 2 Then
                ws.Range(ws.Cells(r - 1, 2), ws.Cells(r - 1, 9)).Copy
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 9)).PasteSpecial xlPasteFormulasAndNumberFormats
                Application.CutCopyMode = False
            End If
            Exit Sub
        End If
    Next r
End Sub

Private Function UI_PersonelTcVar(ByVal tc As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_PERSONEL)
    For r = 2 To UI_PERSONEL_SON
        If Trim$(CStr(ws.Cells(r, 2).Value)) = tc Then
            UI_PersonelTcVar = True
            Exit Function
        End If
    Next r
End Function

Private Sub UI_PersonelAdiniDegistir(ByVal eskiAd As String, ByVal yeniAd As String)
    Dim ws As Worksheet
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ICRA)
    For r = 2 To UI_ICRA_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), eskiAd, vbTextCompare) = 0 Then
            ws.Cells(r, 1).Value = yeniAd
        End If
    Next r

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_KES)
    For r = 2 To UI_KES_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), eskiAd, vbTextCompare) = 0 Then
            ws.Cells(r, 1).Value = yeniAd
        End If
    Next r

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ODEME)
    For r = 2 To UI_ODEME_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), eskiAd, vbTextCompare) = 0 Then
            ws.Cells(r, 1).Value = yeniAd
        End If
    Next r

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_OZET)
    For r = 2 To 100
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), eskiAd, vbTextCompare) = 0 Then
            ws.Cells(r, 1).Value = yeniAd
        End If
    Next r
End Sub

Private Function UI_PersonelIliskiliKayitSayisi(ByVal ad As String) As Long
    Dim ws As Worksheet
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ICRA)
    For r = 2 To UI_ICRA_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), ad, vbTextCompare) = 0 Then
            UI_PersonelIliskiliKayitSayisi = UI_PersonelIliskiliKayitSayisi + 1
        End If
    Next r

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_KES)
    For r = 2 To UI_KES_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), ad, vbTextCompare) = 0 Then
            UI_PersonelIliskiliKayitSayisi = UI_PersonelIliskiliKayitSayisi + 1
        End If
    Next r

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ODEME)
    For r = 2 To UI_ODEME_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), ad, vbTextCompare) = 0 Then
            UI_PersonelIliskiliKayitSayisi = UI_PersonelIliskiliKayitSayisi + 1
        End If
    Next r
End Function

Private Sub UI_PersonelIliskiliKayitlariSil(ByVal ad As String)
    Dim ws As Worksheet
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ICRA)
    For r = UI_ICRA_SON To 2 Step -1
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), ad, vbTextCompare) = 0 Then ws.Rows(r).Delete
    Next r

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_KES)
    For r = UI_KES_SON To 2 Step -1
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), ad, vbTextCompare) = 0 Then ws.Rows(r).Delete
    Next r

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ODEME)
    For r = UI_ODEME_SON To 2 Step -1
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), ad, vbTextCompare) = 0 Then ws.Rows(r).Delete
    Next r
End Sub

Private Sub UI_PersonelOzettenSil(ByVal ad As String)
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_OZET)
    For r = 100 To 2 Step -1
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), ad, vbTextCompare) = 0 Then ws.Rows(r).Delete
    Next r
End Sub

Private Function UI_PersonelAdVar(ByVal ad As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_PERSONEL)
    For r = 2 To UI_PERSONEL_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), Trim$(ad), vbTextCompare) = 0 Then
            UI_PersonelAdVar = True
            Exit Function
        End If
    Next r
End Function

Private Function UI_IcraDosyaVar(ByVal personel As String, ByVal dosyaNo As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ICRA)
    For r = 2 To UI_ICRA_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), Trim$(personel), vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(ws.Cells(r, 5).Value)), Trim$(dosyaNo), vbTextCompare) = 0 Then
                UI_IcraDosyaVar = True
                Exit Function
            End If
        End If
    Next r
End Function

Private Function UI_KesintiVar(ByVal personel As String, ByVal dosyaNo As String, _
                               ByVal yil As Long, ByVal ay As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_KES)
    For r = 2 To UI_KES_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), Trim$(personel), vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(ws.Cells(r, 2).Value)), Trim$(dosyaNo), vbTextCompare) = 0 Then
                If CLng(Val(ws.Cells(r, 3).Value)) = yil Then
                    If StrComp(Trim$(CStr(ws.Cells(r, 4).Value)), Trim$(ay), vbTextCompare) = 0 Then
                        UI_KesintiVar = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next r
End Function

Private Function UI_SonrakiBosSatir(ByVal ws As Worksheet, ByVal kolon As Long, _
                                    ByVal ilkSatir As Long, ByVal sonSatir As Long) As Long
    Dim r As Long
    For r = ilkSatir To sonSatir
        If Len(Trim$(CStr(ws.Cells(r, kolon).Value))) = 0 Then
            UI_SonrakiBosSatir = r
            Exit Function
        End If
    Next r
End Function

Private Function UI_TutarOku(ByVal metin As String, ByRef sonuc As Double, _
                             Optional ByVal zorunlu As Boolean = True) As Boolean
    Dim s As String
    s = Trim$(metin)
    sonuc = 0
    If Len(s) = 0 Then
        UI_TutarOku = Not zorunlu
        Exit Function
    End If
    s = Replace(s, " ", "")
    s = Replace(s, "?", "")
    s = Replace(s, "TL", "", 1, -1, vbTextCompare)
    s = UI_NormalSayiMetni(s)
    If IsNumeric(s) Then
        sonuc = CDbl(s)
        UI_TutarOku = True
    End If
End Function

Private Function UI_NormalSayiMetni(ByVal s As String) As String
    Dim pNokta As Long
    Dim pVirgul As Long
    Dim ayrac As String
    ayrac = Application.DecimalSeparator
    pNokta = InStrRev(s, ".")
    pVirgul = InStrRev(s, ",")

    If pNokta > 0 And pVirgul > 0 Then
        If pVirgul > pNokta Then
            s = Replace(s, ".", "")
            s = Replace(s, ",", ayrac)
        Else
            s = Replace(s, ",", "")
            s = Replace(s, ".", ayrac)
        End If
    ElseIf pNokta > 0 Then
        If Len(s) - pNokta = 3 Then
            s = Replace(s, ".", "")
        Else
            s = Replace(s, ".", ayrac)
        End If
    ElseIf pVirgul > 0 Then
        s = Replace(s, ",", ayrac)
    End If
    UI_NormalSayiMetni = s
End Function

Private Function UI_SadeceRakam(ByVal metin As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(metin)
        ch = Mid$(metin, i, 1)
        If ch >= "0" And ch <= "9" Then UI_SadeceRakam = UI_SadeceRakam & ch
    Next i
End Function

Private Function UI_Sayi(ByVal v As Variant) As Double
    If IsNumeric(v) Then UI_Sayi = CDbl(v)
End Function

Private Function UI_AyGecerli(ByVal ay As String) As Boolean
    Select Case ay
        Case "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", _
             "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"
            UI_AyGecerli = True
    End Select
End Function

Public Function UI_AyAdi(ByVal tarih As Date) As String
    Dim aylar As Variant
    aylar = Array("", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", _
                  "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık")
    UI_AyAdi = CStr(aylar(Month(tarih)))
End Function


' =====================================================================
'  İCRA DAİRESİNE ÖDEME KAYITLARI
' =====================================================================
Public Function UI_OdemeKaydet(ByVal personel As String, ByVal dosyaNo As String, _
                               ByVal tarihMetin As String, ByVal tutarMetin As String, _
                               ByVal dekont As String, ByVal aciklama As String) As String
    On Error GoTo Hata
    Dim ws As Worksheet
    Dim satir As Long
    Dim tarih As Date
    Dim tutar As Double

    personel = Trim$(personel)
    dosyaNo = Trim$(dosyaNo)

    If Len(personel) = 0 Then
        UI_OdemeKaydet = "Personel seçilmelidir."
        Exit Function
    End If
    If Len(dosyaNo) = 0 Then
        UI_OdemeKaydet = "Dosya numarası seçilmelidir."
        Exit Function
    End If
    If Not UI_IcraDosyaVar(personel, dosyaNo) Then
        UI_OdemeKaydet = "Bu personel için böyle bir icra dosyası bulunamadı."
        Exit Function
    End If
    If Len(Trim$(tarihMetin)) = 0 Or Not IsDate(tarihMetin) Then
        UI_OdemeKaydet = "Ödeme tarihi geçerli değil. Örn: 10.06.2026"
        Exit Function
    End If
    tarih = CDate(tarihMetin)
    If Not UI_TutarOku(tutarMetin, tutar, True) Or tutar <= 0 Then
        UI_OdemeKaydet = "Ödeme tutarı geçerli bir tutar olmalıdır."
        Exit Function
    End If

    Dim bekleyen As Double
    bekleyen = UI_DosyaBekleyen(personel, dosyaNo)
    If tutar > bekleyen + 0.005 Then
        If MsgBox("Bu dosyada ödenmesi bekleyen tutar " & Format$(bekleyen, "#,##0.00") & _
                  " TL. Girdiğiniz ödeme bunu aşıyor. Yine de kaydedilsin mi?", _
                  vbYesNo + vbQuestion, "Bekleyen tutar aşımı") <> vbYes Then
            UI_OdemeKaydet = "Kayıt iptal edildi."
            Exit Function
        End If
    End If

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ODEME)
    satir = UI_SonrakiBosSatir(ws, 1, 2, UI_ODEME_SON)
    If satir = 0 Then
        UI_OdemeKaydet = "İCRA ÖDEMELERİ sayfasında boş satır kalmadı."
        Exit Function
    End If

    ws.Cells(satir, 1).Value = personel
    ws.Cells(satir, 2).Value = dosyaNo
    ws.Cells(satir, 3).Value = tarih
    ws.Cells(satir, 4).Value = tutar
    ws.Cells(satir, 5).Value = Trim$(dekont)
    ws.Cells(satir, 6).Value = Trim$(aciklama)
    ws.Cells(satir, 3).NumberFormat = "dd.mm.yyyy"
    ws.Cells(satir, 4).NumberFormat = "#,##0.00"
    Application.Calculate

    UI_OdemeKaydet = "OK|Ödeme kaydedildi: " & Format$(tutar, "#,##0.00") & " TL - " & dosyaNo
    Exit Function
Hata:
    UI_OdemeKaydet = "Ödeme kaydedilemedi: " & Err.Description
End Function

Public Function UI_OdemeSil(ByVal satir As Long) As String
    On Error GoTo Hata
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ODEME)
    If satir < 2 Or Len(Trim$(CStr(ws.Cells(satir, 1).Value))) = 0 Then
        UI_OdemeSil = "Silinecek ödeme kaydı bulunamadı."
        Exit Function
    End If
    ws.Rows(satir).Delete
    Application.Calculate
    UI_OdemeSil = "OK|Ödeme kaydı silindi."
    Exit Function
Hata:
    UI_OdemeSil = "Ödeme silinemedi: " & Err.Description
End Function

' Dosyada kesilen ama daireye henüz ödenmeyen tutar
Public Function UI_DosyaBekleyen(ByVal personel As String, ByVal dosyaNo As String) As Double
    Dim wsI As Worksheet
    Dim r As Long
    Set wsI = ThisWorkbook.Worksheets(UI_SAYFA_ICRA)
    For r = 2 To UI_ICRA_SON
        If StrComp(Trim$(CStr(wsI.Cells(r, 1).Value)), Trim$(personel), vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(wsI.Cells(r, 5).Value)), Trim$(dosyaNo), vbTextCompare) = 0 Then
                UI_DosyaBekleyen = UI_Sayi(wsI.Cells(r, 24).Value)
                Exit Function
            End If
        End If
    Next r
End Function

' =====================================================================
'  İCRA DOSYASI GÜNCELLEME / SİLME
' =====================================================================
Public Function UI_IcraGuncelle(ByVal satir As Long, ByVal personel As String, _
                                ByVal dosyaTarihiMetin As String, ByVal dosyaNo As String, _
                                ByVal oncelikMetin As String, ByVal daire As String, _
                                ByVal alacakli As String, ByVal icraTutarMetin As String, _
                                ByVal kesintiSekli As String, ByVal payMetin As String, _
                                ByVal paydaMetin As String, ByVal sabitMetin As String, _
                                ByVal oncekiTahsilMetin As String, ByVal iban As String, _
                                ByVal aciklama As String) As String
    On Error GoTo Hata
    Dim ws As Worksheet
    Dim eskiDosyaNo As String
    Dim dosyaTarihi As Date
    Dim icraTutar As Double
    Dim oncelik As Long
    Dim pay As Double
    Dim payda As Double
    Dim sabit As Double
    Dim oncekiTahsil As Double

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ICRA)
    personel = Trim$(personel)
    dosyaNo = Trim$(dosyaNo)
    kesintiSekli = Trim$(kesintiSekli)
    eskiDosyaNo = Trim$(CStr(ws.Cells(satir, 5).Value))

    If satir < 2 Or Len(eskiDosyaNo) = 0 Then
        UI_IcraGuncelle = "Güncellenecek dosya satırı bulunamadı."
        Exit Function
    End If
    If StrComp(Trim$(CStr(ws.Cells(satir, 1).Value)), personel, vbTextCompare) <> 0 Then
        UI_IcraGuncelle = "Dosyanın personeli değiştirilemez. Dosyayı silip doğru personele yeniden girin."
        Exit Function
    End If
    If Len(dosyaNo) = 0 Then
        UI_IcraGuncelle = "Dosya numarası zorunludur."
        Exit Function
    End If
    If StrComp(dosyaNo, eskiDosyaNo, vbTextCompare) <> 0 Then
        If UI_IcraDosyaVar(personel, dosyaNo) Then
            UI_IcraGuncelle = "Bu personel için aynı dosya numarası zaten kayıtlı."
            Exit Function
        End If
    End If
    If Len(Trim$(dosyaTarihiMetin)) = 0 Or Not IsDate(dosyaTarihiMetin) Then
        UI_IcraGuncelle = "Dosya tarihi geçerli değil. Örn: 10.06.2026"
        Exit Function
    End If
    dosyaTarihi = CDate(dosyaTarihiMetin)
    If Len(Trim$(oncelikMetin)) = 0 Or Not IsNumeric(oncelikMetin) Then
        UI_IcraGuncelle = "Öncelik sırası sayı olmalıdır."
        Exit Function
    End If
    oncelik = CLng(oncelikMetin)
    If oncelik <= 0 Then
        UI_IcraGuncelle = "Öncelik sırası 1 veya daha büyük olmalıdır."
        Exit Function
    End If
    If Not UI_TutarOku(icraTutarMetin, icraTutar, True) Or icraTutar <= 0 Then
        UI_IcraGuncelle = "İcra tutarı geçerli bir tutar olmalıdır."
        Exit Function
    End If
    If StrComp(kesintiSekli, "Oran", vbTextCompare) = 0 Then
        If Not UI_TutarOku(payMetin, pay, True) Or pay <= 0 Then
            UI_IcraGuncelle = "Oran - Pay geçerli olmalıdır."
            Exit Function
        End If
        If Not UI_TutarOku(paydaMetin, payda, True) Or payda <= 0 Then
            UI_IcraGuncelle = "Oran - Payda geçerli olmalıdır."
            Exit Function
        End If
        sabit = 0
    ElseIf StrComp(kesintiSekli, "Sabit Tutar", vbTextCompare) = 0 Then
        If Not UI_TutarOku(sabitMetin, sabit, True) Or sabit <= 0 Then
            UI_IcraGuncelle = "Sabit aylık kesinti geçerli bir tutar olmalıdır."
            Exit Function
        End If
        pay = 0
        payda = 0
    Else
        UI_IcraGuncelle = "Kesinti şekli Oran veya Sabit Tutar olmalıdır."
        Exit Function
    End If
    If Not UI_TutarOku(oncekiTahsilMetin, oncekiTahsil, False) Then
        UI_IcraGuncelle = "Önceki firmada tahsil geçerli bir tutar olmalıdır."
        Exit Function
    End If

    ws.Cells(satir, 4).Value = dosyaTarihi
    ws.Cells(satir, 5).Value = dosyaNo
    ws.Cells(satir, 6).Value = oncelik
    ws.Cells(satir, 7).Value = Trim$(daire)
    ws.Cells(satir, 8).Value = Trim$(alacakli)
    ws.Cells(satir, 9).Value = icraTutar
    ws.Cells(satir, 10).Value = IIf(StrComp(kesintiSekli, "Oran", vbTextCompare) = 0, "Oran", "Sabit Tutar")
    If pay > 0 Then ws.Cells(satir, 11).Value = pay Else ws.Cells(satir, 11).ClearContents
    If payda > 0 Then ws.Cells(satir, 12).Value = payda Else ws.Cells(satir, 12).ClearContents
    If sabit > 0 Then ws.Cells(satir, 13).Value = sabit Else ws.Cells(satir, 13).ClearContents
    ws.Cells(satir, 15).Value = oncekiTahsil
    ws.Cells(satir, 20).Value = Trim$(iban)
    ws.Cells(satir, 21).Value = Trim$(aciklama)
    UI_IcraFormulleriniYaz ws, satir

    ' Dosya no değiştiyse kesinti ve ödeme kayıtlarını da taşı
    If StrComp(dosyaNo, eskiDosyaNo, vbTextCompare) <> 0 Then
        UI_DosyaNoDegistir personel, eskiDosyaNo, dosyaNo
    End If
    Application.Calculate

    UI_IcraGuncelle = "OK|İcra dosyası güncellendi: " & personel & " - " & dosyaNo
    Exit Function
Hata:
    UI_IcraGuncelle = "İcra dosyası güncellenemedi: " & Err.Description
End Function

Public Function UI_IcraSil(ByVal satir As Long) As String
    On Error GoTo Hata
    Dim ws As Worksheet
    Dim personel As String
    Dim dosyaNo As String
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ICRA)
    personel = Trim$(CStr(ws.Cells(satir, 1).Value))
    dosyaNo = Trim$(CStr(ws.Cells(satir, 5).Value))
    If satir < 2 Or Len(dosyaNo) = 0 Then
        UI_IcraSil = "Silinecek dosya satırı bulunamadı."
        Exit Function
    End If

    ' Dosyaya bağlı kesinti ve ödeme kayıtlarını da sil
    Dim wsK As Worksheet
    Set wsK = ThisWorkbook.Worksheets(UI_SAYFA_KES)
    For r = UI_KES_SON To 2 Step -1
        If StrComp(Trim$(CStr(wsK.Cells(r, 1).Value)), personel, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(wsK.Cells(r, 2).Value)), dosyaNo, vbTextCompare) = 0 Then wsK.Rows(r).Delete
        End If
    Next r
    Set wsK = ThisWorkbook.Worksheets(UI_SAYFA_ODEME)
    For r = UI_ODEME_SON To 2 Step -1
        If StrComp(Trim$(CStr(wsK.Cells(r, 1).Value)), personel, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(wsK.Cells(r, 2).Value)), dosyaNo, vbTextCompare) = 0 Then wsK.Rows(r).Delete
        End If
    Next r
    ws.Rows(satir).Delete
    Application.Calculate

    UI_IcraSil = "OK|İcra dosyası ve bağlı kayıtları silindi: " & dosyaNo
    Exit Function
Hata:
    UI_IcraSil = "İcra dosyası silinemedi: " & Err.Description
End Function

Private Sub UI_DosyaNoDegistir(ByVal personel As String, ByVal eskiNo As String, ByVal yeniNo As String)
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_KES)
    For r = 2 To UI_KES_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), personel, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(ws.Cells(r, 2).Value)), eskiNo, vbTextCompare) = 0 Then
                ws.Cells(r, 2).Value = yeniNo
            End If
        End If
    Next r
    Set ws = ThisWorkbook.Worksheets(UI_SAYFA_ODEME)
    For r = 2 To UI_ODEME_SON
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), personel, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(ws.Cells(r, 2).Value)), eskiNo, vbTextCompare) = 0 Then
                ws.Cells(r, 2).Value = yeniNo
            End If
        End If
    Next r
End Sub
