Option Explicit

' =====================================================================
'  PERSONEL İCRA TAKİP MAKROLARI
'  Kurulum: Excel'de Alt+F11 > Dosya > Dosya Al... > IcraTakip.bas
'           Sonra dosyayı .xlsm olarak kaydedin (KULLANIM sayfasına bakın).
'
'  BUTONLARI_KUR       : PANEL sayfasına butonları ekler (bir kez çalıştırın).
'  KESINTILERI_KAYDET  : PANEL'de seçili yıl/ay için kesintileri hesaplar,
'                        AYLIK KESİNTİLER sayfasına işler. Kural:
'                        havuz = net maaş x pay/payda (veya sabit tutar);
'                        öncelik sırasındaki açık dosyalara dağıtılır,
'                        kalan borç aşılmaz, dosya kapanırsa artan tutar
'                        aynı ay sıradaki dosyaya taşar. Aynı ay için
'                        mükerrer kayıt oluşturmaz.
'  SECILI_AYI_GERI_AL  : Seçili ayın OTOMATİK işaretli satırlarını siler.
' =====================================================================

Private Const SAYFA_ICRA As String = "İCRA DOSYALARI"
Private Const SAYFA_KES As String = "AYLIK KESİNTİLER"
Private Const SAYFA_PANEL As String = "PANEL"
Private Const ICRA_SON As Long = 300
Private Const ETIKET As String = "OTOMATİK"
Private Const MAKS_DOSYA As Long = 64

' --- İCRA DOSYALARI sayfası kolon numaraları ---
Private Const C_AD As Long = 1        ' A  Personel Adı Soyadı
Private Const C_NET As Long = 3       ' C  Net Maaş
Private Const C_DOSYA As Long = 5     ' E  Dosya Numarası
Private Const C_ONCELIK As Long = 6   ' F  Öncelik Sırası
Private Const C_SEKIL As Long = 10    ' J  Kesinti Şekli
Private Const C_PAY As Long = 11      ' K  Oran - Pay
Private Const C_PAYDA As Long = 12    ' L  Oran - Payda
Private Const C_SABIT As Long = 13    ' M  Sabit Aylık Kesinti
Private Const C_KALAN As Long = 18    ' R  Kalan Borç

Public Sub BUTONLARI_KUR()
    Dim ws As Worksheet
    Dim b As Button
    Set ws = ThisWorkbook.Worksheets(SAYFA_PANEL)
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    Set b = ws.Buttons.Add(ws.Range("B11").Left + 4, ws.Range("B11").Top + 6, 270, 40)
    b.Caption = "KESİNTİLERİ HESAPLA ve KAYDET"
    b.OnAction = "KESINTILERI_KAYDET"
    b.Font.Bold = True
    b.Font.Size = 11
    Set b = ws.Buttons.Add(ws.Range("B11").Left + 4, ws.Range("B11").Top + 56, 270, 26)
    b.Caption = "Seçili ayın OTOMATİK kayıtlarını GERİ AL"
    b.OnAction = "SECILI_AYI_GERI_AL"
    b.Font.Size = 9
    MsgBox "Butonlar PANEL sayfasına eklendi." & vbCrLf & _
           "Artık PANEL'den yıl/ay seçip butonla kaydedebilirsiniz.", _
           vbInformation, "Kurulum tamamlandı"
End Sub

Public Sub KESINTILERI_KAYDET()
    Dim wsI As Worksheet
    Dim wsK As Worksheet
    Dim wsP As Worksheet
    Set wsI = ThisWorkbook.Worksheets(SAYFA_ICRA)
    Set wsK = ThisWorkbook.Worksheets(SAYFA_KES)
    Set wsP = ThisWorkbook.Worksheets(SAYFA_PANEL)

    Dim yil As Long
    Dim ay As String
    If Not IsNumeric(wsP.Range("C3").Value) Then
        MsgBox "PANEL sayfasında C3 hücresine yıl girin (örn. 2026).", vbExclamation
        Exit Sub
    End If
    yil = CLng(wsP.Range("C3").Value)
    ay = Trim$(CStr(wsP.Range("C4").Value))
    If yil < 2000 Or yil > 2100 Then
        MsgBox "Geçerli bir yıl girin (2000 - 2100).", vbExclamation
        Exit Sub
    End If
    If Not AyGecerli(ay) Then
        MsgBox "PANEL sayfasında C4 hücresinden ayı seçin.", vbExclamation
        Exit Sub
    End If

    Application.Calculate

    ' Mevcut kayıtları oku (aynı ay için mükerrer girişi önler)
    Dim mevcut As Object
    Set mevcut = CreateObject("Scripting.Dictionary")
    mevcut.CompareMode = vbTextCompare
    Dim sonK As Long
    Dim r As Long
    sonK = wsK.Cells(wsK.Rows.Count, 1).End(xlUp).Row
    For r = 2 To sonK
        If Len(Trim$(CStr(wsK.Cells(r, 1).Value))) > 0 Then
            mevcut(Anahtar(wsK.Cells(r, 1).Value, wsK.Cells(r, 2).Value, _
                           wsK.Cells(r, 3).Value, wsK.Cells(r, 4).Value)) = True
        End If
    Next r

    ' İcra dosyası olan personellerin listesi
    Dim adlar As Object
    Set adlar = CreateObject("Scripting.Dictionary")
    adlar.CompareMode = vbTextCompare
    Dim ad As String
    For r = 2 To ICRA_SON
        ad = Trim$(CStr(wsI.Cells(r, C_AD).Value))
        If Len(ad) > 0 And Len(Trim$(CStr(wsI.Cells(r, C_DOSYA).Value))) > 0 Then
            If Not adlar.Exists(ad) Then adlar.Add ad, True
        End If
    Next r

    Dim eklenen As Long
    Dim atlanan As Long
    Dim toplam As Double
    Dim uyari As String
    Dim k As Variant
    For Each k In adlar.Keys
        PersoneliIsle wsI, wsK, CStr(k), yil, ay, mevcut, eklenen, atlanan, toplam, uyari
    Next k

    Application.Calculate

    Dim mesaj As String
    mesaj = eklenen & " kesinti satırı eklendi." & vbCrLf & _
            "Toplam: " & Format$(toplam, "#,##0.00") & " TL"
    If atlanan > 0 Then
        mesaj = mesaj & vbCrLf & atlanan & " kayıt bu ay için zaten girilmişti, atlandı."
    End If
    If Len(uyari) > 0 Then
        mesaj = mesaj & vbCrLf & vbCrLf & "Uyarılar:" & vbCrLf & uyari
    End If
    MsgBox mesaj, vbInformation, yil & " " & ay & " kesintileri"
End Sub

Private Sub PersoneliIsle(wsI As Worksheet, wsK As Worksheet, ad As String, _
                          yil As Long, ay As String, mevcut As Object, _
                          ByRef eklenen As Long, ByRef atlanan As Long, _
                          ByRef toplam As Double, ByRef uyari As String)
    Dim sat(1 To MAKS_DOSYA) As Long
    Dim onc(1 To MAKS_DOSYA) As Double
    Dim kal(1 To MAKS_DOSYA) As Double
    Dim n As Long
    Dim r As Long

    ' Personelin açık (kalan borçlu) dosyalarını topla
    For r = 2 To ICRA_SON
        If StrComp(Trim$(CStr(wsI.Cells(r, C_AD).Value)), ad, vbTextCompare) = 0 Then
            If Sayi(wsI.Cells(r, C_KALAN).Value) > 0.005 Then
                If n < MAKS_DOSYA Then
                    n = n + 1
                    sat(n) = r
                    onc(n) = Sayi(wsI.Cells(r, C_ONCELIK).Value)
                    kal(n) = Sayi(wsI.Cells(r, C_KALAN).Value)
                End If
            End If
        End If
    Next r
    If n = 0 Then Exit Sub

    ' Öncelik sırasına göre sırala
    Dim i As Long
    Dim j As Long
    Dim tL As Long
    Dim tD As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If onc(j) < onc(i) Then
                tL = sat(i): sat(i) = sat(j): sat(j) = tL
                tD = onc(i): onc(i) = onc(j): onc(j) = tD
                tD = kal(i): kal(i) = kal(j): kal(j) = tD
            End If
        Next j
    Next i

    ' Bu ayın havuzu: aktif (en öncelikli açık) dosyanın kuralı
    Dim aktif As Long
    Dim netM As Double
    Dim havuz As Double
    aktif = sat(1)
    netM = Sayi(wsI.Cells(aktif, C_NET).Value)
    If StrComp(Trim$(CStr(wsI.Cells(aktif, C_SEKIL).Value)), "Oran", vbTextCompare) = 0 Then
        Dim pay As Double
        Dim payda As Double
        pay = Sayi(wsI.Cells(aktif, C_PAY).Value)
        payda = Sayi(wsI.Cells(aktif, C_PAYDA).Value)
        If pay <= 0 Or payda <= 0 Then
            uyari = uyari & "- " & ad & ": oran (pay/payda) eksik, atlandı." & vbCrLf
            Exit Sub
        End If
        If netM <= 0 Then
            uyari = uyari & "- " & ad & ": PERSONEL sayfasında net maaş yok, atlandı." & vbCrLf
            Exit Sub
        End If
        havuz = netM * pay / payda
    Else
        havuz = Sayi(wsI.Cells(aktif, C_SABIT).Value)
        If havuz <= 0 Then
            uyari = uyari & "- " & ad & ": sabit kesinti tutarı boş, atlandı." & vbCrLf
            Exit Sub
        End If
    End If
    If netM > 0 And havuz > netM Then havuz = netM
    havuz = Application.WorksheetFunction.Round(havuz, 2)

    ' Havuzu öncelik sırasındaki açık dosyalara dağıt
    Dim tut As Double
    Dim anah As String
    Dim hedef As Long
    For i = 1 To n
        If havuz <= 0.005 Then Exit For
        tut = havuz
        If tut > kal(i) Then tut = kal(i)
        tut = Application.WorksheetFunction.Round(tut, 2)
        anah = Anahtar(ad, wsI.Cells(sat(i), C_DOSYA).Value, yil, ay)
        If mevcut.Exists(anah) Then
            atlanan = atlanan + 1
            Exit For
        End If
        hedef = wsK.Cells(wsK.Rows.Count, 1).End(xlUp).Row + 1
        wsK.Cells(hedef, 1).Value = ad
        wsK.Cells(hedef, 2).Value = wsI.Cells(sat(i), C_DOSYA).Value
        wsK.Cells(hedef, 3).Value = yil
        wsK.Cells(hedef, 4).Value = ay
        wsK.Cells(hedef, 5).Value = tut
        wsK.Cells(hedef, 6).Value = ETIKET & " (" & Format$(Now, "dd.mm.yyyy hh:nn") & ")"
        mevcut(anah) = True
        eklenen = eklenen + 1
        toplam = toplam + tut
        havuz = havuz - tut
    Next i
End Sub

Public Sub SECILI_AYI_GERI_AL()
    Dim wsK As Worksheet
    Dim wsP As Worksheet
    Set wsK = ThisWorkbook.Worksheets(SAYFA_KES)
    Set wsP = ThisWorkbook.Worksheets(SAYFA_PANEL)

    Dim yil As Long
    Dim ay As String
    If Not IsNumeric(wsP.Range("C3").Value) Then
        MsgBox "PANEL sayfasında C3 hücresine yıl girin.", vbExclamation
        Exit Sub
    End If
    yil = CLng(wsP.Range("C3").Value)
    ay = Trim$(CStr(wsP.Range("C4").Value))
    If Not AyGecerli(ay) Then
        MsgBox "PANEL sayfasında C4 hücresinden ayı seçin.", vbExclamation
        Exit Sub
    End If
    If MsgBox(yil & " " & ay & " dönemine ait OTOMATİK işaretli kesinti satırları silinecek." & _
              vbCrLf & "Elle girdiğiniz satırlara dokunulmaz. Devam edilsin mi?", _
              vbYesNo + vbQuestion, "Geri al") <> vbYes Then Exit Sub

    Dim r As Long
    Dim silinen As Long
    Dim sonK As Long
    sonK = wsK.Cells(wsK.Rows.Count, 1).End(xlUp).Row
    For r = sonK To 2 Step -1
        If Trim$(CStr(wsK.Cells(r, 3).Value)) = CStr(yil) Then
            If StrComp(Trim$(CStr(wsK.Cells(r, 4).Value)), ay, vbTextCompare) = 0 Then
                If InStr(1, CStr(wsK.Cells(r, 6).Value), ETIKET, vbTextCompare) = 1 Then
                    wsK.Rows(r).Delete
                    silinen = silinen + 1
                End If
            End If
        End If
    Next r
    Application.Calculate
    MsgBox silinen & " satır silindi.", vbInformation, "Geri al"
End Sub

Private Function Sayi(v As Variant) As Double
    If IsNumeric(v) Then Sayi = CDbl(v)
End Function

Private Function Anahtar(ad As Variant, dosya As Variant, yil As Variant, ay As Variant) As String
    Anahtar = Trim$(CStr(ad)) & "|" & Trim$(CStr(dosya)) & "|" & _
              Trim$(CStr(yil)) & "|" & Trim$(CStr(ay))
End Function

Private Function AyGecerli(ay As String) As Boolean
    Select Case ay
        Case "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", _
             "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"
            AyGecerli = True
    End Select
End Function
