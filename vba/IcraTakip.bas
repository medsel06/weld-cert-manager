Attribute VB_Name = "IcraTakip"
Option Explicit

' =====================================================================
'  PERSONEL ÝCRA TAKÝP MAKROLARI
'  Kurulum: Excel'de Alt+F11 > Dosya > Dosya Al... > IcraTakip.bas
'           Sonra dosyayý .xlsm olarak kaydedin (KULLANIM sayfasýna bakýn).
'
'  BUTONLARI_KUR       : PANEL sayfasýna butonlarý ekler (bir kez çalýþtýrýn).
'  KESINTILERI_KAYDET  : PANEL'de seçili yýl/ay için kesintileri hesaplar,
'                        AYLIK KESÝNTÝLER sayfasýna iþler. Kural:
'                        havuz = net maaþ x pay/payda (veya sabit tutar);
'                        öncelik sýrasýndaki açýk dosyalara daðýtýlýr,
'                        kalan borç aþýlmaz, dosya kapanýrsa artan tutar
'                        ayný ay sýradaki dosyaya taþar. Ayný ay için
'                        mükerrer kayýt oluþturmaz.
'  SECILI_AYI_GERI_AL  : Seçili ayýn OTOMATÝK iþaretli satýrlarýný siler.
' =====================================================================

Private Const SAYFA_ICRA As String = "ÝCRA DOSYALARI"
Private Const SAYFA_KES As String = "AYLIK KESÝNTÝLER"
Private Const SAYFA_PANEL As String = "PANEL"
Private Const ICRA_SON As Long = 300
Private Const ETIKET As String = "OTOMATÝK"
Private Const MAKS_DOSYA As Long = 64

' --- ÝCRA DOSYALARI sayfasý kolon numaralarý ---
Private Const C_AD As Long = 1        ' A  Personel Adý Soyadý
Private Const C_NET As Long = 3       ' C  Net Maaþ
Private Const C_DOSYA As Long = 5     ' E  Dosya Numarasý
Private Const C_ONCELIK As Long = 6   ' F  Öncelik Sýrasý
Private Const C_SEKIL As Long = 10    ' J  Kesinti Þekli
Private Const C_PAY As Long = 11      ' K  Oran - Pay
Private Const C_PAYDA As Long = 12    ' L  Oran - Payda
Private Const C_SABIT As Long = 13    ' M  Sabit Aylýk Kesinti
Private Const C_KALAN As Long = 18    ' R  Kalan Borç

Public Sub BUTONLARI_KUR()
    Dim ws As Worksheet
    Dim b As Button
    Set ws = ThisWorkbook.Worksheets(SAYFA_PANEL)
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    Set b = ws.Buttons.Add(ws.Range("B11").Left + 4, ws.Range("B11").Top + 6, 270, 40)
    b.Caption = "KESÝNTÝLERÝ HESAPLA ve KAYDET"
    b.OnAction = "KESINTILERI_KAYDET"
    b.Font.Bold = True
    b.Font.Size = 11
    Set b = ws.Buttons.Add(ws.Range("B11").Left + 4, ws.Range("B11").Top + 56, 270, 26)
    b.Caption = "Seçili ayýn OTOMATÝK kayýtlarýný GERÝ AL"
    b.OnAction = "SECILI_AYI_GERI_AL"
    b.Font.Size = 9
    MsgBox "Butonlar PANEL sayfasýna eklendi." & vbCrLf & _
           "Artýk PANEL'den yýl/ay seçip butonla kaydedebilirsiniz.", _
           vbInformation, "Kurulum tamamlandý"
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
        MsgBox "PANEL sayfasýnda C3 hücresine yýl girin (örn. 2026).", vbExclamation
        Exit Sub
    End If
    yil = CLng(wsP.Range("C3").Value)
    ay = Trim$(CStr(wsP.Range("C4").Value))
    If yil < 2000 Or yil > 2100 Then
        MsgBox "Geçerli bir yýl girin (2000 - 2100).", vbExclamation
        Exit Sub
    End If
    If Not AyGecerli(ay) Then
        MsgBox "PANEL sayfasýnda C4 hücresinden ayý seçin.", vbExclamation
        Exit Sub
    End If

    Application.Calculate

    ' Mevcut kayýtlarý oku (ayný ay için mükerrer giriþi önler)
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

    ' Ýcra dosyasý olan personellerin listesi
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
    mesaj = eklenen & " kesinti satýrý eklendi." & vbCrLf & _
            "Toplam: " & Format$(toplam, "#,##0.00") & " TL"
    If atlanan > 0 Then
        mesaj = mesaj & vbCrLf & atlanan & " kayýt bu ay için zaten girilmiþti, atlandý."
    End If
    If Len(uyari) > 0 Then
        mesaj = mesaj & vbCrLf & vbCrLf & "Uyarýlar:" & vbCrLf & uyari
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

    ' Personelin açýk (kalan borçlu) dosyalarýný topla
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

    ' Öncelik sýrasýna göre sýrala
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

    ' Bu ayýn havuzu: aktif (en öncelikli açýk) dosyanýn kuralý
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
            uyari = uyari & "- " & ad & ": oran (pay/payda) eksik, atlandý." & vbCrLf
            Exit Sub
        End If
        If netM <= 0 Then
            uyari = uyari & "- " & ad & ": PERSONEL sayfasýnda net maaþ yok, atlandý." & vbCrLf
            Exit Sub
        End If
        havuz = netM * pay / payda
    Else
        havuz = Sayi(wsI.Cells(aktif, C_SABIT).Value)
        If havuz <= 0 Then
            uyari = uyari & "- " & ad & ": sabit kesinti tutarý boþ, atlandý." & vbCrLf
            Exit Sub
        End If
    End If
    If netM > 0 And havuz > netM Then havuz = netM
    havuz = Application.WorksheetFunction.Round(havuz, 2)

    ' Havuzu öncelik sýrasýndaki açýk dosyalara daðýt
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
        MsgBox "PANEL sayfasýnda C3 hücresine yýl girin.", vbExclamation
        Exit Sub
    End If
    yil = CLng(wsP.Range("C3").Value)
    ay = Trim$(CStr(wsP.Range("C4").Value))
    If Not AyGecerli(ay) Then
        MsgBox "PANEL sayfasýnda C4 hücresinden ayý seçin.", vbExclamation
        Exit Sub
    End If
    If MsgBox(yil & " " & ay & " dönemine ait OTOMATÝK iþaretli kesinti satýrlarý silinecek." & _
              vbCrLf & "Elle girdiðiniz satýrlara dokunulmaz. Devam edilsin mi?", _
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
    MsgBox silinen & " satýr silindi.", vbInformation, "Geri al"
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
        Case "Ocak", "Þubat", "Mart", "Nisan", "Mayýs", "Haziran", _
             "Temmuz", "Aðustos", "Eylül", "Ekim", "Kasým", "Aralýk"
            AyGecerli = True
    End Select
End Function
