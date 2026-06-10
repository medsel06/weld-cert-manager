Attribute VB_Name = "FormStil"
Option Explicit

' =====================================================================
'  FORM STÝL MOTORU  (DPI ölçekleme + ekrana sýðdýrma + kurumsal tema)
'  ---------------------------------------------------------------
'  AMAÇ: UserForm'lar yüksek-DPI (4K / %150) laptoplarda küçük çizilir
'        (VBA formlarý DPI'a duyarlý deðildir). Bu motor her formu:
'          1) ÝÇERÝÐE GÖRE boyutlandýrýr  -> iç taþma / yarým kontrol /
'             gereksiz boþluk olmaz.
'          2) Ekranýn DPI'ýna göre orantýlý büyütür (monitör görünümü
'             korunur, laptopta okunur olur).
'          3) EKRANI AÞMAYACAK þekilde sýnýrlar; sýðmazsa kaydýrma
'             çubuðu açar -> hiçbir kontrol ekran dýþýnda kalmaz.
'          4) Kurumsal renk temasý + Segoe UI uygular, formu ortalar.
'
'  KURULUM: FormStil.bas'i içe aktarýn; 4 formun UserForm_Initialize
'           olayýnýn ÝLK satýrýna þunu ekleyin:  FormStil.Uygula Me
'           (FORM_STIL_KURULUM.md dosyasýna bakýn.)
'
'  AYAR: OLCEK_ZORLA = 0 (otomatik DPI). Elle istiyorsanýz 1.0-2.0 verin.
' =====================================================================

Private Const OLCEK_ZORLA As Double = 0       ' 0 = otomatik DPI
Private Const OLCEK_TABAN As Double = 1#       ' istek faktörü en az 1 (küçültmeyi yalnýz ekran sýnýrý yapar)
Private Const OLCEK_TAVAN As Double = 1.8       ' aþýrý büyütmeyi engelle
Private Const OLCEK_DIP As Double = 0.6         ' ekrana sýðdýrýrken bundan fazla küçültme (altýnda scrollbar)
Private Const MIN_PUNTO As Double = 11#          ' hedef minimum punto (DPI=100 iken)
Private Const KENAR As Double = 14#              ' form iç kenar boþluðu (punto)
Private Const EKRAN_PAY_X As Double = 0.94       ' ekran geniþliðinin kullanýlabilir oraný
Private Const EKRAN_PAY_Y As Double = 0.9        ' ekran yüksekliðinin kullanýlabilir oraný
Private Const YAZI_TIPI As String = "Segoe UI"
Private Const RENK_BEYAZ As Long = 16777215

#If VBA7 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If
Private Const LOGPIXELSX As Long = 88
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1

' =====================================================================
'  TEK GÝRÝÞ NOKTASI - her formun Initialize olayýndan çaðrýlýr.
' =====================================================================
Public Sub Uygula(frm As Object)
    On Error Resume Next
    If InStr(1, frm.Tag, "STIL", vbTextCompare) > 0 Then Exit Sub
    frm.Tag = frm.Tag & " STIL"

    Dim dpi As Double
    dpi = EkranDpi()

    ' 1) Ölçeksiz içerik kutusu (tüm kontrolleri saran sýnýr)
    Dim baseR As Double
    Dim baseB As Double
    IcerikKutusu frm, baseR, baseB
    If baseR < 1 Or baseB < 1 Then Exit Sub

    ' 2) Ýstenen büyütme (DPI + minimum punto)
    Dim istenen As Double
    istenen = IstenenOlcek(frm, dpi)

    ' 3) Ekran sýnýrý (punto): formu buraya sýðdýr
    Dim ekrW As Double
    Dim ekrH As Double
    ekrW = EkranNokta(SM_CXSCREEN, dpi)
    ekrH = EkranNokta(SM_CYSCREEN, dpi)
    Dim kullW As Double
    Dim kullH As Double
    kullW = ekrW * EKRAN_PAY_X
    kullH = ekrH * EKRAN_PAY_Y

    Dim ekranMax As Double
    ekranMax = MinD((kullW - KENAR) / baseR, (kullH - KENAR) / baseB)

    Dim final As Double
    final = istenen
    If final > ekranMax Then final = ekranMax     ' ekrana sýðdýrmak için gerekirse küçült
    If final > OLCEK_TAVAN Then final = OLCEK_TAVAN
    If final < OLCEK_DIP Then final = OLCEK_DIP    ' bundan fazla küçülme; gerisini scrollbar halleder

    ' 4) Uygula
    OlcekleVeBicimle frm, final
    FormaSigdir frm, baseR * final, baseB * final, kullW, kullH
    TemaUygula frm
    Ortala frm, ekrW, ekrH
End Sub

' ---------------------------------------------------------------------
'  Ölçeksiz içerik kutusu: max(Left+Width), max(Top+Height)
' ---------------------------------------------------------------------
Private Sub IcerikKutusu(frm As Object, ByRef r As Double, ByRef b As Double)
    On Error Resume Next
    Dim c As Object
    Dim x As Double
    Dim y As Double
    r = 0: b = 0
    For Each c In frm.Controls
        x = c.Left + c.Width
        y = c.Top + c.Height
        If x > r Then r = x
        If y > b Then b = y
    Next c
    If r <= 0 Then r = frm.InsideWidth
    If b <= 0 Then b = frm.InsideHeight
End Sub

' ---------------------------------------------------------------------
'  Ýstenen ölçek = max(DPI/96, MIN_PUNTO/en-küçük-font, OLCEK_TABAN)
' ---------------------------------------------------------------------
Private Function IstenenOlcek(frm As Object, ByVal dpi As Double) As Double
    Dim dpiF As Double
    If OLCEK_ZORLA > 0 Then
        dpiF = OLCEK_ZORLA
    Else
        dpiF = dpi / 96#
    End If

    Dim enKucuk As Double
    enKucuk = 999
    Dim c As Object
    Dim tn As String
    For Each c In frm.Controls
        tn = TypeName(c)
        If tn = "TextBox" Or tn = "ComboBox" Or tn = "CommandButton" Or tn = "Label" Then
            If c.Font.Size > 0 And c.Font.Size < enKucuk Then enKucuk = c.Font.Size
        End If
    Next c
    Dim icerikF As Double
    icerikF = 1#
    If enKucuk > 0 And enKucuk < 900 Then icerikF = MIN_PUNTO / enKucuk

    Dim f As Double
    f = dpiF
    If icerikF > f Then f = icerikF
    If f < OLCEK_TABAN Then f = OLCEK_TABAN
    IstenenOlcek = f
End Function

Private Function EkranDpi() As Double
    On Error GoTo Varsayilan
    #If VBA7 Then
        Dim hdc As LongPtr
    #Else
        Dim hdc As Long
    #End If
    hdc = GetDC(0)
    If hdc <> 0 Then
        EkranDpi = GetDeviceCaps(hdc, LOGPIXELSX)
        ReleaseDC 0, hdc
    End If
    If EkranDpi < 72 Then EkranDpi = 96
    Exit Function
Varsayilan:
    EkranDpi = 96
End Function

' Ekran boyutunu punto cinsine çevir (mantýksal piksel 96 tabanlýdýr)
Private Function EkranNokta(ByVal eksen As Long, ByVal dpi As Double) As Double
    On Error GoTo Varsayilan
    Dim px As Long
    px = GetSystemMetrics(eksen)
    If px < 200 Then GoTo Varsayilan
    EkranNokta = px * 72# / 96#
    Exit Function
Varsayilan:
    If eksen = SM_CYSCREEN Then EkranNokta = 768 * 0.75 Else EkranNokta = 1366 * 0.75
End Function

' ---------------------------------------------------------------------
'  Tüm kontrolleri tek faktörle ölçekle (oran korunur, iç taþma olmaz)
' ---------------------------------------------------------------------
Private Sub OlcekleVeBicimle(frm As Object, ByVal f As Double)
    On Error Resume Next
    If Abs(f - 1) < 0.0001 Then Exit Sub
    Dim c As Object
    For Each c In frm.Controls
        c.Left = c.Left * f
        c.Top = c.Top * f
        c.Width = c.Width * f
        c.Height = c.Height * f
        If c.Font.Size > 0 Then c.Font.Size = Yuvarla(c.Font.Size * f)
    Next c
    If Not frm.Font Is Nothing Then frm.Font.Size = Yuvarla(frm.Font.Size * f)
End Sub

Private Function Yuvarla(ByVal v As Double) As Double
    Yuvarla = Int(v * 2 + 0.5) / 2
End Function

Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

' ---------------------------------------------------------------------
'  Formu içeriðe göre boyutla; ekraný aþarsa kýrp + kaydýrma çubuðu aç.
' ---------------------------------------------------------------------
Private Sub FormaSigdir(frm As Object, ByVal icerikW As Double, ByVal icerikH As Double, _
                        ByVal kullW As Double, ByVal kullH As Double)
    On Error Resume Next
    Dim gerW As Double
    Dim gerH As Double
    gerW = icerikW + KENAR
    gerH = icerikH + KENAR

    Dim icW As Double
    Dim icH As Double
    icW = gerW
    icH = gerH
    If icW > kullW Then icW = kullW
    If icH > kullH Then icH = kullH

    frm.InsideWidth = icW
    frm.InsideHeight = icH

    If gerW > icW + 1 Or gerH > icH + 1 Then
        frm.ScrollBars = 3                 ' fmScrollBarsBoth
        frm.KeepScrollBarsVisible = 0      ' yalnýz gerekince göster
        frm.ScrollWidth = gerW
        frm.ScrollHeight = gerH
    End If
End Sub

' ---------------------------------------------------------------------
'  Formu ekranda ortala
' ---------------------------------------------------------------------
Private Sub Ortala(frm As Object, ByVal ekrW As Double, ByVal ekrH As Double)
    On Error Resume Next
    frm.StartUpPosition = 0                ' Manuel
    frm.Left = (ekrW - frm.Width) / 2
    frm.Top = (ekrH - frm.Height) / 2
    If frm.Left < 0 Then frm.Left = 0
    If frm.Top < 0 Then frm.Top = 0
End Sub

' ---------------------------------------------------------------------
'  Kurumsal tema: yazý tipi, renkler, rol bazlý buton/etiket biçimi.
' ---------------------------------------------------------------------
Private Sub TemaUygula(frm As Object)
    On Error Resume Next
    frm.BackColor = RGB(247, 249, 252)

    Dim c As Object
    Dim ad As String
    For Each c In frm.Controls
        c.Font.Name = YAZI_TIPI
        ad = c.Name
        Select Case TypeName(c)

            Case "CommandButton"
                c.Font.Bold = True
                If AdIcerir(ad, "Kaydet") Then
                    ButonRenk c, RGB(33, 115, 70), RENK_BEYAZ
                ElseIf AdIcerir(ad, "Kapat") Then
                    ButonRenk c, RGB(192, 57, 43), RENK_BEYAZ
                ElseIf AdIcerir(ad, "Temizle") Or AdIcerir(ad, "GeriAl") Then
                    ButonRenk c, RGB(127, 140, 141), RENK_BEYAZ
                ElseIf AdIcerir(ad, "Hesapla") Then
                    ButonRenk c, RGB(211, 84, 0), RENK_BEYAZ
                Else
                    ButonRenk c, RGB(31, 78, 120), RENK_BEYAZ
                End If

            Case "TextBox", "ComboBox"
                c.BackColor = RENK_BEYAZ
                c.ForeColor = RGB(30, 30, 30)
                c.BorderStyle = 1            ' fmBorderStyleSingle
                c.SpecialEffect = 0          ' fmSpecialEffectFlat

            Case "Label"
                If AdIcerir(ad, "Title") Then
                    c.BackColor = RGB(31, 78, 120)
                    c.ForeColor = RENK_BEYAZ
                    c.Font.Bold = True
                    c.Font.Size = c.Font.Size + 3
                    c.TextAlign = 2          ' fmTextAlignCenter
                    c.Left = 6
                    c.Width = frm.InsideWidth - 12
                ElseIf AdIcerir(ad, "Bolum") Then
                    c.ForeColor = RGB(31, 78, 120)
                    c.Font.Bold = True
                    c.BackColor = RGB(225, 232, 241)
                ElseIf AdIcerir(ad, "Not") Or AdIcerir(ad, "Info") Or AdIcerir(ad, "Alt") _
                       Or AdIcerir(ad, "PersonelBilgi") Then
                    c.ForeColor = RGB(90, 90, 90)
                    c.Font.Italic = True
                Else
                    c.ForeColor = RGB(45, 45, 45)
                End If
        End Select
    Next c
End Sub

Private Sub ButonRenk(c As Object, ByVal zemin As Long, ByVal yazi As Long)
    On Error Resume Next
    c.BackColor = zemin
    c.ForeColor = yazi
End Sub

Private Function AdIcerir(ByVal ad As String, ByVal parca As String) As Boolean
    AdIcerir = (InStr(1, ad, parca, vbTextCompare) > 0)
End Function
