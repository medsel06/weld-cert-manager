Attribute VB_Name = "FormStil"
Option Explicit

' =====================================================================
'  FORM STÝL MOTORU  (DPI ölçekleme + kurumsal tema)
'  ---------------------------------------------------------------
'  AMAÇ: UserForm'lar yüksek çözünürlüklü (4K / %150 ölçekli) laptop
'        ekranlarýnda küçük görünür; çünkü VBA formlarý DPI'a duyarlý
'        deðildir. Bu motor, açýlan her formu ekranýn DPI'ýna göre
'        ORANTILI büyütür (büyük monitörde olduðu gibi kalýr, laptopta
'        okunur hale gelir), tüm yazýlarý Segoe UI + okunur minimum
'        boyuta çeker ve kurumsal renk temasýný uygular.
'
'  KURULUM:
'    1) Bu modülü içe aktarýn (Alt+F11 > Dosya > Dosya Al > FormStil.bas).
'    2) 4 formun her birinin UserForm_Initialize olayýnýn ÝLK satýrýna
'       þunu ekleyin:
'                 FormStil.Uygula Me
'       (frmAnaMenu, frmPersonelGiris, frmIcraGiris, frmKesintiGiris)
'
'  AYAR: Hâlâ küçük/büyük gelirse aþaðýdaki OLCEK_ZORLA deðerini elle
'        verin (0 = otomatik DPI). Örn. 1.3 -> %30 daha büyük.
' =====================================================================

' 0 = otomatik (ekran DPI'ýndan hesaplar). Elle zorlamak için 1.0 - 2.0 arasý verin.
Private Const OLCEK_ZORLA As Double = 0
' Hiçbir zaman küçültme (monitör görünümünü bozma):
Private Const OLCEK_TABAN As Double = 1#
' Üst sýnýr (aþýrý büyümeyi engelle):
Private Const OLCEK_TAVAN As Double = 2#
' Giriþ kutusu/buton/etiket için hedef minimum punto (DPI=100 iken):
Private Const MIN_PUNTO As Double = 11#
Private Const YAZI_TIPI As String = "Segoe UI"

' --- Renkler tema bölümünde RGB() ile verilir; sýk kullanýlan beyaz: ---
Private Const RENK_BEYAZ As Long = 16777215     ' RGB(255,255,255)

#If VBA7 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
#Else
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
#End If
Private Const LOGPIXELSX As Long = 88

' =====================================================================
'  TEK GÝRÝÞ NOKTASI - her formun Initialize olayýndan çaðrýlýr.
' =====================================================================
Public Sub Uygula(frm As Object)
    On Error Resume Next
    ' Ayný form iki kez ölçeklenmesin
    If InStr(1, frm.Tag, "STIL", vbTextCompare) > 0 Then Exit Sub
    frm.Tag = frm.Tag & " STIL"

    Dim olcek As Double
    olcek = OlcekFaktoru(frm)

    OlcekleVeBicimle frm, olcek
    TemaUygula frm
End Sub

' ---------------------------------------------------------------------
'  Ölçek faktörü = max(DPI faktörü, minimum-punto faktörü, taban)
'  -> hem yüksek-DPI laptop sorununu hem de tasarýmdaki küçük fontu çözer.
' ---------------------------------------------------------------------
Private Function OlcekFaktoru(frm As Object) As Double
    Dim dpiF As Double
    If OLCEK_ZORLA > 0 Then
        dpiF = OLCEK_ZORLA
    Else
        dpiF = EkranDpi() / 96#
    End If

    ' Giriþ kontrolleri içindeki en küçük fontu hedef puntoya çýkaracak faktör
    Dim enKucuk As Double
    enKucuk = 999
    Dim ctrl As Object
    For Each ctrl In frm.Controls
        Dim tn As String
        tn = TypeName(ctrl)
        If tn = "TextBox" Or tn = "ComboBox" Or tn = "CommandButton" Or tn = "Label" Then
            If ctrl.Font.Size > 0 And ctrl.Font.Size < enKucuk Then enKucuk = ctrl.Font.Size
        End If
    Next ctrl
    Dim icerikF As Double
    icerikF = 1#
    If enKucuk > 0 And enKucuk < 900 Then icerikF = MIN_PUNTO / enKucuk

    Dim f As Double
    f = dpiF
    If icerikF > f Then f = icerikF
    If f < OLCEK_TABAN Then f = OLCEK_TABAN
    If f > OLCEK_TAVAN Then f = OLCEK_TAVAN
    OlcekFaktoru = f
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

' ---------------------------------------------------------------------
'  Formu ve tüm kontrolleri tek faktörle orantýlý büyüt (taþma olmaz).
' ---------------------------------------------------------------------
Private Sub OlcekleVeBicimle(frm As Object, ByVal f As Double)
    On Error Resume Next
    If f <= 1.0001 Then Exit Sub   ' monitörde deðiþiklik yok

    Dim ctrl As Object
    For Each ctrl In frm.Controls
        ctrl.Left = ctrl.Left * f
        ctrl.Top = ctrl.Top * f
        ctrl.Width = ctrl.Width * f
        ctrl.Height = ctrl.Height * f
        If ctrl.Font.Size > 0 Then ctrl.Font.Size = Yuvarla(ctrl.Font.Size * f)
    Next ctrl

    frm.Width = frm.Width * f
    frm.Height = frm.Height * f
    If Not frm.Font Is Nothing Then frm.Font.Size = Yuvarla(frm.Font.Size * f)
End Sub

Private Function Yuvarla(ByVal v As Double) As Double
    Yuvarla = Int(v * 2 + 0.5) / 2   ' 0.5 puntoya yuvarla
End Function

' ---------------------------------------------------------------------
'  Kurumsal tema: yazý tipi, renkler, rol bazlý buton/etiket biçimi.
' ---------------------------------------------------------------------
Private Sub TemaUygula(frm As Object)
    On Error Resume Next
    frm.BackColor = RGB(247, 249, 252)

    Dim ctrl As Object
    For Each ctrl In frm.Controls
        ctrl.Font.Name = YAZI_TIPI
        Dim ad As String
        ad = ctrl.Name
        Select Case TypeName(ctrl)

            Case "CommandButton"
                ctrl.Font.Bold = True
                If AdIcerir(ad, "Kaydet") Then
                    ButonRenk ctrl, RGB(33, 115, 70), RENK_BEYAZ
                ElseIf AdIcerir(ad, "Kapat") Then
                    ButonRenk ctrl, RGB(192, 57, 43), RENK_BEYAZ
                ElseIf AdIcerir(ad, "Temizle") Or AdIcerir(ad, "GeriAl") Then
                    ButonRenk ctrl, RGB(127, 140, 141), RENK_BEYAZ
                ElseIf AdIcerir(ad, "Hesapla") Then
                    ButonRenk ctrl, RGB(211, 84, 0), RENK_BEYAZ
                Else
                    ' menü / genel butonlar - marka rengi
                    ButonRenk ctrl, RGB(31, 78, 120), RENK_BEYAZ
                End If

            Case "TextBox", "ComboBox"
                ctrl.BackColor = RENK_BEYAZ
                ctrl.ForeColor = RGB(30, 30, 30)
                ctrl.BorderStyle = 1            ' fmBorderStyleSingle
                ctrl.SpecialEffect = 0          ' fmSpecialEffectFlat

            Case "Label"
                If AdIcerir(ad, "Title") Then
                    ' Üst baþlýk bandý
                    ctrl.BackColor = RGB(31, 78, 120)
                    ctrl.ForeColor = RENK_BEYAZ
                    ctrl.Font.Bold = True
                    ctrl.Font.Size = ctrl.Font.Size + 3
                    ctrl.TextAlign = 2          ' fmTextAlignCenter
                    ctrl.Left = 6
                    ctrl.Width = frm.InsideWidth - 12
                ElseIf AdIcerir(ad, "Bolum") Then
                    ' Bölüm baþlýklarý
                    ctrl.ForeColor = RGB(31, 78, 120)
                    ctrl.Font.Bold = True
                    ctrl.BackColor = RGB(225, 232, 241)
                ElseIf AdIcerir(ad, "Not") Or AdIcerir(ad, "Info") Or AdIcerir(ad, "Alt") _
                       Or AdIcerir(ad, "PersonelBilgi") Then
                    ctrl.ForeColor = RGB(90, 90, 90)
                    ctrl.Font.Italic = True
                Else
                    ctrl.ForeColor = RGB(45, 45, 45)
                End If
        End Select
    Next ctrl
End Sub

Private Sub ButonRenk(ctrl As Object, ByVal zemin As Long, ByVal yazi As Long)
    On Error Resume Next
    ctrl.BackColor = zemin
    ctrl.ForeColor = yazi
End Sub

Private Function AdIcerir(ByVal ad As String, ByVal parca As String) As Boolean
    AdIcerir = (InStr(1, ad, parca, vbTextCompare) > 0)
End Function
