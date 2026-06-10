Option Explicit

Private mSecildi As Boolean
Private mTarih As Date

Public Function TarihSec(Optional ByVal baslangic As Date = 0) As Variant
    If baslangic = 0 Then baslangic = Date
    mSecildi = False
    TarihiYukle baslangic
    Me.Show vbModal
    If mSecildi Then
        TarihSec = mTarih
    Else
        TarihSec = Empty
    End If
End Function

Private Sub UserForm_Initialize()
    FormStil.Uygula Me
End Sub

Private Sub TarihiYukle(ByVal tarih As Date)
    Dim i As Long
    cboGun.Clear
    For i = 1 To 31
        cboGun.AddItem CStr(i)
    Next i

    cboAy.Clear
    UI_AyComboYukle cboAy

    cboYil.Clear
    For i = Year(Date) - 10 To Year(Date) + 10
        cboYil.AddItem CStr(i)
    Next i

    cboGun.Value = CStr(Day(tarih))
    cboAy.ListIndex = Month(tarih) - 1
    cboYil.Value = CStr(Year(tarih))
    lblSecili.Caption = Format$(tarih, "dd.mm.yyyy")
End Sub

Private Sub cboGun_Change()
    SecimiGuncelle
End Sub

Private Sub cboAy_Change()
    SecimiGuncelle
End Sub

Private Sub cboYil_Change()
    SecimiGuncelle
End Sub

Private Sub SecimiGuncelle()
    Dim t As Date
    If SecimTarihi(t) Then
        lblSecili.Caption = Format$(t, "dd.mm.yyyy")
    Else
        lblSecili.Caption = "Geçersiz tarih"
    End If
End Sub

Private Function SecimTarihi(ByRef tarih As Date) As Boolean
    On Error GoTo Hata
    Dim gun As Long
    Dim ay As Long
    Dim yil As Long
    gun = CLng(Val(cboGun.Value))
    ay = cboAy.ListIndex + 1
    yil = CLng(Val(cboYil.Value))
    If gun < 1 Or ay < 1 Or yil < 1900 Then GoTo Hata
    tarih = DateSerial(yil, ay, gun)
    If Day(tarih) <> gun Or Month(tarih) <> ay Or Year(tarih) <> yil Then GoTo Hata
    SecimTarihi = True
    Exit Function
Hata:
    SecimTarihi = False
End Function

Private Sub cmdBugun_Click()
    TarihiYukle Date
End Sub

Private Sub cmdSec_Click()
    Dim t As Date
    If Not SecimTarihi(t) Then
        MsgBox "Geçerli bir tarih seçin.", vbExclamation, "Tarih"
        Exit Sub
    End If
    mTarih = t
    mSecildi = True
    Hide
End Sub

Private Sub cmdIptal_Click()
    mSecildi = False
    Hide
End Sub
