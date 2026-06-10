Option Explicit

Private Sub UserForm_Initialize()
    FormStil.Uygula Me
    txtYil.Value = CStr(Year(Date))
    UI_AyComboYukle cboAy
    cboAy.Value = UI_AyAdi(Date)
    OnizlemeGuncelle
End Sub

Private Sub txtYil_Change()
    OnizlemeGuncelle
End Sub

Private Sub cboAy_Change()
    OnizlemeGuncelle
End Sub

Private Sub OnizlemeGuncelle()
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long
    Dim toplam As Double
    Application.Calculate
    Set ws = ThisWorkbook.Worksheets("İCRA DOSYALARI")
    For r = 2 To 300
        If Len(Trim$(CStr(ws.Cells(r, 5).Value))) > 0 Then
            If IsNumeric(ws.Cells(r, 14).Value) Then toplam = toplam + ws.Cells(r, 14).Value
        End If
    Next r
    lblOnizleme.Caption = "Bu dönem kesilecek (önizleme): " & Format$(toplam, "#,##0.00") & " TL"
End Sub

Private Function DonemiPaneleYaz() As Boolean
    Dim ws As Worksheet
    If Not IsNumeric(txtYil.Value) Then
        MsgBox "Geçerli bir yıl girin (örn. 2026).", vbExclamation, "Dönem"
        Exit Function
    End If
    If Len(Trim$(cboAy.Value)) = 0 Then
        MsgBox "Ay seçin.", vbExclamation, "Dönem"
        Exit Function
    End If
    Set ws = ThisWorkbook.Worksheets("PANEL")
    ws.Range("C3").Value = CLng(txtYil.Value)
    ws.Range("C4").Value = Trim$(cboAy.Value)
    DonemiPaneleYaz = True
End Function

Private Sub cmdHesapla_Click()
    If Not DonemiPaneleYaz() Then Exit Sub
    If MsgBox(Trim$(txtYil.Value) & " " & Trim$(cboAy.Value) & _
              " dönemi için otomatik kesintiler kaydedilecek. Devam edilsin mi?", _
              vbYesNo + vbQuestion, "Ay sonu kesinti") <> vbYes Then Exit Sub
    KESINTILERI_KAYDET
    OnizlemeGuncelle
End Sub

Private Sub cmdGeriAl_Click()
    If Not DonemiPaneleYaz() Then Exit Sub
    SECILI_AYI_GERI_AL
    OnizlemeGuncelle
End Sub

Private Sub cmdKapat_Click()
    Unload Me
End Sub
