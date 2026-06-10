Option Explicit

Private Sub UserForm_Initialize()
    FormStil.Uygula Me
    OzetYukle
End Sub

Private Sub OzetYukle()
    Dim wsP As Worksheet
    Dim wsI As Worksheet
    Dim r As Long
    Dim i As Long
    Dim ix As Long
    Dim ad As String
    Application.Calculate
    Set wsP = ThisWorkbook.Worksheets("PERSONEL")
    Set wsI = ThisWorkbook.Worksheets("İCRA DOSYALARI")

    lstOzet.Clear
    For r = 2 To 300
        ad = Trim$(CStr(wsP.Cells(r, 1).Value))
        If Len(ad) > 0 Then
            Dim adet As Long
            Dim icra As Double
            Dim tahsil As Double
            Dim kalan As Double
            Dim odenen As Double
            Dim bekleyen As Double
            adet = 0: icra = 0: tahsil = 0: kalan = 0: odenen = 0: bekleyen = 0
            For i = 2 To 300
                If StrComp(Trim$(CStr(wsI.Cells(i, 1).Value)), ad, vbTextCompare) = 0 Then
                    If Len(Trim$(CStr(wsI.Cells(i, 5).Value))) > 0 Then
                        adet = adet + 1
                        icra = icra + Sayi(wsI.Cells(i, 9).Value)
                        tahsil = tahsil + Sayi(wsI.Cells(i, 17).Value)
                        kalan = kalan + Sayi(wsI.Cells(i, 18).Value)
                        odenen = odenen + Sayi(wsI.Cells(i, 23).Value)
                        bekleyen = bekleyen + Sayi(wsI.Cells(i, 24).Value)
                    End If
                End If
            Next i
            lstOzet.AddItem ad
            ix = lstOzet.ListCount - 1
            lstOzet.List(ix, 1) = Format$(Sayi(wsP.Cells(r, 3).Value), "#,##0.00")
            lstOzet.List(ix, 2) = CStr(adet)
            lstOzet.List(ix, 3) = Format$(icra, "#,##0.00")
            lstOzet.List(ix, 4) = Format$(tahsil, "#,##0.00")
            lstOzet.List(ix, 5) = Format$(kalan, "#,##0.00")
            lstOzet.List(ix, 6) = Format$(odenen, "#,##0.00")
            lstOzet.List(ix, 7) = Format$(bekleyen, "#,##0.00")
        End If
    Next r
    lblSayac.Caption = lstOzet.ListCount & " personel"
End Sub

Private Function Sayi(ByVal v As Variant) As Double
    If IsNumeric(v) Then Sayi = CDbl(v)
End Function

Private Sub cmdYenile_Click()
    OzetYukle
End Sub

Private Sub cmdKapat_Click()
    Unload Me
End Sub
