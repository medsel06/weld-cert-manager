Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    UYGULAMA_BASLAT
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    UYGULAMA_NORMAL_GORUNUM
End Sub
