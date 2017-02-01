Private Sub Workbook_BeforeClose(Cancel As Boolean)
Application.OnKey "+^v", ""
End Sub

Private Sub Workbook_Open()
Application.OnKey "+^v", "copyduoishow.xlam!copyshow"
End Sub