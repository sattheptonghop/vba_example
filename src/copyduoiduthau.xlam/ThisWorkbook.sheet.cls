Option Base 1

Private Sub Addins()
Dim cmd As CommandBar
Set cmd = Application.CommandBars("Tools")
Dim com As CommandBarControl
Set com = cmd.Controls.Add(msoControlButton)
com.Caption = "Copy duoi du thau"
com.OnAction = "copyshow"
End Sub

Private Sub Resetbara()
Dim cmd As CommandBar
Set cmd = Application.CommandBars("Tools")
Dim com As CommandBarControl
Set com = cmd.Controls("Copy duoi du thau")
com.Delete
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
If Cancel Then
    Else
    Application.OnKey "+^v", ""
    Resetbara
    'ThisWorkbook.IsAddin = False
    End If
End Sub

Private Sub Workbook_Open()
Application.OnKey "+^v", "copyduoiduthau.xlam!copyshow"
ThisWorkbook.IsAddin = True
Addins
End Sub

