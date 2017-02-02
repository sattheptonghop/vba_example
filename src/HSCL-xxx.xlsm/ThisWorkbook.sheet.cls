Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
'xoa cai cu
If Cancel = True Then
    Dim icbc As Object
    For Each icbc In Application.CommandBars("cell").Controls
    If icbc.Tag = "Make" Then icbc.Delete
    Next icbc
    'phuc hoi lai phim f9 voi lenh tinh lai toan bo bang tinh
    Application.OnKey "{F9}", ""
End If
End Sub

Private Sub Workbook_Open()
'tao lenh moi
'If Not Application.Intersect(Target, [Area]) Is Nothing Then
    With Application.CommandBars("cell").Controls _
        .Add(Type:=msoControlButton, Before:=1, temporary:=True)
        .Caption = "Copy"
        .OnAction = "FmainShow"
        .Tag = "Make"
        .FaceId = 123
    End With
    'gan phim f9 voi lenh tinh lai sheet
    Application.OnKey "{F9}", "ActiveShCal1"
'End If
End Sub



