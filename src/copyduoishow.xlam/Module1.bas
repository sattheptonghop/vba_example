Attribute VB_Name = "Module1"
Sub copyshow()
    copyForm.Show
End Sub
Sub autoo(Dongcopyt, colrunt, coldowt)
Attribute autoo.VB_Description = "Macro recorded 11/30/2013 by DHC"
Attribute autoo.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' autoo Macro
' Macro recorded 11/30/2013 by DHC
'

On Error GoTo Hal
'tat de nhanh hon
Nhanh

'Gan thong so
dongCopy = Dongcopyt '"19:27"  'dong copy theo
soDongcopy = Rows(dongCopy).Rows.Count
iRow = Selection.Row
colRun = CInt(colrunt) '1 ' cot giong so thu tu, moi cong viec chi co 1 noi dung
coldow = CInt(coldowt) '4  'cot noi dung lien tuc khong ngat quang

'lenh chinh
Do While Cells(Selection.Row + 0, colRun).End(xlDown) <> "" 'Cells(Selection.Row + 2, coldow).End(xlDown) = ""
    Cells(Selection.Row, colRun).End(xlDown).Select
    Select Case Cells(Selection.Row, colRun).End(xlDown)
    Case ""
        Cells(Selection.Row + 0, coldow).End(xlDown).Offset(0, colRun - coldow).Select
    Case Else
        Cells(Selection.Row + 0, colRun).End(xlDown).Offset(-1, 0).Select
    End Select
        
    makeeInsert (dongCopy)
Loop

Hal:
'mo lai trang thai tinh
Thuong
End Sub

Sub makee(dongCopy As String)
              i = Selection.Row
            'Rows("16:16").Copy
            Rows(i).Insert Shift:=xlDown
            Rows(dongCopy).Copy
            ActiveSheet.Paste
            'Rows(i + 1).Insert Shift:=xlDown
            'Cells(Selection.Row + 2, 1).End(xlDown).Offset(-1, 0).Select

End Sub
Sub makeeInsert(dongCopy As String)
    i = Selection.Row
    Application.StatusBar = "copy: " & i
    
    Rows(dongCopy).Copy
    Rows(i + 1).Insert Shift:=xlDown

End Sub

Sub Nhanh()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.ActiveSheet.DisplayPageBreaks = False
End Sub

Sub Thuong()
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
MsgBox "xong"
End Sub

Sub autothanhtien()
'
' autoo Macro
' Macro recorded 11/30/2013 by DHC
'

'
Do Until Cells(Selection.Row + 2, Selection.Column).End(xlDown) = ""
Cells(Selection.Row, Selection.Column) = "=" & Cells(Selection.Row + 1, Selection.Column).End(xlDown).Address
Cells(Selection.Row + 1, Selection.Column).End(xlDown).Offset(1, 0).Select
    'makee
Loop
'Cells(Selection.Row + 2, 3).End(xlDown).Select
    'makee
     
End Sub
