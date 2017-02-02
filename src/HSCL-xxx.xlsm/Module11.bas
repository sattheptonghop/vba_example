Attribute VB_Name = "Module11"
Public acSheet As Worksheet ' for HyperlinksSort

Sub Run(SelectCells As Range, cmd)
On Error GoTo EndRun:
Fast

'Tao sheet
Dim sData, sSheet1
Set sData = ActiveSheet 'ActiveWorkbook.Sheets("Danhmuc")
        Dim inoidung As Range
        For Each inoidung In SelectCells '.SpecialCells(xlCellTypeVisible)
            'Set sSheet1 = inoidung
            If inoidung <> "" And _
                inoidung.EntireRow.Hidden = 0 And _
                inoidung.EntireColumn.Hidden = 0 Then
                
                Select Case cmd
                    Case "taovanbanTT"
                        Call TaoVanBan(inoidung, "soi0")
                    Case "taovanbansoi"
                        Call TaoVanBan(inoidung, "soi")
                    Case "taovanbanin"
                        Call TaoVanBan(inoidung, "in")
                    Case "taovanban"
                        Call TaoVanBan(inoidung, 1)
                    Case "soi"
                        Call HyperlinksPrint(inoidung, 1)
                    Case "in"
                        Call HyperlinksPrint(inoidung, 0)
                    Case "xoa"
                        Call HyperlinksXoa(inoidung)
                    Case "F5"
                        Call HyperlinksF5(inoidung)
                    Case "Sort"
                        Call HyperlinksSort(inoidung)
                End Select
                
            End If
        Next inoidung
    
'Bat lai man hinh
EndRun:
    sData.Select
    unFast
    On Error GoTo 0
End Sub

Private Sub TaoVanBan(SelectCell As Range, Copy As String)
' lenh kiem tra, in truoc neu da co link
On Error GoTo hala
Application.DisplayAlerts = False
SelectCell.Hyperlinks(1).Follow
If ActiveSheet.Name <> SelectCell.Parent.Name Then
    
 'neu co lenh copy sheet thi thuc hien va tao link
    Select Case Copy
        Case "in"
            ActiveSheet.PrintOut
        Case "soi"
            ActiveSheet.PrintPreview
    End Select
    SelectCell.Parent.Select
    Exit Sub
Else
End If
hala:
Application.DisplayAlerts = True

' lenh thuc hien chinh
On Error GoTo Hal:
    Dim Asheet As Worksheet
    Dim aVi As Boolean  'Bien tinh trang sheet an hay hien
    
    'Neu ton tai sheet co ten nhu vay thi thuc hien
    If WorksheetExists(SelectCell.Value) Then
        Set Asheet = Sheets(SelectCell.Value)
        
        'truong hop sheet dang bi an
        aVi = True
        If Asheet.Visible = False Then
            aVi = False
            Asheet.Visible = True
        End If
        
        'Kich hoat sheet mau
        Asheet.Select
        
        'neu co comment run thi gan
        Dim iComment As Range
        Set iComment = Range_Comment("a_run")
        If Not iComment Is Nothing Then
            iComment.Formula = "=Row(" & SelectCell.AddressLocal(, , , 1) & ")"
        End If
        Call Click
        
        'neu co lenh copy sheet thi thuc hien va tao link
        Select Case Copy
            Case 1
                ActiveSheet.Copy after:=Sheets(Sheets.Count)
                Call CopyRand
                SelectCell.Hyperlinks.Add SelectCell, "", "'" & ActiveSheet.Name & "'" & "!a1"
            Case 0
            Case "in"
                If Not Range_Rand() Is Nothing Then
                    ActiveSheet.Copy after:=Sheets(Sheets.Count)
                    Call CopyRand
                    SelectCell.Hyperlinks.Add SelectCell, "", "'" & ActiveSheet.Name & "'" & "!a1"
                End If
                ActiveSheet.PrintOut
                'HyperlinksPrint SelectCell, 0
            Case "soi0"
                If Not Range_Rand() Is Nothing Then
                    ActiveSheet.Copy after:=Sheets(Sheets.Count)
                    Call CopyRand
                    SelectCell.Hyperlinks.Add SelectCell, "", "'" & ActiveSheet.Name & "'" & "!a1"
                End If
                'ActiveSheet.PrintPreview
                'HyperlinksPrint SelectCell, 1
            Case "soi"
                If Not Range_Rand() Is Nothing Then
                    ActiveSheet.Copy after:=Sheets(Sheets.Count)
                    Call CopyRand
                    SelectCell.Hyperlinks.Add SelectCell, "", "'" & ActiveSheet.Name & "'" & "!a1"
                End If
                ActiveSheet.PrintPreview
                'HyperlinksPrint SelectCell, 1
        End Select
        'phuc hoi lai neu sheet da an tu truoc
        If aVi = False Then Asheet.Visible = aVi
    End If
Hal:
End Sub
Private Sub HyperlinksSort(SelectCell As Range)
'On Error Resume Next
On Error GoTo Hal
Application.DisplayAlerts = False
SelectCell.Hyperlinks(1).Follow
If ActiveSheet.Name <> SelectCell.Parent.Name Then
    If acSheet Is Nothing Then
        Set acSheet = ActiveSheet
    Else
        ActiveSheet.Move after:=acSheet
        Set acSheet = ActiveSheet
    End If
    SelectCell.Parent.Select
    
End If
Hal:
Application.DisplayAlerts = True
On Error GoTo 0
End Sub

Private Sub HyperlinksPrint(SelectCell As Range, Preview As Boolean)
'On Error Resume Next
On Error GoTo Hal
Application.DisplayAlerts = False
SelectCell.Hyperlinks(1).Follow
If ActiveSheet.Name <> SelectCell.Parent.Name Then
    If Preview = True Then ActiveSheet.PrintPreview
    If Preview = False Then ActiveSheet.PrintOut
Else
End If
Hal:
Application.DisplayAlerts = True
On Error GoTo 0
End Sub
Private Sub HyperlinksF5(SelectCell As Range)
'On Error Resume Next
On Error GoTo Hal
Application.DisplayAlerts = False
SelectCell.Hyperlinks(1).Follow
If ActiveSheet.Name <> SelectCell.Parent.Name Then
    ActiveShCal
End If
Hal:
Application.DisplayAlerts = True
On Error GoTo 0
End Sub
Private Sub HyperlinksXoa(SelectCell As Range)
'On Error Resume Next
On Error GoTo Hal
Application.DisplayAlerts = False
SelectCell.Hyperlinks(1).Follow
If ActiveSheet.Name <> SelectCell.Parent.Name Then
    ActiveSheet.Delete
    SelectCell.Parent.Select
    SelectCell.Hyperlinks.Delete
End If
Hal:
Application.DisplayAlerts = True
On Error GoTo 0
End Sub
Private Sub ActiveShCal()
ActiveSheet.Calculate
End Sub
Private Sub ActiveShCal1()
Selection.Calculate
End Sub
Private Sub FmainShow()
FMain.Show
End Sub
Private Sub unFast()
    With Application
        '.Calculation = xlManual
        '.MaxChange = 0.001
        '.CalculateBeforeSave = False
        .ScreenUpdating = True
        .StatusBar = "Done"
    End With
End Sub
Private Sub Fast()
'Tat tinh nang tinh lai, man hinh
    'ActiveShCal
    With Application
        .Calculation = xlManual
        .MaxChange = 0.001
        .CalculateBeforeSave = False
        .ScreenUpdating = False
        .StatusBar = "Doing"
    End With
End Sub


'''''''''''''''''
'   Nut cap nhat noi dung bien ban theo ma
'''''''''''''''''
Sub Click()
On Error Resume Next
If Application.ScreenUpdating = True Then
    Dim iFast
    Fast
    iFast = 1
End If

If Range_Comment("a_run") <> 0 Then
ActiveShCal
    Dim target
    Set target = Range_Comment("a_run").Offset(1, 0).End(xlToRight)
    Call RunCopy(target)
    Call RunCopy(target.Offset(0, 1))
    Call RunCopy(target.Offset(0, 2))
    Call RunCopy(target.Offset(0, 3))
ActiveShCal
Else
    Application.StatusBar = "Nho tao comment 'run' trong sheet bien ban"
End If

If iFast = 1 Then unFast
On Error GoTo 0
End Sub

''''''''''''''''
'   Chuc nang gan vi tri cua o co chua comment "a_run" o sheet hien tai
''''''''''''''''

Function Range_Comment(ByVal Comment As String) As Range
Dim temp As Range
Set temp = ActiveSheet.Cells.Find(what:="a_run", LookIn:=xlComments)
If temp Is Nothing Then
    Set Range_Comment = Nothing
    Else
    Set Range_Comment = temp
End If
End Function

''''''''''''''''
'   Chuc nang gan vi tri cua o co chua "rand"
''''''''''''''''

Function Range_Rand() As Range
Dim temp As Range
Set temp = ActiveSheet.Cells.Find(what:="rand", LookIn:=xlFormulas)
If temp Is Nothing Then
    Set Range_Rand = Nothing
    Else
    Set Range_Rand = temp
End If
End Function
Private Sub RunChange(ByVal target As Range)
'chay tiep neu loi
'
'kiem tra ton tai cua sheet
'
Set SheetDanhMuc = ThisWorkbook.Sheets(target.Offset(1, 0).Value)
    For Each irun In target.Parent.Range(target.Offset(2, 0), target.End(xlDown))
        irun.Offset(0, 1) = "=" & SheetDanhMuc.Cells(target.Value, irun.Value).Address(, , , 1)
    Next irun
End Sub

'''''''''''''''''
'   Chay copy dong tu sheet co ten nhu vay
'   tim theo 2 cot Ma so cong viec va cot Ma phieu
'''''''''''''''''

Sub RunCopy(ByVal target As Range)
Dim shAC, shKL, cCV, cDG, sKL, vDG, rSt, rEn
If WorksheetExists(target.Value) = False Then Exit Sub
Set shAC = ActiveSheet
Set shKL = Sheets(target.Value)
cCV = target.Offset(1, 0).Value
cDG = target.Offset(2, 0).Value
sKL = target.Value
vDG = target.Offset(3, 0).Value
Set rSt = target.Offset(4, 0)
Set rEn = target.Offset(5, 0)

Dim rCV_KL, rCV_A, rId_KL, rId_A
Dim sHave As Boolean, sSum As Boolean

'xoa khoi luong cu
If shAC.Range(rSt.Formula).Offset(1, 0).Address <> _
    shAC.Range(rEn.Formula).Address Then
        shAC.Range(rSt.Formula).Offset(1, 0).Resize(Range(rEn.Formula).Row - Range(rSt.Formula).Row - 1, cDG - 1).Delete Shift:=xlUp
End If

If vDG = "0" Then Exit Sub
If target.Value = "" Then Exit Sub

'gan dong bat dau
rId_A = shAC.Range(rSt.Formula).Offset(1, 0).Row

Select Case cCV

    'Truong hop copy theo 1 cot, khong can tinh tong
    Case 0
        For rId_KL = 1 To shKL.UsedRange.Rows.Count + 1
            If shKL.Cells(rId_KL, cDG) = vDG Then
            Dim i
                For i = rId_KL To shKL.UsedRange.Rows.Count + 1
                    'shAC.Rows(rId_A).Insert Shift:=xlDown
                    If shKL.Cells(i, cDG) <> vDG Then
                        'copy do?ng co? ma? giô?ng  sang AS
                        shKL.Cells(rId_KL, 1).Resize(i - rId_KL, cDG - 1).Copy
                        shAC.Cells(rId_A, 1).Insert Shift:=xlDown
                        't?ng sô? do?ng AS
                        'rId_A = rId_A + 1
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
        
    'Truong hop con lai copy theo 2 cot, tinh tong theo cot dau tien
    Case Else
        For rId_KL = 1 To shKL.UsedRange.Rows.Count
        'cha?y theo cô?t "3"
        
            'pha?t hiê?n tên công viê?c theo "2"
            'neu "2" khong trong tuc la bat dau ten cong viec
            If shKL.Cells(rId_KL, cCV) <> "" Then
                'l?u sô? do?ng de sau nay tinh tong
                rCV_KL = rId_KL
                'ga?n tra?ng tha?i m??i l?u de copy ten cong viec
                sHave = True
                
                'nê?u tra?ng tha?i câ?n ti?nh tô?ng
                If sSum = True Then
                    'ti?nh tô?ng cho AS
                    shAC.Cells(rCV_A, cDG - 2) = "=sum(" & Range(Cells(rCV_A + 1, cDG - 2), Cells(rId_A - 1, cDG - 2)).Address & ")"
                    '"=Sum(P" & rCV_A + 1 & ":P" & rId_A - 1 & ")"
                    'bo? tra?ng tha?i câ?n ti?nh tô?ng
                    sSum = False
                End If
            End If
            
            'pha?t hiê?n cô?t "3" co? ma? giô?ng "V:1033"
            If shKL.Cells(rId_KL, cDG) = vDG Then
                'nê?u tra?ng tha?i m??i
                If sHave = True Then
                    'copy do?ng tên công viê?c sang AS
                    shKL.Cells(rCV_KL, 1).Resize(1, cDG - 1).Copy
                    shAC.Cells(rId_A, 1).Insert Shift:=xlDown
                    'l?u sô? do?ng
                    rCV_A = rId_A
                    't?ng sô? do?ng AS
                    rId_A = rId_A + 1
                    'bo? tra?ng tha?i m??i l?u
                    sHave = False
                End If
                'copy do?ng co? ma? giô?ng  sang AS
                shKL.Cells(rId_KL, 1).Resize(1, cDG - 1).Copy
                shAC.Cells(rId_A, 1).Insert Shift:=xlDown
                't?ng sô? do?ng AS
                rId_A = rId_A + 1
                'ga?n tra?ng tha?i câ?n ti?nh tô?ng
                sSum = True
            End If
        Next
        If sSum = True Then
            'ti?nh tô?ng cho AS
            shAC.Cells(rCV_A, cDG - 2) = "=sum(" & Range(Cells(rCV_A + 1, cDG - 2), Cells(rId_A - 1, cDG - 2)).Address & ")"
            'bo? tra?ng tha?i câ?n ti?nh tô?ng
            sSum = False
        End If
End Select

End Sub

''''''''''
'         Xoa tat ca sheet da copy
'         tim nhung sheet co chua ten la " ("
''''''''''
Sub XoaSheetDaCopy()
Application.DisplayAlerts = False

    For Each ish In ThisWorkbook.Sheets
        If Replace(ish.Name, " (", "") <> ish.Name Then ish.Delete
    Next ish
    
Application.DisplayAlerts = True
End Sub


''''''''''''''''''
'   Chuc nang kiem tra ton tai sheet
''''''''''''''''''
Public Function WorksheetExists(ByVal WorksheetName As String) As Boolean
On Error Resume Next

    WorksheetExists = (Sheets(WorksheetName).Name <> "")

On Error GoTo 0
End Function

''''''''''''''''''
'   Chuc nang dua ra ten cua sheet chua o do
''''''''''''''''''
Function SheetName(rCell As Range, Optional UseAsRef As Boolean) As String

    Application.Volatile

        If UseAsRef = True Then

            SheetName = "'" & rCell.Parent.Name & "'!"

        Else

            SheetName = rCell.Parent.Name

        End If

End Function

''''''''''''''''''
'   copy chet so rand
''''''''''''''''''

Sub CopyRand()
'Application.Calculation = xlCalculationManual
On Error GoTo Hal

Dim FoundCell As Range
Dim FirstAddr As String
Set FoundCell = Cells.Find(what:="rand", LookIn:=xlFormulas)

If Not FoundCell Is Nothing Then
    FirstAddr = FoundCell.Address
End If
Do Until FoundCell Is Nothing
    FoundCell = FoundCell.Value
    Set FoundCell = Cells.FindNext(after:=FoundCell)
    If FoundCell.Address = FirstAddr Then
        Exit Do
    End If
Loop

Hal:
'Application.Calculation = xlCalculationAutomatic
End Sub

Function UpperUni(chuoi As String) As String
UpperUni = Application.WorksheetFunction.Trim(UCase(chuoi))
End Function
