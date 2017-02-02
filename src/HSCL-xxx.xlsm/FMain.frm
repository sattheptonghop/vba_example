VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FMain 
   Caption         =   "Cay Cau Giay - beta"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3195
   OleObjectBlob   =   "FMain.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub F5active_Click()
    Call Click
End Sub

'
'   Main
'
Private Sub UserForm_Initialize()
    RefEdit1.Value = Selection.Address
End Sub

Private Sub RefEdit1_Change()
    With RefEdit1
       .Value = Replace(RefEdit1, ActiveSheet.Name, "")
       .Value = Replace(RefEdit1, "'", "")
       .Value = Replace(RefEdit1, "!", "")
       .Value = Replace(RefEdit1, "$", "")
    End With
End Sub

''
' Chinh sua va xoa
''
Private Sub Sort_Click()
    FMain.Hide
    Call Run(Range(RefEdit1.Value), "Sort")
    FMain.Show
End Sub

Private Sub F5_Click()
    FMain.Hide
    Call Run(Range(RefEdit1.Value), "F5")
    FMain.Show
End Sub

Private Sub xoa_Click()
    FMain.Hide
    Call Run(Range(RefEdit1.Value), "xoa")
    FMain.Show
End Sub

Private Sub XoaTat_Click()
    FMain.Hide
    'Call XoaTatca
    FMain.Show
End Sub

''
' In an
''
' Hypelink
'
Private Sub Soi_Click()
    FMain.Hide
    Call Run(Range(RefEdit1.Value), "soi")
    FMain.Show
End Sub

Private Sub TaoVB_Click()
    FMain.Hide
    Call Run(Range(RefEdit1.Value), "taovanban")
    FMain.Show
End Sub

Private Sub InVB_Click()
    FMain.Hide
    Call Run(Range(RefEdit1.Value), "in")
    FMain.Show
End Sub

'
' Truc tiep
'

Private Sub SoiTructiep_Click()
    FMain.Hide
    Call Run(Range(RefEdit1.Value), "taovanbansoi")
    FMain.Show
End Sub


Private Sub TaoTructiep_Click()
    FMain.Hide
    Call Run(Range(RefEdit1.Value), "taovanbanTT")
    FMain.Show
End Sub

Private Sub InTructiep_Click()
    FMain.Hide
    Call Run(Range(RefEdit1.Value), "taovanbanin")
    FMain.Show
End Sub














