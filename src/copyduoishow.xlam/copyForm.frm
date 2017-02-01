VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} copyForm 
   Caption         =   "Copy duôi"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2805
   OleObjectBlob   =   "copyForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "copyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ChayButton_Click()
    Unload Me
    Call autoo(Dongcopyt.Value, cotrunt.Value, cotdowt.Value)
End Sub


Private Sub UserForm_Initialize()
    Dongcopyt.Value = Selection.EntireRow.Address
    For Each i In Selection.EntireRow.Columns
        If Application.WorksheetFunction.CountBlank(i) = 0 Then
            cotdowt.Value = i.Column
            Exit For
        End If
    Next i
End Sub
