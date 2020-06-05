VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LocChoice 
   Caption         =   "Location Choice"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "LocChoice.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "LocChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LB2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    For i = 0 To LB2.ListCount - 1
        If LB2.Selected(i) = True Then
            Cells(5, 6).Value = LocChoice.LB2.List(i)
            LocChoice.Hide
            LB2.Clear
            Exit Sub
        End If
    Next i
End Sub

Private Sub LB2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
     If KeyAscii = 13 Then
        For i = 0 To LB2.ListCount - 1
        If LB2.Selected(i) = True Then
            Cells(5, 6).Value = LocChoice.LB2.List(i)
            LocChoice.Hide
            LB2.Clear
            Exit Sub
        End If
    Next i
    End If
End Sub

Private Sub UserForm_Initialize()
    With LocChoice
        .Top = Application.Top + 125 '< change 125 to what u want
        .Left = Application.Left + 25 '< change 25 to what u want
    End With
End Sub
