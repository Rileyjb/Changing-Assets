VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BranchChoice 
   Caption         =   "BranchChoice"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "BranchChoice.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "BranchChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LB_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    For i = 0 To LB.ListCount - 1
        If LB.Selected(i) = True Then
            Cells(4, 6).Value = BranchChoice.LB.List(i)
            BranchChoice.Hide
            LB.Clear
            Exit Sub
        End If
    Next i
End Sub

Private Sub LB_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
     Dim i As Integer
     If KeyAscii = 13 Then
        For i = 0 To LB.ListCount - 1
            If LB.Selected(i) = True Then
                Cells(4, 6).Value = BranchChoice.LB.List(i)
                BranchChoice.Hide
                LB.Clear
                Exit Sub
            End If
        Next i
    End If
End Sub

Private Sub UserForm_Initialize()
    With BranchChoice
        .Top = Application.Top + 125 '< change 125 to what u want
        .Left = Application.Left + 25 '< change 25 to what u want
    End With
End Sub
