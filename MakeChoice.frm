VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MakeChoice 
   Caption         =   "Make Choice"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "MakeChoice.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MakeChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LBMake_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    
    For i = 0 To LBMake.ListCount - 1
        If LBMake.Selected(i) = True Then
            Cells(5, 6).Value = MakeChoice.LBMake.List(i)
            MakeChoice.Hide
            LBMake.Clear
            Exit Sub
        End If
    Next i
End Sub



Private Sub LBMake_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        For i = 0 To LBMake.ListCount - 1
        If LBMake.Selected(i) = True Then
            Cells(5, 6).Value = MakeChoice.LBMake.List(i)
            MakeChoice.Hide
            LBMake.Clear
            Exit Sub
        End If
    Next i
    End If
End Sub

Private Sub UserForm_Initialize()
    With MakeChoice
        .Top = Application.Top + 125 '< change 125 to what u want
        .Left = Application.Left + 25 '< change 25 to what u want
    End With
End Sub
