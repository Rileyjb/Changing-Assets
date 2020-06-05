VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeviceChoice 
   Caption         =   "DeviceChoice"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "DeviceChoice.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeviceChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LBDevice_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    For i = 0 To LBDevice.ListCount - 1
        If LBDevice.Selected(i) = True Then
            Cells(4, 6).Value = DeviceChoice.LBDevice.List(i)
            DeviceChoice.Hide
            LBDevice.Clear
            Exit Sub
        End If
    Next i
End Sub

Private Sub LBDevice_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    If KeyAscii = 13 Then
        For i = 0 To LBDevice.ListCount - 1
        If LBDevice.Selected(i) = True Then
            Cells(4, 6).Value = DeviceChoice.LBDevice.List(i)
            DeviceChoice.Hide
            LBDevice.Clear
            Exit Sub
        End If
    Next i
        
    End If
End Sub

Private Sub UserForm_Initialize()
    With DeviceChoice
        .Top = Application.Top + 125 '< change 125 to what u want
        .Left = Application.Left + 25 '< change 25 to what u want
    End With
End Sub
