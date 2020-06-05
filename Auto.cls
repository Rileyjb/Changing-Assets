VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Auto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub start()
    Workbooks.Open ("https://naf365.sharepoint.com/sites/HelpDeskGroup/Shared%20Documents/General/Equipment/current%20inv%20count.xlsx?web=1")
    
    Workbooks.Open ("C:/compare assets macro.xlsm")
    
    ThisWorkbook.Activate
End Sub

Sub fill_temp()
    Dim i As Integer
    Dim y As Integer
    Dim col As String
    Dim endcell As Integer
    Dim tags As Variant
    
    y = 2
    endcell = 2
    
    Workbooks("current inv count.xlsx").Activate
    
    For i = 1 To ActiveWorkbook.Worksheets.Count
        ActiveWorkbook.Worksheets(i).Activate
        For x = 1 To 18
            If Cells(1, x).Value <> vbNullString Then
                Select Case x
                    Case 1
                        col = "A"
                    Case 2
                        col = "B"
                    Case 3
                        col = "C"
                    Case 4
                        col = "D"
                    Case 5
                        col = "E"
                    Case 6
                        col = "F"
                    Case 7
                        col = "G"
                    Case 8
                        col = "H"
                    Case 9
                        col = "I"
                    Case 10
                        col = "J"
                    Case 11
                        col = "K"
                    Case 12
                        col = "L"
                    Case 13
                        col = "M"
                    Case 14
                        col = "N"
                    Case 15
                        col = "O"
                    Case 16
                        col = "P"
                    Case 17
                        col = "Q"
                    Case 18
                        col = "R"
                End Select
                Do While Cells(y, x).Value <> vbNullString
                    y = y + 1
                Loop
                If y <> 2 Then
                    tags = Range(col & "2", col & CStr(y - 1)).Value
                    Workbooks("compare assets macro.xlsm").Activate
                    Range("B" & endcell, "B" & endcell + y - 3).Value = tags
                    endcell = endcell + (y - 2)
                    Workbooks("current inv count.xlsx").Activate
                End If
                y = 2
            End If
        Next x
    Next i
End Sub

Sub close_sheet()
    Workbooks("current inv count.xlsx").Close
End Sub
