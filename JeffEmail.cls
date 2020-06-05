VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JeffEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub start()
    Workbooks.Open ("https://naf365.sharepoint.com/sites/HelpDeskGroup/Shared%20Documents/General/Equipment/current%20inv%20count.xlsx?web=1")
    
    ThisWorkbook.Activate
End Sub

Sub fill_temp()
    Dim i As Integer
    Dim z As Integer
    Dim y As Integer
    Dim col As String
    Dim endcell As Integer
    Dim tags As Variant
    Dim p1 As String
    
    Dim olApp As Outlook.Application
    Set olApp = CreateObject("Outlook.Application")
    Dim olMail As Outlook.MailItem
    Dim oAccount As Object
    Set olMail = olApp.CreateItem(olMailItem)
    
    olMail.display
    
    sig = olMail.HTMLBody
    
    p1 = ""
    
    y = 2
    endcell = 1
    
    Workbooks("current inv count.xlsx").Activate
    
    For i = 1 To ActiveWorkbook.Worksheets.Count
        ActiveWorkbook.Worksheets(i).Activate
        If ActiveWorkbook.Worksheets(i).Name = "G4" Or ActiveWorkbook.Worksheets(i).Name = "G5" Then
            GoTo here
        End If
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
                    If Cells(1, x).Value = "Nursery" Or Cells(1, x).Value = "Stack 2" Or _
                        Cells(1, x).Value = "Stack 3" Or Cells(1, x).Value = "Stack 5" Or _
                        Cells(1, x).Value = "Stack 4" Or Cells(1, x).Value = "Stack 1" Then
                            If LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "desktops" Or _
                                LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "standard monitors" Or _
                                LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "phones" Then
                                
                                ThisWorkbook.Worksheets("Email Template").Cells(i, 10).Value = LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & ":" & y - 2
                                p1 = p1 + LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & " being imaged: " & y - 2 & "<br>"
                            Else
                                ThisWorkbook.Worksheets("Email Template").Cells(i, 10).Value = LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & ":" & y - 2
                                p1 = p1 + LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & "s being imaged: " & y - 2 & "<br>"
                            End If
                    Else
                        tags = Range(col & "2", col & CStr(y - 1)).Value
                        test = Cells(1, x).Value
                        ThisWorkbook.Worksheets("Jeff").Activate
                        Range("A" & endcell, "A" & endcell + y - 3).Value = tags
                        endcell = endcell + (y - 2)
                    End If
                End If
                Workbooks("current inv count.xlsx").Activate
                y = 2
            End If
        Next x
'here:
        ThisWorkbook.Worksheets("Jeff").Activate
        z = 1
        Do While Cells(z, 1).Value <> vbNullString
            z = z + 1
        Loop
        If LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "desktops" Or _
            LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "standard monitors" Or _
            LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "phones" Then
            
            ThisWorkbook.Worksheets("Email Template").Cells(i, 9).Value = LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & ":" & z - 1
            p1 = p1 + LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & " in storage: " & z - 1 & "<br>"
        Else
            ThisWorkbook.Worksheets("Email Template").Cells(i, 9).Value = LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & ":" & z - 1
            p1 = p1 + LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & "s in storage: " & z - 1 & "<br>"
        End If
here:
        'sends the email
        olMail.To = "Jeff.Farinich@nafinc.com"
        olMail.CC = "ITInventoryControl@nafinc.com"
        olMail.Subject = "Equipment Count for " & Date
        olMail.HTMLBody = sig
        
        Range("A:A").Clear
        endcell = 1
        Workbooks("current inv count.xlsx").Activate
    Next i
    If p1 <> "" Then
        'fill in table
        For v = 1 To 8
            If ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value <> vbNullString Then
            
                If InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, "widescreen") > 0 Then
                    ThisWorkbook.Worksheets("Email Template").Cells(7, 2).Value = _
                    Mid(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, ":") + 1, Len(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value))
                
                ElseIf InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, "docking station") > 0 Then
                    ThisWorkbook.Worksheets("Email Template").Cells(4, 2).Value = _
                    Mid(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, ":") + 1, Len(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value))
            
                ElseIf InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, "desktops") > 0 Then
                    ThisWorkbook.Worksheets("Email Template").Cells(5, 3).Value = "NAF Imaged: " & _
                    Mid(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, ":") + 1, Len(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value)) _
                    & vbNewLine & "NAF Image In Progress: " & Mid(ThisWorkbook.Worksheets("Email Template").Cells(v, 10).Value, InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 10).Value, ":") + 1, Len(ThisWorkbook.Worksheets("Email Template").Cells(v, 10).Value))

                    ThisWorkbook.Worksheets("Email Template").Cells(5, 2).Value = CInt(Mid(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, ":") + 1, Len(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value))) _
                    + CInt(Mid(ThisWorkbook.Worksheets("Email Template").Cells(v, 10).Value, InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 10).Value, ":") + 1, Len(ThisWorkbook.Worksheets("Email Template").Cells(v, 10).Value)))

                    
                ElseIf InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, "standard monitors") > 0 Then
                   ThisWorkbook.Worksheets("Email Template").Cells(6, 2).Value = _
                    Mid(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, ":") + 1, Len(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value))

                ElseIf InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, "latitude") > 0 Then
                    ThisWorkbook.Worksheets("Email Template").Cells(2, 3).Value = "NAF Imaged: " & _
                    Mid(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, ":") + 1, Len(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value)) _
                    & vbNewLine & "NAF Image In Progress: " & Mid(ThisWorkbook.Worksheets("Email Template").Cells(v, 10).Value, InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 10).Value, ":") + 1, Len(ThisWorkbook.Worksheets("Email Template").Cells(v, 10).Value))
                    
                    ThisWorkbook.Worksheets("Email Template").Cells(2, 2).Value = CInt(Mid(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value, ":") + 1, Len(ThisWorkbook.Worksheets("Email Template").Cells(v, 9).Value))) _
                    + CInt(Mid(ThisWorkbook.Worksheets("Email Template").Cells(v, 10).Value, InStr(1, ThisWorkbook.Worksheets("Email Template").Cells(v, 10).Value, ":") + 1, Len(ThisWorkbook.Worksheets("Email Template").Cells(v, 10).Value)))

                End If
                
            
            End If
        Next v
        
        
        ThisWorkbook.Worksheets("Email Template").Range("A1:C7").Copy
        
        olMail.display
        'SendKeys ("^v")
        DoEvents
    End If
End Sub

Sub close_sheet()
    Workbooks("current inv count.xlsx").Close
End Sub


