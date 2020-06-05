VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimEmail"
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
    
    p1 = ""
    
    olMail.display
    
    sig = olMail.HTMLBody
    
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
                
                ' if its empty
                If Cells(1, x).Value = "Floor Dell" Or Cells(1, x).Value = "Dell 24inch" Or Cells(1, x).Value = "Dell 27inch" Or Cells(1, x).Value = "Dell Dock" Or Cells(1, x).Value = _
                    "MacBooks" Then
                    If y = 2 Then
                        If LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "desktops" Or _
                            LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "standard monitors" Or _
                            LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "phones" Then
                            p1 = p1 + "We have no " & Cells(1, x).Value & " " & LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & " in storage.<br>"
                        Else
                            p1 = p1 + "We have no " & Cells(1, x).Value & " " & LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & "s in storage.<br>"
                        End If
                        olMail.HTMLBody = p1 & vbNewLine & sig
                    End If
                End If
                
                'not empty
                If y <> 2 Then
                    ' ignore widescreens
                    If Cells(1, x).Value = "3rd shelf" And ActiveWorkbook.Worksheets(i).Name = "WIDESCREEN" Then
                        GoTo skip
                    End If
                    tags = Range(col & "2", col & CStr(y - 1)).Value
                    test = Cells(1, x).Value
                    
                    'if 24 or targus
                    If Cells(1, x).Value = "Dell 24inch" Or Cells(1, x).Value = "Targus" Or Cells(1, x).Value = "Dev Dell" Then
                        If y - 2 < 7 Then
                            p1 = p1 + "We only have " & y - 2 & " " & test & " " & LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & "s in storage.<br>"
                        
                            olMail.To = "Tim.Nguyen@nafinc.com"
                            olMail.CC = "ITInventoryControl@nafinc.com"
                            olMail.Subject = "Daily Low Equipment Check"
                            olMail.HTMLBody = p1 & vbNewLine & sig
                            
                            y = 0
                            tags = Null
                        End If
                        
                    'if other important ones
                    ElseIf Cells(1, x).Value = "Floor Dell" Or Cells(1, x).Value = "Dell 27inch" _
                        Or Cells(1, x).Value = "Dell Dock" Then
                        If y - 2 < 15 Then
                            If LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "desktops" Or _
                                LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "standard monitors" Or _
                                LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "phones" Then
                                p1 = p1 + "We only have " & y - 2 & " " & test & " " & LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & " in storage.<br>"
                            Else
                                p1 = p1 + "We only have " & y - 2 & " " & test & " " & LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & "s in storage.<br>"
                            End If
                            
                            olMail.To = "Tim.Nguyen@nafinc.com"
                            olMail.CC = "ITInventoryControl@nafinc.com"
                            olMail.Subject = "Daily Low Equipment Check"
                            olMail.HTMLBody = p1 & vbNewLine & sig
                            
                            y = 0
                            tags = Null
                        End If
                        
                    'for macs
                    ElseIf Cells(1, x).Value = "MacBooks" Then
                        If y - 2 < 2 Then
                            p1 = p1 + "We only have " & y - 2 & " " & test & " " & LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & "s in storage.<br>"
                        End If
                        
                        olMail.To = "Tim.Nguyen@nafinc.com"
                        olMail.CC = "ITInventoryControl@nafinc.com"
                        olMail.Subject = "Daily Low Equipment Check"
                        olMail.HTMLBody = p1 & vbNewLine & sig
                        
                        y = 0
                        tags = Null
                    
                    'for latitudes only takes floor aka new
                    ElseIf Cells(1, x).Value <> "Floor/Dell 3500" Or Cells(1, x).Value <> "Floor/Dell 3510" Then 'And ActiveWorkbook.Worksheets(i).Name = "LATITUDE" Then
                        GoTo skip
                    'everything else
                    Else
                        ThisWorkbook.Worksheets("Tim").Activate
                        Range("A" & endcell, "A" & endcell + y - 3).Value = tags
                        endcell = endcell + (y - 2)
                        Workbooks("current inv count.xlsx").Activate
                   End If
                End If
skip:
                y = 2
            End If
        Next x
        ThisWorkbook.Worksheets("Tim").Activate
        z = 1
        Do While Cells(z, 1).Value <> vbNullString
            z = z + 1
        Loop
        
        If z - 1 < 150 And z - 1 > 0 And LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "latitude" Then
            p1 = p1 + "We only have " & z - 1 & " " & LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & "s in storage.<br>"
                
            olMail.To = "Tim.Nguyen@nafinc.com"
            olMail.CC = "ITInventoryControl@nafinc.com"
            olMail.Subject = "Daily Low Equipment Check"
            olMail.HTMLBody = p1 & vbNewLine & sig
                
        ElseIf z - 1 < 15 And z - 1 > 0 Then
            'body
                If LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "desktops" Or _
                    LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "standard monitors" Or _
                    LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) = "phones" Then
                    p1 = p1 + "We only have " & z - 1 & " " & LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & " in storage.<br>"
                Else
                    p1 = p1 + "We only have " & z - 1 & " " & LCase(Workbooks("current inv count.xlsx").Worksheets(i).Name) & "s in storage.<br>"
                End If

            'sends the email
            olMail.To = "Tim.Nguyen@nafinc.com"
            olMail.CC = "ITInventoryControl@nafinc.com"
            olMail.Subject = "Daily Low Equipment Check"
            olMail.HTMLBody = p1 & vbNewLine & sig

        End If
here:
        ThisWorkbook.Worksheets("Tim").Activate
        Range("A:A").Clear
        endcell = 1
        Workbooks("current inv count.xlsx").Activate
    Next i
    If p1 <> "" Then
        olMail.display
    Else
        olMail.HTMLBody = "We have sufficient stock of all equipment." & vbNewLine & sig
        olMail.display
    End If
End Sub

Sub close_sheet()
    Workbooks("current inv count.xlsx").Close
End Sub

