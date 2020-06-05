VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AssignAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created by Riley Bunce with the help of good and stack overflow
'sheet2(hidden) needs to be upadated with branches location device and models semi regualarly from BV
Public assetSheet As String
Public AssetMsg As String
Public Missing As String
Public Storage As Boolean
Public totalNum As Long
Public progress As Double
Public barwidth As Long
Public percent As Double

'Check which button is clicked, run correct program
Sub Check()
On Error GoTo badboi
    Dim StorageB As OptionButton
    Dim salvage As OptionButton
    Dim assign As OptionButton
    Dim Laptops As OptionButton
    Dim Assets As OptionButton
    Dim Inv As OptionButton
    Dim Remove As OptionButton
    
    AssetMsg = ""
    
    totalNum = WorksheetFunction.CountA(Range("B1:ZZ1"))
    
    'creates buttons
    With Sheets("Sheet1")
        Set StorageB = .Shapes("StorageButton").OLEFormat.Object
        Set salvage = .Shapes("SalvageButton").OLEFormat.Object
        Set assign = .Shapes("AssignButton").OLEFormat.Object
        Set Laptops = .Shapes("LaptopButton").OLEFormat.Object
        Set Assets = .Shapes("NewEntry").OLEFormat.Object
        Set Inv = .Shapes("CurInv").OLEFormat.Object
        Set Remove = .Shapes("RemoveInv").OLEFormat.Object
    End With
    
    'if storage box is checked put things in current inv
    If (ActiveSheet.Shapes("InStorage").ControlFormat.Value = xlOn) Then
        Storage = True
    End If
    
    'runs option chosen
    If StorageB > 0 Then
        'opens inv & return sheet
        Call Open_Files
        'opens conversion sheet
        Call start
        Call Run_Storage
        Call close_sheet
        Unload LoadingScreen
        'if not empty then send email
        If AssetMsg <> "" Then Call Email_Location
        MsgBox "", , "DONE"
        'clears cells
        Range("B1:Z1").Value = Null
        Range("B1").Select
        
    ElseIf salvage > 0 Then
        'opens inv & return sheet
        Call Open_Files
        'opens conversion sheet
        Call start
        Call Run_Salvage
        Call close_sheet
        Unload LoadingScreen
        MsgBox "Done"
        'clears cells0
        Range("B1:Z1").Value = Null
        Range("B1").Select
        
    ElseIf assign > 0 Then
        'opens inv & return sheet
        Call Open_Files
        'opens conversion sheet
        Call start
        Call Run_Assign
        Call close_sheet
        Unload LoadingScreen
        MsgBox "Done"
        'clears Cells
        Range("B1:Z1").Value = Null
        Range("J4").Select
        
    ElseIf Laptops > 0 Then
        'opens inv & return sheet
        Call Open_Files
        'opens conversion sheet
        Call start
        Call Run_Laptop
        'if box checked then place in sheet
        If Storage = True Then Call Cur_inv
        Call close_sheet
        Unload LoadingScreen
        'if not emplty then send email
        If AssetMsg <> "" Then Call Email_Location
        MsgBox "", , "DONE"
        Range("B1").Select
        
    ElseIf Assets > 0 Then
        'checks if user cleared IE instances before running
        answer = MsgBox("Did you close all instances of Internet Explorer in task manager? If you haven't press No, if you have press Yes", vbQuestion & vbYesNo, "Add New Entry")
        If answer = vbNo Then
            End
        End If
        MsgBox "Remember to click the IE icon on the task bar when it shows up."
        Call Add_New_Entry
        Unload LoadingScreen
        MsgBox AssetMsg, , "DONE"
        MsgBox "Please download a new asset sheet."
        Call Download_Sheet
        'clears cells
        MsgBox "Run Add To Current Inv to find any missed assets", , "Check New Assets"
        Range("B2:ZZ2").Value = Null
        Range("B1").Select
        
    ElseIf Inv > 0 Then
        Missing = "Missing:" + vbNewLine
        'opens inv & return sheet
        Call Open_Files
        Call start
        Call Cur_inv
        Call close_sheet
        Unload LoadingScreen
        'if not empty email location of assets
        If AssetMsg <> "" Then Call Email_Location
        'assets not added to bv
        MsgBox Missing, , "Done - add these assets manually"
        'clears cells
        Range("B1:ZZ2").Value = Null
        Range("B1").Select
        
    ElseIf Remove > 0 Then
        'opens inv & return sheet
        Call Open_Files
        Call start
        Call Removal
        Call close_sheet
        Unload LoadingScreen
        MsgBox AssetMsg, , "DONE"
        'clears cells
        Range("B1:ZZ2").Value = Null
        Range("B1").Select
        
    End If
    Exit Sub
badboi:
    Call Send_Error("check")
End Sub

'takes assets out of current inv sheet
Sub Removal()
On Error GoTo badboi
    Dim ie As Object
    Dim found As Boolean
    Dim newTag As String
    Set ie = CreateObject("InternetExplorer.Application")

    With ie
        
        Dim realNum As String, item As Object, itemNum As Integer, itemName As String, deviceName As String
        
        'load progress bar
        Call load
        
        i = 2
        Do While (Cells(1, i).Value <> vbNullString)
            'stuff for bar
            progress = (i - 1) / totalNum
            barwidth = LoadingScreen.border.Width * progress
            percent = Round(progress * 100, 0)
            
            LoadingScreen.bar.Width = barwidth
            LoadingScreen.text.Caption = percent & "% Complete"
            
            DoEvents
            ' end stuff for bar
            
            'converts tag number
            newTag = Cells(1, i).Value
            realNum = Convert_Tag(newTag)
            
            'if asset is found
            If realNum <> "-1" Then
            
                'goes to asset number
                .navigate "https://thebrokernetwork.com/IT/WF_IT_AssetTracking_Entry.aspx?entryId=" & realNum
                .Visible = False
                
                ieBusy ie
                
                Set item = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlDevice")(0)
                itemNum = item.Value
                
                Set device = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlMakeModel")(0)
                deviceName = device.Value
                
                ThisWorkbook.Activate
            
                'device type selector
                Select Case itemNum
                Case 1
                    itemName = "Phones"
                Case 2
                    itemName = "Desktops"
                Case 3
                    itemName = "Standard Monitors"
                Case 7
                    itemName = "Latitude"
                Case 8
                    itemName = "Docking Station"
                Case Else
                    itemName = ""
                End Select
                
                'takes asset out of inv
                x = Remove_inv(itemName, newTag, deviceName)
            End If

            i = i + 1
        Loop
    End With
    
    ie.Quit
    Exit Sub
badboi:
    ie.Quit
    Call Send_Error("Remove From Inventory")
End Sub

'only adds assets to current inv
Sub Cur_inv()
    Dim ie As Object
    Dim newTag As String
    Set ie = CreateObject("InternetExplorer.Application")

    With ie
        
        Dim realNum As String, item As Object, itemNum As Integer, itemName As String, _
            make As Object, makeString As String
        
        'Loads progress bar
        Call load
        
        i = 2
continue:
On Error GoTo badboi
        Do While (Cells(1, i).Value <> vbNullString)
            'stuff for bar
            progress = (i - 1) / totalNum
            barwidth = LoadingScreen.border.Width * progress
            percent = Round(progress * 100, 0)
            
            LoadingScreen.bar.Width = barwidth
            LoadingScreen.text.Caption = percent & "% Complete"
            
            DoEvents
            ' end stuff for bar
            
            'converts tag number
            newTag = Cells(1, i).Value
            realNum = Convert_Tag(newTag)
            
            'if asset is found
            If realNum <> "-1" Then
                'navigates to web page
                .navigate "https://thebrokernetwork.com/IT/WF_IT_AssetTracking_Entry.aspx?entryId=" & realNum
                .Visible = False
                
                'puts correct values in boxes to mark asset as storage
                ieBusy ie
                Set item = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlDevice")(0)
                itemNum = item.Value
                Set make = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlMakeModel")(0)
                makeString = make.Value
                ThisWorkbook.Activate
            
                'device type selection
                Select Case itemNum
                Case 1
                    itemName = "Phones"
                Case 2
                    itemName = "Desktops"
                Case 3
                    itemName = "Standard Monitors"
                Case 7
                    itemName = "Latitude"
                Case 8
                    itemName = "Docking Station"
                Case Else
                    itemName = ""
                End Select
                
                'puts asset in current inv sheet
                x = Input_inv(itemName, makeString, newTag)
            End If
            
            i = i + 1
        Loop
    End With
    
    ie.Quit
    
    Exit Sub
badboi:
    'keeps track of missing assets
    Missing = Missing + newTag + ", " + vbNewLine
    GoTo continue
End Sub

'run if storage button is chosen
Sub Run_Storage()
On Error GoTo badboi
    Dim ie As Object
    Dim found As Boolean
    Dim newTag As String
    Dim person As String
    Set ie = CreateObject("InternetExplorer.Application")

    With ie
        
        Dim ogNotes As String, realNum As String, notes As Object, oStatus As Object, oSubmit As Object, _
            assign As Object, location As Object, branch As Object, item As Object, itemNum As Integer, itemName As String, _
            make As Object, makeString As String
           
        'loads progress bar
        Call load

        i = 2
        Do While (Cells(1, i).Value <> vbNullString)
        
            'stuff for bar
            progress = (i - 1) / totalNum
            barwidth = LoadingScreen.border.Width * progress
            percent = Round(progress * 100, 0)
            
            LoadingScreen.bar.Width = barwidth
            LoadingScreen.text.Caption = percent & "% Complete"
            
            DoEvents
            ' end stuff for bar
            
            'converts tag number
            newTag = Cells(1, i).Value
            realNum = Convert_Tag(newTag)
            
            'if asset is found
            If realNum <> "-1" Then
                'navigates to web page
                .navigate "https://thebrokernetwork.com/IT/WF_IT_AssetTracking_Entry.aspx?entryId=" & realNum
                .Visible = False
                
                'puts correct values in boxes to mark asset as storage
                ieBusy ie
                Set oStatus = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlStatus")(0)
                oStatus.Value = 1
                Set assign = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlEmployeeAssigned")(0)
                assign.Value = 0
                Set branch = .document.getElementsByName("ctl00$ContentPlaceHolder1$DropDownList_Domain")(0)
                branch.Value = 68
                Set location = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlLocation")(0)
                location.Value = "SUITE 250"
                Set oSubmit = .document.getElementsByName("ctl00$ContentPlaceHolder1$btnSave")(0)
                Set notes = .document.getElementsByName("ctl00$ContentPlaceHolder1$txtNotes")(0)
                ogNotes = notes.innertext
                notes.Value = "RETURNED " & Date & vbNewLine & ogNotes
                Set item = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlDevice")(0)
                itemNum = item.Value
                Set make = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlMakeModel")(0)
                makeString = make.Value
                oSubmit.Click
                ieBusy ie
                ThisWorkbook.Activate
                
                'device type selection
                Select Case itemNum
                Case 1
                    itemName = "Phones"
                Case 2
                    itemName = "Desktops"
                Case 3
                    itemName = "Standard Monitors"
                Case 7
                    itemName = "Latitude"
                Case 8
                    itemName = "Docking Station"
                Case Else
                    itemName = ""
                End Select
                
                'is storage is checked adds to current inv
                If Storage = True Then
                    x = Input_inv(itemName, makeString, newTag)
                End If
                
            End If
            
            i = i + 1
        Loop
    End With
    
    ie.Quit
    
    Exit Sub
badboi:
    ie.Quit
    Call Send_Error("Stroage")
End Sub

'run if salvage button is chosen
Sub Run_Salvage()
    Dim ie As Object
    Dim found As Boolean
    Dim newTag As String
    Dim person As String
    Set ie = CreateObject("InternetExplorer.Application")
    
    'loads progress bar
    Call load

    With ie
        
        Dim ogNotes As String, tag As String, realNum As String, notes As Object, oStatus As Object, _
            oSubmit As Object, assign As Object, dropOptions As Object, location As Object, branch As Object, _
            item As Object, itemNum As Integer, itemName As String, device As Object, deviceName As String
        
        i = 2
here:
On Error Resume Next
        Do While (Cells(1, i).Value <> vbNullString)
            
            'stuff for bar
            progress = (i - 1) / totalNum
            barwidth = LoadingScreen.border.Width * progress
            percent = Round(progress * 100, 0)
            
            LoadingScreen.bar.Width = barwidth
            LoadingScreen.text.Caption = percent & "% Complete"
            
            DoEvents
            ' end stuff for bar
            
            'converts tag number
            newTag = Cells(1, i).Value
            realNum = Convert_Tag(newTag)
            
            'if asset is found
            If realNum <> "-1" Then
                .navigate "https://thebrokernetwork.com/IT/WF_IT_AssetTracking_Entry.aspx?entryId=" & realNum
                .Visible = False
                
                'puts correct values in boxes to mark asset as salvage
                ieBusy ie
                Set oStatus = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlStatus")(0)
                oStatus.Value = 3
                Set assign = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlEmployeeAssigned")(0)
                assign.Value = 0
                Set branch = .document.getElementsByName("ctl00$ContentPlaceHolder1$DropDownList_Domain")(0)
                branch.Value = 68
                Set location = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlLocation")(0)
                location.Value = "E-WASTE"
                Set oSubmit = .document.getElementsByName("ctl00$ContentPlaceHolder1$btnSave")(0)
                Set notes = .document.getElementsByName("ctl00$ContentPlaceHolder1$txtNotes")(0)
                ogNotes = notes.innertext
                notes.Value = "SALVAGED " & Date & vbNewLine & ogNotes
                
                Set item = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlDevice")(0)
                itemNum = item.Value
                Set device = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlMakeModel")(0)
                deviceName = device.Value
                
                oSubmit.Click
                ieBusy ie
                ThisWorkbook.Activate
            End If
            
            'device type selection
            Select Case itemNum
            Case 1
                itemName = "Phones"
            Case 2
                itemName = "Desktops"
            Case 3
                itemName = "Standard Monitors"
            Case 7
                itemName = "Latitude"
            Case 8
                itemName = "Docking Station"
            Case Else
                itemName = ""
            End Select
            
            'remove asset by going to sheet
            If deviceName = "HP E202" Or deviceName = "HP E201" Or deviceName = "DELL U2419HC" Or deviceName = "DELL U2412MB" _
                    Or deviceName = "DELL P2419H" Or deviceName = "DELL U2719D" Then
                Workbooks("current inv count.xlsx").Sheets("WIDESCREEN").Activate
            ElseIf deviceName = "HP 850 G4" Then
                Workbooks("current inv count.xlsx").Sheets("G4").Activate
            ElseIf deviceName = "HP 850 G5" Then
                Workbooks("current inv count.xlsx").Sheets("G5").Activate
            Else
                Workbooks("current inv count.xlsx").Sheets(itemName).Activate
            End If
            
            Cells.Find(What:=newTag, After:=ActiveCell, LookIn:=xlFormulas2, LookAt _
                :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=False).Select
                
            Selection.Delete Shift:=xlUp
        
            ThisWorkbook.Activate
            
            
            i = i + 1
        Loop
    End With
    
    ie.Quit
    
    Exit Sub
badboi:
    i = i + 1
    ThisWorkbook.Activate
    GoTo here
End Sub

'assigns asset to user
Sub Run_Assign()
On Error GoTo badboi
    Dim ie As Object
    Dim found As Boolean
    Dim newTag As String
    Dim allassets As String
    Set ie = CreateObject("InternetExplorer.Application")

    With ie
        
        allassets = "Assets: "
        
        Dim ogNotes As String, tag As String, realNum As String, _
            notes As Object, branch As Object, oStatus As Object, oSubmit As Object, assign As Object, _
            dropOptions As Object, location As Object, ogBranch As String, newLocation As String, _
            status As Integer, person As Integer, ogPerson As String, item As Object, itemNum As Integer, _
            itemName As String, deviceName As String
            
        Worksheets("Sheet2").Visible = True
        Worksheets("Sheet2").Activate
        
        i = 1
        
        'gets selection menues from hidden sheet2
        Do While Cells(i, 1).Value <> vbNullString
            BranchChoice.LB.AddItem (Cells(i, 1).Value)
            i = i + 1
        Loop
        
        i = 1
        
        Do While Cells(i, 1).Value <> vbNullString
            LocChoice.LB2.AddItem (Cells(i, 2).Value)
            i = i + 1
        Loop
        
        Worksheets("Sheet1").Activate
        Worksheets("Sheet2").Visible = False
        
        'shows menues
        BranchChoice.Show
        LocChoice.Show
        
        'user inputs where and who equipment is for
        ogBranch = Cells(4, 6).Value
        Cells(4, 6).Value = ""
        
        newLocation = Cells(5, 6).Value
        Cells(5, 6).Value = ""
        
        newLocation = UCase(newLocation)
        ogPerson = InputBox("Person")
        
        'loads progress bar
        Call load
        
        i = 2
        Do While (Cells(1, i).Value <> vbNullString)
            
            allassets = allassets & Cells(1, i).Value & ", "
            
            'stuff for bar
            progress = (i - 1) / totalNum
            barwidth = LoadingScreen.border.Width * progress
            percent = Round(progress * 100, 0)
            
            LoadingScreen.bar.Width = barwidth
            LoadingScreen.text.Caption = percent & "% Complete"
            
            DoEvents
            ' end stuff for bar
            Application.StatusBar = "Number of assets " & i
            
            'converts tag number
            newTag = Cells(1, i).Value
            realNum = Convert_Tag(newTag)
            
            'if asset is found
            If realNum <> "-1" Then
                'goes to asset number
                .navigate "https://thebrokernetwork.com/IT/WF_IT_AssetTracking_Entry.aspx?entryId=" & realNum
                .Visible = False
                
                ieBusy ie
                
                Set branch = .document.getElementsByName("ctl00$ContentPlaceHolder1$DropDownList_Domain")(0)
                
                'gets branch from drop down
                For Each elem In branch.getElementsByTagName("option")
                    If elem.innertext = ogBranch Then
                        elem.Selected = True
                        found = True
                        Exit For
                    End If
                Next elem
                
                'checks if branch is valid
                If found = False Then
                    MsgBox "Cant find " & ogBranch
                    Call close_sheet
                    End
                End If
                found = False
                
                'sets location
                Set location = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlLocation")(0)
                
                For Each elem In location.getElementsByTagName("option")
                    If elem.innertext = newLocation Then 'If InStr(elem.innertext, ogPerson) > 0 Then
                        elem.Selected = True
                        found = True
                        Exit For
                    End If
                Next elem
                
                If found = False Then
                    MsgBox "cannot find " & newLocation
                    Call close_sheet
                    End
                End If
                found = False
                
                'sets status
                Set oStatus = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlStatus")(0)
                oStatus.Value = 2
                
                Set assign = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlEmployeeAssigned")(0)
                 
                'gets name from drop down
                For Each elem In assign.getElementsByTagName("option")
                    If elem.innertext = ogPerson Then
                        elem.Selected = True
                        found = True
                        Exit For
                    End If
                Next elem
                
                'checks if name is valid
                If found = False Then
                    MsgBox "Cant find " & ogPerson
                    Call close_sheet
                    End
                End If
                found = False
                
                Set notes = .document.getElementsByName("ctl00$ContentPlaceHolder1$txtNotes")(0)
                
                'puts notes into all caps
                ogNotes = UCase(notes.innertext)
                
                'looks for these words
                reissued = InStr(ogNotes, "REISSUED")
                reissue = InStr(ogNotes, "REISSUE")
                returned = InStr(ogNotes, "RETURNED")
                
                'if they show up put reissue if not put issue
                If reissue <> 0 Or reissued <> 0 Or returned <> 0 Then
                    notes.Value = "REISSUED " & Date & vbNewLine & ogNotes
                Else
                    notes.Value = "ISSUED " & Date & vbNewLine & ogNotes
                End If
                
                Set item = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlDevice")(0)
                itemNum = item.Value
                
                Set device = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlMakeModel")(0)
                deviceName = device.Value
                
                'submit
                Set oSubmit = .document.getElementsByName("ctl00$ContentPlaceHolder1$btnSave")(0)
                oSubmit.Click
                
                ieBusy ie
                ThisWorkbook.Activate
            
                'device type selection
                Select Case itemNum
                Case 1
                    itemName = "Phones"
                Case 2
                    itemName = "Desktops"
                Case 3
                    itemName = "Standard Monitors"
                Case 7
                    itemName = "Latitude"
                Case 8
                    itemName = "Docking Station"
                Case Else
                    itemName = ""
                End Select
                
                'removes asset from current inv sheet
                x = Remove_inv(itemName, newTag, deviceName)
                
            End If

            i = i + 1
            progress = progress + 1
        Loop
    End With
    
    ie.Quit
    
    'allows for easy copy paste into ticket
    Cells(4, 10).Value = Mid(allassets, 1, Len(allassets) - 2)
    
    Exit Sub
badboi:
    ie.Quit
    Call Send_Error("Run Assign")
End Sub

'runs if laptop is being returned
Sub Run_Laptop()
On Error GoTo badboi
    Dim ie As Object
    Dim start As Integer
    Dim endnum As Integer
    Dim found As Boolean
    Dim newTag As String
    Dim person As String
    Dim sheetNum As Integer
    Set ie = CreateObject("InternetExplorer.Application")
    
    start = 1
    endnum = 1
    
    'gets month
    monNum = Month(Now)
    mon = MonthName(monNum)
    mon = UCase(mon)
    
    'goes to sheet month
    sheetNum = Workbooks("2020 RETURNS.xlsx").Worksheets.Count
    
    'if sheet doesn't exist create it
    If Workbooks("2020 RETURNS.xlsx").Worksheets(sheetNum).Name = mon Then
        Workbooks("2020 RETURNS.xlsx").Sheets(mon).Activate
    Else
        Workbooks("2020 RETURNS.xlsx").Activate
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = mon
        Cells(1, 1).Value = "Return Date"
        Cells(1, 2).Value = "1st Issued"
        Cells(1, 3).Value = "Condition"
        Cells(1, 4).Value = "Tag Number"
        Cells(1, 5).Value = "Device"
        Cells(1, 6).Value = "Computer Name"
        Cells(1, 7).Value = "Make"
        Cells(1, 8).Value = "Serial / MAC Number"
        Cells(1, 9).Value = "Branch"
        Cells(1, 10).Value = "Location"
        Cells(1, 11).Value = "Assigned To"
        Cells(1, 12).Value = "Status"
        Cells(1, 13).Value = "Money Back"
    End If
    
    'find starting point
    Do While Cells(start, 1).Value <> vbNullString
        start = start + 1
    Loop
    
    ThisWorkbook.Activate
    
    'loads progress bar
    Call load

    With ie
        
        Dim ogNotes As String, realNum As String, notes As Object, oStatus As Object, oSubmit As Object, _
        assign As Object, dropOptions As Object, location As Object, branch As Object, tagString As String, _
        deviceString As String, makeString As String, snString As String, branchString As String, _
        locationString As String, assignString As String, statusString As String, comString As String, _
        tag As Object, device As Object, make As Object, sn As Object, com As Object
        
        i = 2
        Do While (Cells(1, i).Value <> vbNullString)
            
            'stuff for bar
            progress = (i - 1) / totalNum
            barwidth = LoadingScreen.border.Width * progress
            percent = Round(progress * 100, 0)
            
            LoadingScreen.bar.Width = barwidth
            LoadingScreen.text.Caption = percent & "% Complete"
            
            DoEvents
            ' end stuff for bar
            
            'converts tag
            newTag = Cells(1, i).Value
            realNum = Convert_Tag(newTag)
            
            'if asset found
            If realNum <> "-1" Then
                'goes to web page
                .navigate "https://thebrokernetwork.com/IT/WF_IT_AssetTracking_Entry.aspx?entryId=" & realNum
                .Visible = False
                
                'assigns vars to html
                ieBusy ie
                Set oStatus = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlStatus")(0)
                Set assign = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlEmployeeAssigned")(0)
                Set branch = .document.getElementsByName("ctl00$ContentPlaceHolder1$DropDownList_Domain")(0)
                Set location = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlLocation")(0)
                Set oSubmit = .document.getElementsByName("ctl00$ContentPlaceHolder1$btnSave")(0)
                Set notes = .document.getElementsByName("ctl00$ContentPlaceHolder1$txtNotes")(0)
                Set tag = .document.getElementsByName("ctl00$ContentPlaceHolder1$txtAssetTagNumber")(0)
                Set device = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlDevice")(0)
                Set make = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlMakeModel")(0)
                Set sn = .document.getElementsByName("ctl00$ContentPlaceHolder1$txtSerialMACNumber")(0)
                Set com = .document.getElementsByName("ctl00$ContentPlaceHolder1$txtComputerName")(0)
                
                
                'get data for return sheet
                statusString = oStatus.Value
                
                'device status selection
                Select Case statusString
                Case "2"
                    statusString = "In Production"
                Case "3"
                    statusString = "Salvage"
                Case "1"
                    statusString = "Storage"
                Case "4"
                    statusString = "Unaccounted"
                Case Else
                    statusString = ""
                End Select
                
                assignString = assign.Value
                
                'gets name from drop down
                For Each elem In assign.getElementsByTagName("option")
                    If StrComp(elem.Value, assignString) = 0 Then
                        If elem.Value = 0 Then
                            assignString = ""
                        Else
                            assignString = elem.innertext
                        End If
                        Exit For
                    End If
                Next elem
                
                branchString = branch.Value
                
                'gets branch from drop down
                For Each bra In branch.getElementsByTagName("option")
                    If StrComp(bra.Value, branchString) = 0 Then
                        branchString = bra.innertext
                        Exit For
                    End If
                Next bra
                
                locationString = location.Value
                tagString = tag.Value
                deviceString = device.Value
                
                'device type selection
                Select Case deviceString
                Case 1
                    deviceString = "Phones"
                Case 2
                    deviceString = "Desktops"
                Case 3
                    deviceString = "Standard Monitors"
                Case 7
                    deviceString = "Latitude"
                Case 8
                    deviceString = "Docking Station"
                Case Else
                    deviceString = ""
                End Select
                    
                makeString = make.Value
                snString = sn.Value
                comString = com.Value
                ogNotes = notes.innertext
                
                'adds infor to returns sheet
                Call Open_return(tagString, branchString, deviceString, comString, makeString, snString, _
                locationString, statusString, assignString, ogNotes)

                'assigns values to vars
                oStatus.Value = 1
                assign.Value = 0
                branch.Value = 68
                location.Value = "SUITE 250"
                notes.Value = "RETURNED " & Date & vbNewLine & ogNotes
                
                oSubmit.Click
                ieBusy ie
    
            End If
            
            i = i + 1
        Loop
    End With
    
    ie.Quit
    
    Workbooks("2020 RETURNS.xlsx").Sheets(mon).Activate
    
    'finds new end point
    Do While Cells(endnum, 1).Value <> vbNullString
        endnum = endnum + 1
    Loop
    
    'sends email with returned info
    Call Email_Laptops(start, endnum - 1)
    
    Exit Sub
badboi:
    ie.Quit
    Call Send_Error("Run Laptop")
End Sub

'sends error message with function and user name with error
Sub Send_Error(issue As String)
    Dim olApp As Outlook.Application
    Set olApp = CreateObject("Outlook.Application")
    Dim olMail As Outlook.MailItem
    Dim oAccount As Object
    Set olMail = olApp.CreateItem(olMailItem)
    
    Dim p1 As String
    Dim xWB As Workbook
    
    olMail.display
    
    sig = olMail.HTMLBody

    'body
     p1 = "Changing Assets had an issue at " & issue & " with " & Application.UserName
                        
    'sends the email
    olMail.To = "Riley.Bunce@nafinc.com"
    olMail.Subject = "Error"
    olMail.HTMLBody = p1 & vbNewLine & sig

    olMail.send
    
    ThisWorkbook.Activate
    
    For Each xWB In Application.Workbooks
        If Not (xWB Is Application.ThisWorkbook) Then
            xWB.Close
        End If
    Next
    
    MsgBox "The program had an error. A report has been sent, you should get a reply shortly."
    
    End
End Sub

'takes sn and asset num from sheet
'inputs new assets in bv
'have to click save
Sub Add_New_Entry()
On Error Resume Next
    Dim ie As Object
    Dim found As Boolean
    Dim wd As Object
    Dim newTag As String
    Set ie = CreateObject("InternetExplorer.Application")
    Dim oHtml As HTMLDocument

    With ie
        
        Dim tag As String, oTag As Object, notes As Object, branch As Object, oStatus As Object, _
        oSubmit As Object, location As Object, device As Object, devi As String, deviNum As Integer, _
        ogMake As String
        
        Worksheets("Sheet2").Visible = True
        Worksheets("Sheet2").Activate
        
        i = 1
        
        'makes prompts for device and model
        Do While Cells(i, 3).Value <> vbNullString
            MakeChoice.LBMake.AddItem (Cells(i, 3).Value)
            i = i + 1
        Loop
        
        i = 1
        
        Do While Cells(i, 4).Value <> vbNullString
            DeviceChoice.LBDevice.AddItem (Cells(i, 4).Value)
            i = i + 1
        Loop
        
        Worksheets("Sheet1").Activate
        Worksheets("Sheet2").Visible = False
        
        'shows prompts
        DeviceChoice.Show
        MakeChoice.Show
        
        devi = UCase(Cells(4, 6).Value)
        Cells(4, 6).Value = ""
        
        ogMake = Cells(5, 6).Value
        Cells(5, 6).Value = ""
        
        'device type selection
        Select Case devi
            Case "COMPUTER"
                deviNum = 2
            Case "DOCKING STATION"
                deviNum = 8
            Case "LAPTOP"
                deviNum = 7
            Case "MONITOR"
                deviNum = 3
            Case "PHONE"
                deviNum = 1
            Case "SERVER"
                deviNum = 4
        End Select
        
        'load progress bar
        Call load
        
        i = 2
        Do While (Cells(1, i).Value <> vbNullString)
        
            'stuff for bar
            progress = (i - 1) / totalNum
            barwidth = LoadingScreen.border.Width * progress
            percent = Round(progress * 100, 0)
            
            LoadingScreen.bar.Width = barwidth
            LoadingScreen.text.Caption = percent & "% Complete"
            
            DoEvents
            ' end stuff for bar
            
            ThisWorkbook.Activate
        
            tag = Cells(1, i).Value
            snNum = Cells(2, i).Value
            
            'goes to asset page
            .navigate "https://thebrokernetwork.com/IT/WF_IT_AssetTracking.aspx#"
            .Visible = True
            
            ieBusy ie
            
            'clicks add button
            .document.parentWindow.execScript "add()"
            
            'waits for web page to load
            Application.Wait (Now + TimeValue("0:00:02"))
           
            'find new window
             For Each wd In CreateObject("Shell.Application").Windows
               If wd.Name = "Internet Explorer" Then
                    Application.Wait (Now + TimeValue("0:00:03"))
                    tit = wd.document.Title
                    If InStr(wd.document.Title, "Asset Entry") <> 0 Then
                        With wd
                            found = True
                            Set oStatus = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlStatus")(0)
                            Set oTag = .document.getElementsByName("ctl00$ContentPlaceHolder1$txtAssetTagNumber")(0)
                            Set branch = .document.getElementsByName("ctl00$ContentPlaceHolder1$DropDownList_Domain")(0)
                            Set location = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlLocation")(0)
                            Set oSubmit = .document.getElementsByName("ctl00$ContentPlaceHolder1$btnSave")(0)
                            Set notes = .document.getElementsByName("ctl00$ContentPlaceHolder1$txtNotes")(0)
                            Set make = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlMakeModel")(0)
                            Set device = .document.getElementsByName("ctl00$ContentPlaceHolder1$ddlDevice")(0)
                            Set sn = .document.getElementsByName("ctl00$ContentPlaceHolder1$txtSerialMACNumber")(0)
                            
                            wd.document.Focus
                            
                            oTag.Focus
                            SendKeys tag
                            
                            branch.Value = 68
                            device.Value = deviNum
                            make.Value = ogMake
                            sn.Value = snNum
                            location.Value = "SUITE 250"
                            oStatus.Value = 1
                            notes.Value = "NEW " & Date
                            
                            wd.document.Focus
                            
                            sn.Focus
                            SendKeys "{ENTER}"
                            Application.Wait (Now + TimeValue("0:00:02"))
                            ieBusy wd
                            
                            ieBusy ie
                            
                            ThisWorkbook.Activate
                        
                        
                        End With
                    Exit For
                   End If
               End If
            Next wd
            If found = False Then i = i - 1
            i = i + 1
            found = False
        Loop
    .Quit
    wd.Quit
    End With
    Exit Sub
badboi:
    ie.Quit
    wd.Quit
    Call Send_Error("Add new entry")
End Sub














'****************************************************HELPER METHODS**********************************************'















'searches through other work book for asset number and returns the inv number
'if no number found returns -1
Function Convert_Tag(num As String) As String
    Dim i As Long
    Dim found As Boolean
    
    Workbooks(assetSheet).Activate
    
    If (Left(num, 1) = "0") Then
        num = Mid(num, 2, Len(num) - 1)
    End If
    
    found = False
    
    i = 2
    Do While Cells(i, 4).Value <> vbNullString
        If Cells(i, 4).Value = num Then
            Convert_Tag = Cells(i, 1).Value
            ThisWorkbook.Activate
            found = True
            Exit Do
        End If
        i = i + 1
    Loop
    If (found = False) Then
        'MsgBox "cant find asset enter on your own " & num
        Missing = Missing + num & ", "
        Convert_Tag = "-1"
        ThisWorkbook.Activate
    End If
    ThisWorkbook.Activate
End Function

'closes sheet when done
Sub close_sheet()
On Error GoTo badboi
    Workbooks(assetSheet).Close
    Workbooks("current inv count.xlsx").Save
    Workbooks("current inv count.xlsx").Close
    Workbooks("2020 RETURNS.xlsx").Save
    Workbooks("2020 RETURNS.xlsx").Close
    ThisWorkbook.Activate
    Exit Sub
badboi:
    Call Send_Error("close sheet")
End Sub

'opens workbook that has asset number conversions
Sub start()
On Error GoTo badboi
    Workbooks.Open "C:\Asset List\" & assetSheet
    ThisWorkbook.Activate
    
    'opens return sheet
    Workbooks.Open ("https://naf365.sharepoint.com/sites/HelpDeskGroup/Shared%20Documents/General/Equipment/2020%20RETURNS.xlsx?web=1")
    
    'opens inv count sheet
    Workbooks.Open ("https://naf365.sharepoint.com/sites/HelpDeskGroup/Shared%20Documents/General/Equipment/current%20inv%20count.xlsx?web=1")
    
    ThisWorkbook.Activate
    Exit Sub
badboi:
    MsgBox "Make sure your asset sheet is in the Asset List folder"
    End
End Sub

'give web page time to load
Sub ieBusy(ie As Object)
    Do While ie.Busy Or ie.readyState < 4
        DoEvents
    Loop
End Sub


'gets the last cell of a column
Function Last_Cell(targetRow As Integer, max As Integer) As Integer
    Dim i As Integer
    
    i = 2
    
    Do While Cells(i, targetRow).Value <> vbNullString
        i = i + 1
    Loop
    
    'no space on shelf send error
    If i > max Or i < 1 Then
        Last_Cell = -1
        Exit Function
    End If
    
    Last_Cell = i
End Function

'opens files from c drive and sets name
Sub Open_Files()
On Error GoTo badboi
    Dim fso As Scripting.FileSystemObject
    Dim folder As Scripting.folder
    Dim file As Scripting.file
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder("C:\Asset List")
    
    For Each file In folder.Files
        Cells(250, 1).Value = file
        assetSheet = Mid(Cells(250, 1), 15, Len(Cells(250, 1)))
        Cells(250, 1).Value = Null
    Next
    Exit Sub
badboi:
    Send_Error ("Open Files")
End Sub

'loads the progress bar
Sub load()
    With LoadingScreen
        .bar.Width = 0
        .text.Caption = "0% Complete"
        .Show vbModeless
                
    End With
End Sub

'emails placement of equipment in current inv sheet
Sub Email_Location()
    Dim olApp As Outlook.Application
    Set olApp = CreateObject("Outlook.Application")
    Dim olMail As Outlook.MailItem
    Dim oAccount As Object
    Set olMail = olApp.CreateItem(olMailItem)
    
    olMail.display
    
    sig = olMail.HTMLBody
                        
    'sends the email
    olMail.To = Replace(Application.UserName, " ", ".")
    olMail.Subject = "Equipment Placement"
    olMail.HTMLBody = AssetMsg & vbNewLine & sig

    olMail.send
End Sub

'Downloads new asset sheet, used adter new assets are added
Sub Download_Sheet()
    'download new asset sheet
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")

On Error GoTo done

    With ie

        Dim downloadButton As Object

        'go to homepage
        .navigate "https://thebrokernetwork.com/IT/WF_IT_AssetTracking.aspx"
        .Visible = True

        'click save as and put into asset list
        ieBusy ie
        Set downloadButton = .document.getElementById("ctl00_ContentPlaceHolder1_lnkDownload")
        downloadButton.Click
    End With
    
    'clear directory of all files
    Kill "C:\Asset List\*.*"
    DoEvents
Exit Sub
    
done:
    Exit Sub
End Sub

'sends email for laptops
Sub Email_Laptops(start As Integer, endnum As Integer)
On Error GoTo here
    Dim olApp As Outlook.Application
    Set olApp = CreateObject("Outlook.Application")
    Dim olMail As Outlook.MailItem
    Dim oAccount As Object
    Set olMail = olApp.CreateItem(olMailItem)
    
    Dim p1 As String
    
    Workbooks("2020 RETURNS.xlsx").Activate
    
    olMail.display
    
    sig = olMail.HTMLBody
    
    p1 = ""
    
    For x = start To endnum
        'formats cells for email
       p1 = p1 & Cells(x, 4).Value & " | " & Cells(x, 5).Value & " | " & Cells(x, 6).Value & " | " & Cells(x, 7).Value & " | " & Cells(x, 8).Value & " | " & Cells(x, 9).Value & " | " & Cells(x, 10).Value & " | " & Cells(x, 11).Value & "<br>"
    Next x
                        
    'sends the email
    olMail.To = "TustinHD@nafinc.com"
    olMail.Subject = "Returned Equipment"
    olMail.HTMLBody = p1 & vbNewLine & sig
    
    ThisWorkbook.Activate

    olMail.send
    Exit Sub
here:
    MsgBox "Could not send email", , "Error"
End Sub

'fills in the return sheet with the fields, you manually input first date given and the amount to be returned
Function Open_return(asset As String, branch As String, device As String, _
    com As String, make As String, sn As String, location As String, status As String, assign As String, _
    ogNotes As String)
   
    Dim start As Integer
    Dim mon As String
    Dim monNum As Integer
    Dim i As Integer
    Dim isNew As Integer
    Dim returnDate As String
    
    i = 1
    
    'gets month
    monNum = Month(Now)
    mon = MonthName(monNum)
    mon = UCase(mon)
    
On Error GoTo badsheet
    
    'goes to month sheet
    Workbooks("2020 RETURNS.xlsx").Sheets(mon).Activate
    
    'finds end point
    Do While Cells(i, 1).Value <> vbNullString
        i = i + 1
    Loop
    
    start = i
    
    If device = "Latitude" Then device = "Laptop"
    
    ogNotes = UCase(ogNotes)
    
    'determines if new date is in notes
    isNew = InStr(ogNotes, "NEW")
    
    Cells(start, 1).Value = Date
    Cells(start, 3).Value = InputBox("G Or B For " & asset)
    Cells(start, 4).Value = asset
    Cells(start, 5).Value = device
    Cells(start, 6).Value = com
    Cells(start, 7).Value = make
    Cells(start, 8).Value = sn
    Cells(start, 9).Value = branch
    Cells(start, 10).Value = location
    Cells(start, 11).Value = assign
    Cells(start, 12).Value = status
    
    'if new date is in notes grade laptop
    If isNew <> 0 Then
        On Error GoTo continue
        
            'get date
            returnDate = Mid(ogNotes, isNew + 4, Len(ogNotes))
            
            'gets rid of rest of notes
            If InStr(returnDate, vbNewLine) <> 0 Or InStr(returnDate, " ") <> 0 Then
                returnDate = Left(returnDate, InStr(returnDate, vbNewLine))
            End If
            
            Cells(start, 2).Value = returnDate
            
            'formats date
            datedif = DateDiff("yyyy", returnDate, Date)
            
            'less than year
            If datedif < 1 Then
                If Cells(start, 3).Value = "G" Then
                    Cells(start, 13).Value = "$400"
                Else
                    Cells(start, 13).Value = "$250"
                End If
                
            '1 year
            ElseIf datedif = 1 Then
                If Cells(start, 3).Value = "G" Then
                    Cells(start, 13).Value = "$250"
                Else
                    Cells(start, 13).Value = "$100"
                End If
                
            ' > 1 year
            Else
                If Cells(start, 3).Value = "G" Then
                    Cells(start, 13).Value = "$100"
                Else
                    Cells(start, 13).Value = "$0"
                End If
            End If
    End If
    
continue:
    
    ThisWorkbook.Activate
Exit Function
badsheet:
    ThisWorkbook.Activate
    Call Send_Error("Open_Return")
    End
End Function

'will remove asset form inventory count for you
'takes type of device and asset number
Function Remove_inv(item As String, asset As String, deviceName As String)
    
    item = UCase(item)
'if no number is found show message
On Error GoTo oopsie
    
    'widescreens
    If deviceName = "HP E202" Or deviceName = "HP E201" Or deviceName = "DELL U2419HC" Or deviceName = "DELL U2412MB" _
            Or deviceName = "DELL P2419H" Or deviceName = "DELL U2719D" Then
        Workbooks("current inv count.xlsx").Sheets("WIDESCREEN").Activate
    'G4's
    ElseIf deviceName = "HP 850 G4" Then
        Workbooks("current inv count.xlsx").Sheets("G4").Activate
    'G5's
    ElseIf deviceName = "HP 850 G5" Then
        Workbooks("current inv count.xlsx").Sheets("G5").Activate
    Else
        Workbooks("current inv count.xlsx").Sheets(item).Activate
    End If
    
    'deletes asset and shifts everything up
    Cells.Find(What:=asset, After:=ActiveCell, LookIn:=xlFormulas2, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Select
        
    Selection.Delete Shift:=xlUp

    ThisWorkbook.Activate
       
    Exit Function
oopsie:
    ThisWorkbook.Activate
    MsgBox "Remove " & asset & " From Inventory Count"
End Function

'will put asset in inventory count for you
'takes type of device and asset number and item type
Function Input_inv(item As String, device As String, asset As String)

'if no space is available
On Error GoTo oopsie
    Dim endcell As Integer
    Dim i As Integer
    Dim maxNumber As Integer
    Dim newRow As Integer
    Dim found As Boolean
    
    item = UCase(item)
    
    endcell = -1
    
    found = False
    
    Workbooks("current inv count.xlsx").Sheets(item).Activate
    
    'for phones
    If item = "PHONES" Then
        maxNumber = 48
        
        'for 1 to last column not empty
        For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
            'looks for certain headers
            If Cells(1, i).Value = "Shelf 3" Then
                'checks if first empty cell is less than max
                If Cells(Rows.Count, i).End(xlUp).Row < maxNumber Then
                    'places asset and adds asset message
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For
                End If
            ElseIf Cells(1, i).Value = "Shelf 4" Then
                If Cells(Rows.Count, i).End(xlUp).Row < maxNumber Then
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For
                End If
            ElseIf Cells(1, i).Value = "Shelf 5" Then
                If Cells(Rows.Count, i).End(xlUp).Row < maxNumber Then
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For
                End If
            End If
        Next i
        
        'if found do nothing, else add error message
        If found = False Then
            GoTo oopsie
        End If
        
    'for desktops
    ElseIf item = "DESKTOPS" Then
        
        'for 1 to last column not empty
        For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
            'looks for certain headers
            If Cells(1, i).Value = "Stack 4" Then
                'places asset and adds asset message
                Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                    " for " & item & "<br>" & vbNewLine
                found = True
                Exit For
            End If
        Next i
        
        'if found do nothing, else add error message
        If found = False Then
            GoTo oopsie
        End If
    
    'for monitors
    ElseIf item = "STANDARD MONITORS" Then
        
        'if widescreen go to else to switch page or just stay on page
        If device <> "HP E202" And device <> "HP E201" And device <> "DELL U2419HC" And device <> "DELL U2412MB" _
            And device <> "DELL P2419H" And device <> "DELL U2719D" Then 'add other models?
            
            'for 1 to last column not empty
            For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
            
                'looks for certain headers
                If Cells(1, i).Value = "First Shelf" And Cells(Rows.Count, i).End(xlUp).Row + 1 < 25 Then
                    
                    'places asset and adds asset message
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For
                    
                ElseIf Cells(1, i).Value = "2nd Shelf" And Cells(Rows.Count, i).End(xlUp).Row + 1 < 25 Then
                    
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For
                    
                ElseIf Cells(1, i).Value = "Floor" Then

                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For

                End If
                
            Next i
            
            'if found do nothing, else add error message
            If found = False Then
                GoTo oopsie
            End If
            
        Else
            item = "Widescreen"
            Sheets(item).Activate
            
            i = 4
            maxNumber = 100
            newRow = 4
            
            If device = "DELL U2719D" Then
                'for 1 to last column not empty
                For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
                
                    'looks for certain headers
                    If Cells(1, i).Value = "Dell 27inch" Then

                        'places asset and adds asset message
                        Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                        AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                            " for " & item & "<br>" & vbNewLine
                        found = True
                        Exit For
                    End If
                Next i
                
                'if found do nothing, else add error message
                If found = False Then
                    GoTo oopsie
                End If
            
            ElseIf device = "DELL U2419HC" Then
                
                'for 1 to last column not empty
                For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
                
                    'looks for certain headers
                    If Cells(1, i).Value = "Dell 24inch" Then

                        'places asset and adds asset message
                        Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                        AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                            " for " & item & "<br>" & vbNewLine
                        found = True
                        Exit For
                    End If
                Next i
                
                'if found do nothing, else add error message
                If found = False Then
                    GoTo oopsie
                End If
                
            Else
                
                'for 1 to last column not empty
                For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
                
                    'looks for certain headers
                    If Cells(1, i).Value = "3rd shelf" Then

                        'places asset and adds asset message
                        Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                        AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                            " for " & item & "<br>" & vbNewLine
                        found = True
                        Exit For
                    End If
                Next i
                
                'if found do nothing, else add error message
                If found = False Then
                    GoTo oopsie
                End If
                
            End If
        End If

    'for laptops
    ElseIf item = "LATITUDE" Then
        
        'for g4s
        If device = "HP 850 G4" Then
            
            item = "G4"
            Sheets(item).Activate
            
            'for 1 to last column not empty
            For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
            
                'looks for certain headers
                If Cells(1, i).Value = "Stack 2" Then

                    'places asset and adds asset message
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For
                End If
            Next i
            
            'if found do nothing, else add error message
            If found = False Then
                GoTo oopsie
            End If
        
        'dev laptops
        ElseIf device = "DELL PRECISION M4800" Or device = "DELL PRECISION 5540" Then
            
            'for 1 to last column not empty
            For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
            
                'looks for certain headers
                If Cells(1, i).Value = "Dev Dell" Then

                    'places asset and adds asset message
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For
                End If
            Next i
            
            'if found do nothing, else add error message
            If found = False Then
                GoTo oopsie
            End If
            
        'for g5s
        ElseIf device = "HP 850 G5" Then
            
            item = "G5"
            Sheets(item).Activate
        
            'for 1 to last column not empty
            For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
            
                'looks for certain headers
                If Cells(1, i).Value = "Stack 1" Then

                    'places asset and adds asset message
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For
                End If
            Next i
            
            'if found do nothing, else add error message
            If found = False Then
                GoTo oopsie
            End If
        
        'for macbooks
        ElseIf InStr(1, UCase(device), "APPLE") > 0 Then
            
            'for 1 to last column not empty
            For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
            
                'looks for certain headers
                If Cells(1, i).Value = "MacBooks" Then

                    'places asset and adds asset message
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For
                End If
            Next i
            
            'if found do nothing, else add error message
            If found = False Then
                GoTo oopsie
            End If
                
        'for latitude
        ElseIf device = "DELL LATITUDE 3500" Or device = "DELL LATITUDE 3510" Then
            
           'for 1 to last column not empty
            For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
            
                'looks for certain headers
                If Cells(1, i).Value = "Stack 3" Then

                    'places asset and adds asset message
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For
                End If
            Next i
            
            'if found do nothing, else add error message
            If found = False Then
                GoTo oopsie
            End If

        Else
            AssetMsg = AssetMsg + "I'm not sure where to put " & asset & ", please add it to inventory count on your own." & "<br>" & vbNewLine
            ThisWorkbook.Activate
        End If
    
    'for docking stations
    ElseIf item = "DOCKING STATION" Then
        
        'hp dock
        If device = "HP ULTRASLIM DOCKING STATION" Then
            maxNumber = 50
            
            'for 1 to last column not empty
            For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
                'looks for certain headers
                If Cells(1, i).Value = "HP Shelf" Then
                    'checks if first empty cell is less than max
                    If Cells(Rows.Count, i).End(xlUp).Row < maxNumber Then
                        'places asset and adds asset message
                        Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                        AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                            " for " & item & "<br>" & vbNewLine
                        found = True
                        Exit For
                    End If
                ElseIf Cells(1, i).Value = "Suite 200 (HP)" Then
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For
                End If
            Next i
            
            'if found do nothing, else add error message
            If found = False Then
                GoTo oopsie
            End If
           
        'dell dock
        ElseIf device = "DELL DOCK WD19" Then
            
            'for 1 to last column not empty
            For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
                'looks for certain headers
                If Cells(1, i).Value = "Dell Dock" Then
                    'places asset and adds asset message
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For

                ElseIf Cells(1, i).Value = "Suite 200 (HP)" Then
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For
                End If
            Next i
            
            'if found do nothing, else add error message
            If found = False Then
                GoTo oopsie
            End If
        
        'surface dock
        ElseIf device = "MICROSOFT SURFACE DOCK" Then
            
            maxNumber = 28
            
            'for 1 to last column not empty
            For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
                'looks for certain headers
                If Cells(1, i).Value = "Surface" Then
                    'checks if first empty cell is less than max
                    If Cells(Rows.Count, i).End(xlUp).Row < maxNumber Then
                        'places asset and adds asset message
                        Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                        AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                            " for " & item & "<br>" & vbNewLine
                        found = True
                        Exit For
                    End If
                End If
            Next i
            
            'if found do nothing, else add error message
            If found = False Then
                GoTo oopsie
            End If
        
        'targus
        ElseIf device = "TARGUS DOCK" Then
            
            maxNumber = 28
            
            'for 1 to last column not empty
            For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
                'looks for certain headers
                If Cells(1, i).Value = "Targus" Then
                    'checks if first empty cell is less than max
                    If Cells(Rows.Count, i).End(xlUp).Row < maxNumber Then
                        'places asset and adds asset message
                        Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                        AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                            " for " & item & "<br>" & vbNewLine
                        found = True
                        Exit For
                    End If
                End If
            Next i
            
            'if found do nothing, else add error message
            If found = False Then
                GoTo oopsie
            End If
        
        'zbook
        ElseIf device = "HP ZBOOK DOCK" Then
            
            maxNumber = 28
            
            'for 1 to last column not empty
            For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
                'looks for certain headers
                If Cells(1, i).Value = "Zbook" Then
                    'checks if first empty cell is less than max
                    If Cells(Rows.Count, i).End(xlUp).Row < maxNumber Then
                        'places asset and adds asset message
                        Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                        AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                            " for " & item & "<br>" & vbNewLine
                        found = True
                        Exit For
                    End If
                End If
            Next i
            
            'if found do nothing, else add error message
            If found = False Then
                GoTo oopsie
            End If
        
        'd6000
        ElseIf device = "DELL D6000" Then
            
            'for 1 to last column not empty
            For i = 1 To Cells(1, Columns.Count).End(xlToLeft).Column + 1
                'looks for certain headers
                If Cells(1, i).Value = "Dell D6000" Then
                    'places asset and adds asset message
                    Cells(Cells(Rows.Count, i).End(xlUp).Row + 1, i).Value = asset
                    AssetMsg = AssetMsg + asset & " is on " & Cells(1, i).Value & _
                        " for " & item & "<br>" & vbNewLine
                    found = True
                    Exit For
                End If
            Next i
            
            'if found do nothing, else add error message
            If found = False Then
                GoTo oopsie
            End If
            
        Else
            AssetMsg = AssetMsg + "I'm not sure where to put " & asset & ", please add it to inventory count on your own." & "<br>" & vbNewLine
            ThisWorkbook.Activate
        End If
        
    Else
        AssetMsg = AssetMsg + "I'm not sure where to put " & asset & ", please add it to inventory count on your own." & "<br>" & vbNewLine
        ThisWorkbook.Activate
    End If

    ThisWorkbook.Activate
       
    Exit Function
    
'error
oopsie:
ThisWorkbook.Activate
AssetMsg = AssetMsg + "Cannot find available space to put the " & device & " (" & asset & ") please add it to inventor count on your own." & "<br>" & vbNewLine
End Function

