Attribute VB_Name = "Module1"
'go button
Sub Do_Things()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim obj As New AssignAssets
    obj.Check
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = "Completed"
End Sub

Sub Open_return()
    'opens return sheet
    Workbooks.Open ("https://naf365.sharepoint.com/sites/HelpDeskGroup/Shared%20Documents/General/Equipment/2020%20RETURNS.xlsx?web=1")
End Sub

Sub Open_inv()
    'opens inv count sheet
    Workbooks.Open ("https://naf365.sharepoint.com/sites/HelpDeskGroup/Shared%20Documents/General/Equipment/current%20inv%20count.xlsx?web=1")
End Sub

'makes pie chart of equipment
Sub Make_Chart()
    Application.ScreenUpdating = False
    Dim obj2 As New Class1
    obj2.start
    obj2.fill_temp
    obj2.Make_Chart
    obj2.close_sheet
    Application.ScreenUpdating = True
End Sub

'closes pie chart
Sub Close_Chart()
    Application.ScreenUpdating = False
    Dim obj3 As New Class2
    obj3.end_chart
    Application.ScreenUpdating = True
End Sub

'shows help box
Sub Help()
    HelpBox.Show
End Sub

'creates email for Jeff F on inventory
Sub jeff()
    MsgBox "Ready?", , "Jeff Email"
    Application.ScreenUpdating = False
    Dim obj As New JeffEmail
    obj.start
    obj.fill_temp
    obj.close_sheet
    Application.ScreenUpdating = True
    MsgBox "Done"
End Sub

'creates email for Tim on low equipment
Sub tim()
    MsgBox "Ready?", , "Tim Email"
    Application.ScreenUpdating = False
    Dim obj As New TimEmail
    obj.start
    obj.fill_temp
    obj.close_sheet
    Application.ScreenUpdating = True
    MsgBox "Done"
End Sub

'runs the compare assets macro for friday inv check
Sub Auto_Load()
    Dim obj As New Auto
    Application.DisplayAlerts = False
    obj.start
    obj.fill_temp
    obj.close_sheet
    Application.DisplayAlerts = True
End Sub
