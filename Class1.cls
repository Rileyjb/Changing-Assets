VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
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
    Dim rowlet As String
    Dim endcell As Integer
    Dim tags As Variant
    Dim p1 As String
    
    Application.StatusBar = "Making Chart"
    
    y = 2
    endcell = 2
    
    Workbooks("current inv count.xlsx").Activate
    ThisWorkbook.Worksheets("Sheet3").Visible = True
    
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
                    Select Case i
                    Case 1
                        rowlet = "A"
                    Case 2
                        rowlet = "B"
                    Case 3
                        rowlet = "C"
                    Case 4
                        rowlet = "D"
                    Case 5
                        rowlet = "E"
                    Case 6
                        rowlet = "F"
                    Case 7
                        rowlet = "G"
                    Case 8
                        rowlet = "H"
                    End Select
                    tags = Range(col & "2", col & CStr(y - 1)).Value
                    ThisWorkbook.Worksheets("Sheet3").Activate
                    Cells(1, i).Value = Workbooks("current inv count.xlsx").Worksheets(i).Name
                    Range(rowlet & endcell, rowlet & endcell + y - 3).Value = tags
                    endcell = endcell + (y - 2)
                    Workbooks("current inv count.xlsx").Activate
                End If
                y = 2
            End If
        Next x
        ThisWorkbook.Activate
        endcell = 2
        Workbooks("current inv count.xlsx").Activate
    Next i

End Sub

Sub Make_Chart()
    ThisWorkbook.Worksheets("Charts").Visible = True
    ThisWorkbook.Worksheets("Sheet3").Activate
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet3!R1C1:R300C8", Version:=6).CreatePivotTable TableDestination:= _
        "Charts!R1C1", TableName:="PivotTable4", DefaultVersion:=6 '
    Sheets("Charts").Select
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("PivotTable4")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable4").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable4").RepeatAllLabels xlRepeatLabels
    ActiveWorkbook.ShowPivotTableFieldList = False
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Charts!$A$1:$C$18")
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("WIDESCREEN"), "Sum of WIDESCREEN", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("DOCKING STATION"), "Sum of DOCKING STATION", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("LATITUDE"), "Sum of LATITUDE", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("DESKTOPS"), "Sum of DESKTOPS", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("STANDARD MONITORS"), "Sum of STANDARD MONITORS", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("G5"), "Sum of G5", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("PHONES"), "Sum of PHONES", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("G4"), "Sum of G4", xlSum
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Sum of WIDESCREEN")
        .Caption = "Count of WIDESCREEN"
        .Function = xlCount
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Sum of DOCKING STATION")
        .Caption = "Count of DOCKING STATION"
        .Function = xlCount
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Sum of LATITUDE")
        .Caption = "Count of LATITUDE"
        .Function = xlCount
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Sum of DESKTOPS")
        .Caption = "Count of DESKTOPS"
        .Function = xlCount
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Sum of STANDARD MONITORS")
        .Caption = "Count of STANDARD MONITORS"
        .Function = xlCount
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Sum of G5")
        .Caption = "Count of G5"
        .Function = xlCount
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Sum of PHONES")
        .Caption = "Count of PHONES"
        .Function = xlCount
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Sum of G4")
        .Caption = "Count of G4"
        .Function = xlCount
    End With
    ActiveChart.ChartType = xlPie
    With ActiveChart.PivotLayout.PivotTable.DataPivotField
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveChart.ShowValueFieldButtons = False
    
    ActiveChart.Legend.Select
    Selection.Position = xlTop
    
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Laptops:"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = _
        "=GETPIVOTDATA(""Count of LATITUDE"",R1C1)+GETPIVOTDATA(""Count of G5"",R1C1)+GETPIVOTDATA(""Count of G4"",R1C1)"
    Range("D2").Select
    
    ActiveCell.FormulaR1C1 = "Monitors:"
    Range("E2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=GETPIVOTDATA(""Count of STANDARD MONITORS"",R1C1)+GETPIVOTDATA(""Count of WIDESCREEN"",R1C1)"
    Range("E3").Select
    
    ThisWorkbook.Worksheets("Sheet3").Visible = False
End Sub

Sub close_sheet()
    Workbooks("current inv count.xlsx").Close
    Application.StatusBar = "Complete"
End Sub

