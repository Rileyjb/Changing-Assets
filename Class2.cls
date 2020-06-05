VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub end_chart()
On Error Resume Next
    Worksheets("Charts").Activate

    For Each myChart In Worksheets("Charts").ChartObjects
        myChart.Delete
    Next myChart
    Range("A1:E50").Clear
    Worksheets("Charts").Visible = False
    Worksheets("Sheet3").Visible = True
    Worksheets("Sheet3").Activate
    Range("A1:J500").Clear
    Worksheets("Sheet3").Visible = False
    Application.StatusBar = ""
End Sub
