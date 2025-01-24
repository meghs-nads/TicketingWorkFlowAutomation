VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Worksheet_Activate()
    Call ptRefresh
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim ShOverview As Worksheet: Set ShOverview = ThisWorkbook.Sheets("PRB - Overview")
ShOverview.PivotTables("PivotTable2").PivotCache.Refresh
End Sub
