Attribute VB_Name = "Overview"
Sub ptRefresh()
   Dim pt As PivotTable
For Each pt In ActiveSheet.PivotTables
        pt.RefreshTable
    Next pt
End Sub
