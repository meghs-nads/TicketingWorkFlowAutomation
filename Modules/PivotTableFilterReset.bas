Attribute VB_Name = "Module6"
Sub OverviewFilterReset()
Attribute OverviewFilterReset.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("A9").Select
            
       ActiveSheet.PivotTables("PivotTable2").ClearAllFilters
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Status")
        .PivotItems("(blank)").Visible = False
        .PivotItems("Completed").Visible = False
        .PivotItems("Cancelled").Visible = False
        .PivotItems("Opportunity").Visible = False
    End With
     
     
    Application.GoTo Reference:=Range("A1"), Scroll:=True
         
    
End Sub



