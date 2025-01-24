Attribute VB_Name = "newProblemSelect"
Sub newProblem()

    Dim ws As Worksheet
    Dim statusRange As Long
    Dim activeRow As Variant
    Dim lastRow As Long
    

    
    Set ws = ThisWorkbook.Sheets("Problems")
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).row
    
    activeRow = lastRow + 1
    ws.Cells(activeRow, "G").Select
    
    
End Sub
