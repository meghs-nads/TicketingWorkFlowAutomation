VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Text

Private Sub Worksheet_Change(ByVal Target As Range)
  
   
'Dim owner As Integer, signature As String
'MsgBox "Hi In Change Method of WorkSheet"

    'Dim ws As Worksheet
    'Dim activeRow As Long
    'Dim Watchcell As Range
    
    'Set ws = ThisWorkbook.Worksheets("Problems")
    'Set Watchcell = Me.Columns("G")
    
    'activeRow = ActiveCell.Row
    
    'If frmTktInput.Visible = False Then
        'If the 'status' is changed then only show the form
     '   If Not Intersect(Target, Watchcell) Is Nothing Then
    
            'If the active cell has finished or cancelled ticket then dont open the form
      '      If ws.Cells(activeRow, 7).Value <> "Completed" And ws.Cells(activeRow, 7).Value <> "Cancelled" Then
       '         frmTktInput.Show
        '    End If
       ' End If
    'End If
    
    
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startDateCol As Range
    Dim ageCol As Range
    Dim currentDate As Date
    Dim issueDate As Variant
    Dim weeksOld As Long
    Dim i As Long
    Dim Watchcell As Range
    Dim activeRow As Long
    
    ' Set the worksheet
    Set ws = Me ' Refers to the sheet where this code is placed
    
    ' Get the current date
    currentDate = Date
    
    ' Find the last filled row in the "Issue Start Date" column (Column I)
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ' Define the range for the "Issue Start Date" and "Age in Weeks" columns
    Set startDateCol = ws.Range("I4:I" & lastRow) ' Adjust column "I" as needed
    Set ageCol = ws.Range("E4:E" & lastRow) ' Column for "Age in Weeks" (adjust if needed)
    
    ' Define the range for the Watchcell (Column G)
    Set Watchcell = Me.Columns("G")
    
    ' Get the active row (row number where the change occurred)
    activeRow = Target.Row
    
    ' Handling Issue Start Date and Age in Weeks calculation
    If Not Intersect(Target, startDateCol) Is Nothing Then
        Application.EnableEvents = False ' Disable events to prevent infinite loop
        
        ' Loop through each cell in the "Issue Start Date" column
        For i = 1 To startDateCol.Rows.Count
            issueDate = startDateCol.Cells(i, 1).Value
            
            ' Validate if the cell contains a valid date
            If IsDate(issueDate) Then
                ' Calculate the age in weeks
                weeksOld = Int((currentDate - CDate(issueDate)) / 7)
                
                ' Ensure the value is not negative
                If weeksOld < 0 Then weeksOld = 0
                
                ' Update the "Age in Weeks" column
                ageCol.Cells(i, 1).Value = weeksOld
            Else
                ' Clear the "Age in Weeks" column if the date is invalid
                ageCol.Cells(i, 1).Value = ""
            End If
        Next i
        
        Application.EnableEvents = True ' Re-enable events
    End If
    
    ' Handling the Watchcell (Column G) for showing the form
    If frmTktInput.Visible = False Then
        ' If the 'status' is changed in Column G, then show the form
        If Not Intersect(Target, Watchcell) Is Nothing Then
            ' If the active cell in Column G has "Completed" or "Cancelled", don't show the form
            If ws.Cells(activeRow, 7).Value <> "Completed" And ws.Cells(activeRow, 7).Value <> "Cancelled" Then
                frmTktInput.Show
            End If
        End If
    End If


   
   
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Check if cell F1 is clicked
    If Not Intersect(Target, Me.Range("F1")) Is Nothing Then
        ' Call the newProblem macro
        Call newProblem
    End If
End Sub



