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

    Dim ws As Worksheet
    Dim activeRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Problems")
    
    If frmTktInput.Visible = False Then
        
        activeRow = ActiveCell.row
        
        'If the active cell has finished or cancelled ticket then dont open the form
        If ws.Cells(activeRow, 7).Value <> "Completed" And ws.Cells(activeRow, 7).Value <> "Cancelled" Then
            frmTktInput.Show
        End If
    End If
    
   
   'Dim rowNum As Long
   'Dim ws As Worksheet

    ' Reference the active worksheet
    'Set ws = ThisWorkbook.Worksheets("Problems")
    
    ' Ensure the form is not already visible
    'If frmTktInput.Visible = False Then
        ' Validate the active cell's row and column
     '   If ActiveCell.Column > 6 And ActiveCell.row > 3 Then
      '      rowNum = ActiveCell.row
            
            ' Check if the row has enough non-empty cells to be considered valid
       '     If Application.WorksheetFunction.CountA(ws.Rows(rowNum)) > 7 Then
        '        With frmTktInput
         '           ' Populate form fields with values from the active row
          '          .txt_IssueDescription.Value = ws.Cells(rowNum, 8).Value
           '         .txtIssueDate.Value = ws.Cells(rowNum, 9).Value
            '        .cmb_Frequency.Value = ws.Cells(rowNum, 10).Value
             '       .cmb_SeverityPriority.Value = ws.Cells(rowNum, 11).Value
              '      .cmb_ComponentsEffected.Value = ws.Cells(rowNum, 12).Value
               '     .cmb_EnvEffected.Value = ws.Cells(rowNum, 15).Value
                '    .cmb_TransactName.Value = ws.Cells(rowNum, 16).Value
                 '   .txtDateEffected.Value = ws.Cells(rowNum, 17).Value
                    
                    
                  '   .txtRecentChanges.Value = ws.Cells(rowNum, 18).Value
                     
                   ' If ws.Cells(rowNum, 19).Value = "No" Then
                    '    .op_WrkArndAvlbl_No.Value = True
                     '   .txtWrkArndAvlbl.Visible = False
                   ' Else
                    '    .ob_WrkArndAvlbl_Yes.Value = True
                     '   .txtWrkArndAvlbl.Visible = True
                      '  .txtWrkArndAvlbl.Value = ws.Cells(rowNum, 19).Value
                   ' End If
                    
                   'If ws.Cells(rowNum, 20).Value = "Yes" Then
                    '    .ob_AbleToReproduce_Yes.Value = True
                   '     .txtAblToReproduce.Visible = False
                    'Else
                     '   .ob_AbleToReproduce_No.Value = True
                      '  .txtAblToReproduce.Visible = True
                       ' .txtAblToReproduce.Value = ws.Cells(rowNum, 20).Value
                    'End If
                    
                    ' Populate steps to reproduce
                   ' .txtStepsToReproduce.Value = ws.Cells(rowNum, 21).Value
                    
                    ' Display the form
                   ' .Show
               ' End With
            
           ' Else 'If the rows are empty
           '      With frmTktInput
           '     .txt_IssueDescription.Value = ""
           '    .txtIssueDate.Value = ""
           '     .cmb_Frequency.Value = ""
           '     .cmb_SeverityPriority.Value = ""
           '     .cmb_ComponentsEffected.Value = ""
           '     .cmb_EnvEffected.Value = ""
           '     .cmb_TransactName.Value = ""
           '     .txtDateEffected.Value = ""
           '     .txtRecentChanges.Value = ""
           '     .txtWrkArndAvlbl.Value = ""
           '     .txtAblToReproduce.Value = ""
           '     .txtStepsToReproduce.Value = ""
           '     .op_WrkArndAvlbl_No.Value = False
           '     .ob_WrkArndAvlbl_Yes.Value = False
           '     .ob_AbleToReproduce_Yes.Value = False
           '     .ob_AbleToReproduce_No.Value = False
           '     .txtWrkArndAvlbl.Visible = False
           '     .txtAblToReproduce.Visible = False
                
                ' Display the empty form
           '     .Show
           ' End With
        
             
           ' End If
        
       ' End If
        
    'End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Check if cell F1 is clicked
    If Not Intersect(Target, Me.Range("F1")) Is Nothing Then
        ' Call the newProblem macro
        Call newProblem
    End If
End Sub


