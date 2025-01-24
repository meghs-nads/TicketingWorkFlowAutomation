VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
    If ActiveWorkbook.ReadOnly Then
'Call Non_PRB_Dematic_Open_Cases_Reset_Filter
Else
Call SetTimer
'Call Non_PRB_Dematic_Open_Cases_Reset_Filter
MsgBox "Problem Tracking is updated by many users - please make your changes and close it!"
Application.OnTime Now + TimeValue("00:10:00"), "my_email"
Application.OnTime Now + TimeValue("00:20:00"), "my_email2"
End If
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call StopTimer
End Sub
Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
    ' Call StopTimer
    ' Call SetTimer
End Sub
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, _
  ByVal Target As Excel.Range)
  
  'MsgBox "In Selection Change"
  
    'Call StopTimer
    'Call SetTimer
End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Dim result, config_date_check As Double
  'result = WorksheetFunction.Sum(Range("AW:AW"))
  'config_date_check = WorksheetFunction.Sum(Range("BG:BG"))
If result > 0 Then
'Cancel = True
'MsgBox "Can't Save - please fill out all the mandatory fields !"
Else
If config_date_check > 0 Then
'Cancel = True
'MsgBox "Can't Save - All PRB's that have a code / config filled in need a Business Go Live Date !"
Else
'ChecktoSend
End If
End If
End Sub


