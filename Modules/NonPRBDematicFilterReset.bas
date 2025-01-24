Attribute VB_Name = "Module8"
Sub Non_PRB_Dematic_Open_Cases_Reset_Filter()
Attribute Non_PRB_Dematic_Open_Cases_Reset_Filter.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Non_PRB_Dematic_Open_Cases_Reset_Filter Macro
'

Application.GoTo Reference:=Range("A21"), Scroll:=True
  On Error Resume Next
    ActiveSheet.ShowAllData
  On Error GoTo 0

'If ActiveSheet.AutoFilterMode = True Then
'MsgBox "Auto Filter is turned on"
'Else
'MsgBox "Auto Filter is turned off"
'End If


'
'If ActiveSheet.AutoFilterMode = True Then
'If Sheets("Non PRB - Dematic Open Cases").AutoFilterMode = True Then
'If ActiveSheet.AutoFilterMode = True Then
'MsgBox ("TEST")
'Worksheets("Non PRB - Dematic Open Cases").Unprotect Password:="hh"
'ActiveSheet.ShowAllData
'Worksheets("Non PRB - Dematic Open Cases").Protect Password:="hh"
'Else
'MsgBox ("SEE")
'End If
'ActiveSheet.ShowAllData

 
End Sub
