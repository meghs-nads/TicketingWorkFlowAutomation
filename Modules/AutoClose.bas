Attribute VB_Name = "AutoClose"
Dim DownTime As Date

Sub SetTimer()
    DownTime = Now + TimeValue("00:25:00")
   ' Application.OnTime EarliestTime:=DownTime, _
   '   Procedure:="ShutDown", Schedule:=True
End Sub
Sub StopTimer()
    On Error Resume Next
    'Application.OnTime EarliestTime:=DownTime,      Procedure:="ShutDown", Schedule:=False
 End Sub
Sub ShutDown()
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    ThisWorkbook.Close
End Sub
