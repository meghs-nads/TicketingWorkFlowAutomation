Attribute VB_Name = "Module4"
Sub Debert_Var()
Attribute Debert_Var.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Debert_Var Macro
'

'
Call NewConfig
    ActiveSheet.Range("$A$6:$N$9488").AutoFilter Field:=12, Criteria1:="1"
    Application.GoTo Reference:=Range("A6"), Scroll:=True

End Sub


Sub STJ_Var()
'
' STJ_Var Macro
'

'
Call NewConfig
    ActiveSheet.Range("$A$6:$N$9488").AutoFilter Field:=13, Criteria1:="1"
    Application.GoTo Reference:=Range("A6"), Scroll:=True

End Sub



Sub WET_Var()
'
' STJ_Var Macro
'

'
Call NewConfig
    ActiveSheet.Range("$A$6:$N$9488").AutoFilter Field:=14, Criteria1:="1"
    Application.GoTo Reference:=Range("A6"), Scroll:=True

End Sub
