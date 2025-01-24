Attribute VB_Name = "Module7"
Sub Sorting_Problems()
Attribute Sorting_Problems.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Sorting_Problems Macro
'

'
    ActiveWorkbook.Worksheets("Problems").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Problems").Sort.SortFields.Add2 Key:=Range( _
        "B4:B9488"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Problems").Sort.SortFields.Add2 Key:=Range( _
        "E4:E9488"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Problems").Sort.SortFields.Add2 Key:=Range( _
        "F4:F9488"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Problems").Sort
        .SetRange Range("B12:BB9488")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
