Attribute VB_Name = "Module1"
Function CheckFilters(r As Range) As String

Set AWS = ActiveSheet
fstate = ""

If AWS.FilterMode Then
    c = AWS.AutoFilter.Filters.Count

    'go through each column and check for filters
    For i = 1 To c Step 1
       If AWS.AutoFilter.Filters(i).On Then
            fstate = fstate & r(i + 1).Value & ", "
       End If
    Next i

    'removes the last comma
    fstate = Left(fstate, Len(fstate) - 2)
Else
    fstate = "NO ACTIVE FILTERS"
End If

CheckFilters = fstate

End Function

