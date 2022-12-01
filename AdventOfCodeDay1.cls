Sub elfCaloryCounter()
Set topList = CreateObject("System.Collections.ArrayList")
For i = 1 To Range("A" & Rows.count).End(xlUp).Row
    count = count + Worksheets("Tabelle1").Cells(i, 1).Value
    If IsEmpty(Worksheets("Tabelle1").Cells(i, 1)) Then
        topList.Add count
        count = 0
    End If
Next i
topList.Sort
topList.Reverse
MsgBox "Amount max " & topList(0) & " Sum Top3 " & (topList(0) + topList(1) + topList(2))
End Sub
