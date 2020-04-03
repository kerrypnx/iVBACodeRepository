Sub unMergeCellsAndFliter()
    Sheet1.Range("A:A").UnMerge
    preText = "K"
    Sheet1.Range("A:A").UnMerge
    For Each cell In Sheet1.Range("A1:A" + CStr(Sheet1.UsedRange.Rows.Count))
        If Len(cell.Value) > 2 Then
            preText = cell.Value
        ElseIf preText = "" Then
            Exit For
        Else
            cell.Value = preText
        End If
        Next
End Sub

Sub deleteSpecialRowsByRowName()
    Dim rngResults As Range, rngToDelete As Range
    Dim strFirstAddress As String
    With Sheet1.Range("C:C") 'Adjust to your particular worksheet
        Set rngResults = .Cells.Find(What:="Sub-Total") 'Adjust what you want it to find
        If Not rngResults Is Nothing Then
            Set rngToDelete = rngResults
            strFirstAddress = rngResults.Address
            Set rngResults = .FindNext(After:=rngResults)
            Do Until rngResults.Address = strFirstAddress
                Set rngToDelete = Application.Union(rngToDelete, rngResults)
                Set rngResults = .FindNext(After:=rngResults)
            Loop
        End If
    End With
    If Not rngToDelete Is Nothing Then rngToDelete.EntireRow.Delete
    Set rngResults = Nothing
End Sub