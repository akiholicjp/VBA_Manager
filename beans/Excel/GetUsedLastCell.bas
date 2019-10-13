Function GetUsedLastCell(Optional ByVal oSheet As Worksheet) As Range
   If oSheet Is Nothing Then Set oSheet = ActiveSheet
   Dim oRange As Range
   On Error Resume Next
   Set oRange = Union(oSheet.UsedRange.SpecialCells(xlCellTypeConstants), _
                  oSheet.UsedRange.SpecialCells(xlCellTypeFormulas))
   If Err.Number = 1004 Then
      Err.Clear
      Set oRange = oSheet.UsedRange.SpecialCells(xlCellTypeConstants)
   End If
   If Err.Number = 1004 Then
      Err.Clear
      Set oRange = oSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
   End If
   If Err.Number <> 0 Then
      Err.Clear
      Set GetUsedLastCell = Nothing
      Exit Function
   End If
   Set GetUsedLastCell = oSheet.Cells(oRange.Rows.Count, oRange.Columns.Count)
End Function
