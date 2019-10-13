Sub SaveCSV(ByVal filename As String, tSheet As Worksheet)
   Dim nLastRow As Long, nLastColumn As Long
   Dim i As Long, j As Long, k As Long, cCnt As Long
   Dim strOut As String

   Open filename For Output As #1

   With tSheet
      nLastRow = .UsedRange.Rows(.UsedRange.Rows.Count).row
      nLastColumn = .UsedRange.Columns(.UsedRange.Columns.Count).column
      For i = 1 To nLastRow
         cCnt = 0
         strOut = .Cells(i, 1).Value
         For j = 2 To nLastColumn
            If Trim(.Cells(i, j).Value) <> "" Then
               For k = 1 To cCnt
                  strOut = strOut + ","
               Next k
               cCnt = 0
               strOut = strOut + "," + .Cells(i, j).Value
            Else
               cCnt = cCnt + 1
            End If
         Next j
         Print #1, strOut
      Next i
   End With
   Close #1
End Sub
