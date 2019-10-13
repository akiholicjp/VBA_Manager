Function SplitEx(str As String, Optional Delim As String = ",", Optional Quote As String = """") As Collection
   Dim lpos As Long
   Dim i As Long
   Dim ch As String
   Dim QuoteMode As Boolean
   lpos = 0
   QuoteMode = False
   Set SplitEx = New Collection
   For i = 1 To Len(str)
      ch = Mid(str, i, 1)
      If Not QuoteMode Then
         If ch = Delim Then
            SplitEx.Add Mid(str, lpos + 1, i - lpos - 1)
            lpos = i
         ElseIf ch = Quote Then
            QuoteMode = True
         End If
      Else
         If ch = Quote Then
            QuoteMode = False
         End If
      End If
   Next i
   If i <> lpos Then
      SplitEx.Add Mid(str, lpos + 1, i - lpos - 1)
   End If
End Function
