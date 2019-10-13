Function VBA_GetCommentOfLine(ByVal sLine As String) As Long
   Dim b As Boolean
   Dim iComment As Long
   Dim iQuoteBeg As Long, iQuoteEnd As Long

   iComment = 0
   Do While True
      iComment = InStr(iComment + 1, sLine, "'")
      If iComment = 0 Then
         b = False
         Exit Do
      End If

      iQuoteEnd = 0
      Do While True
         iQuoteBeg = InStr(iQuoteEnd + 1, sLine, """")
         iQuoteEnd = InStr(iQuoteBeg + 1, sLine, """")

         If iQuoteBeg = 0 Or iQuoteEnd = 0 Then
            b = True
            Exit Do
         End If

         If iQuoteBeg < iComment And iComment < iQuoteEnd Then
            b = False
            Exit Do
         End If
      Loop
      If b Then Exit Do
   Loop
   If b Then
      VBA_GetCommentOfLine = iComment
   Else
      VBA_GetCommentOfLine = 0
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_VBA_GetCommentOfLine()
   Dim s As String
   s = "Public Sub Test(a As String) ' TEST"
   assert "' TEST", Mid(s, VBA_GetCommentOfLine(s))
   s = "Public Sub Test(a As String = ""' TEST"")"
   assert 0, VBA_GetCommentOfLine(s)
   s = "Public Sub Test(a As String = ""' TEST"") ' TEST2 "
   assert "' TEST2 ", Mid(s, VBA_GetCommentOfLine(s))
   s = "Public Sub Test(a As String = ""' TEST"", ""AB'TEST2"") ' TEST3 _"
   assert "' TEST3 _", Mid(s, VBA_GetCommentOfLine(s))
End Sub
