Function IsDateStr(ByVal a_inStr As String) As Boolean
   Dim strInDate As String
   Dim pos As Long
   If a_inStr = "" Then
      IsDateStr = False
      Exit Function
   End If
   strInDate = a_inStr

   pos = InStr(strInDate, ".")
   If 0 <> pos Then
      strInDate = Mid(strInDate, 1, pos - 1)
   End If

   If Not IsDate(strInDate) Then
      IsDateStr = False
      Exit Function
   End If

   IsDateStr = True
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_IsDateStr()
   assert True, IsDateStr("1/31/2011")
   assert True, IsDateStr("January 31 2011")
   assert True, IsDateStr("Jan 31 2011")
   assert True, IsDateStr("Jan,31,2011")
   assert True, IsDateStr("31/1/2011")
   assert True, IsDateStr("31,January,2011")
   assert True, IsDateStr("31 January 2011")
   assert True, IsDateStr("31 Jan 2011")
   assert True, IsDateStr("31,Jan,2011")
   assert True, IsDateStr("31/January/2011")
   assert True, IsDateStr("31/Jan/2011")
   assert True, IsDateStr("2011�N1��31��")
   assert True, IsDateStr("����23�N1��31��")
   assert True, IsDateStr("2011/1/31")
   assert True, IsDateStr("2011-1-31")
   assert True, IsDateStr("2011,1,31")
   assert True, IsDateStr("2011 1 31")
   assert True, IsDateStr("12:23:34")
   assert True, IsDateStr("AM 12:23:34")
   assert True, IsDateStr("�ߑO 12:23:34")
   assert True, IsDateStr("2011/1/31 12:23:34")
   assert True, IsDateStr("2011     1     31")
   assert True, IsDateStr("2011�@1�@31")
   assert True, IsDateStr("2013/1/2 10:02:05.0")
   assert True, IsDateStr("2013/1/2")
   assert True, IsDateStr("2013�N3��4��")
   assert True, IsDateStr("1:2:3")
   assert True, IsDateStr("2013/1/2 3:4:5")

   assert False, IsDateStr("")
   assert False, IsDateStr("20130102")
   assert False, IsDateStr("2011.1.31")
   assert False, IsDateStr("�ꌎ���")
End Sub
