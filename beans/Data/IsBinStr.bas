' Function IsBinStr(ByVal a_inStr As String) As Boolean
'    Dim i As Long

'    For i = 0 To Len(a_inStr) - 1
'       If (Mid(a_inStr, i, 1) <> "0") And (Mid(a_inStr, i, 1) <> "1") Then
'          IsBinStr = False
'          Exit Function
'       End If
'    Next i
'    IsBinStr = True
' End Function

Function IsBinStr(ByVal a_inStr As String) As Boolean
   If a_inStr = "" Then
      IsBinStr = False
   Else
      IsBinStr = Replace(Replace(a_inStr, "0", ""), "1", "") = ""
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_IsBinStr()
   assert True, IsBinStr("10101010")
   assert True, IsBinStr("1010101010101010")
   assert True, IsBinStr("10101010101010101010101010101010")
   assert True, IsBinStr("1010101010101010101010101010101010101010101010101010101010101010")
   assert True, IsBinStr("10101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010")
   assert False, IsBinStr("0123456789")
   assert False, IsBinStr("0123456789abcdef")
   assert False, IsBinStr("0123456789ABCDEF")
   assert False, IsBinStr("x")
   assert False, IsBinStr("")
End Sub
