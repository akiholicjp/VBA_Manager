Function ChangeRadix(ByVal InData As String, _
                            ByVal InRadix As Long, ByVal OutRadix As Long, Optional ByVal Length As Long = -1) As String
   Const C_DIGIT_STRING = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
   Dim tmpRem As Long
   Dim tmpQuot As Long
   Dim sQuot As String
   Dim inlen As Long
   Dim col As Long
   Dim sOutData As String
   Dim n As Long
   ChangeRadix = ""
   InData = Replace(Replace(InData, " ", ""), vbLf, "")
   If InData = "" Then Exit Function
   If InRadix <= 36 Then InData = StrConv(InData, vbLowerCase)
   sOutData = ""
   Do
      sQuot = ""
      tmpRem = 0
      inlen = Len(InData)
      For n = 1 To inlen
         col = InStr(C_DIGIT_STRING, Mid(InData, n, 1))
         If col = 0 Then Exit Function
         tmpRem = tmpRem * InRadix + col - 1
         tmpQuot = tmpRem \ OutRadix
         tmpRem = tmpRem Mod OutRadix
         If tmpQuot <> 0 Or sQuot <> "" Then
            sQuot = sQuot & Mid(C_DIGIT_STRING, tmpQuot + 1, 1)
         End If
      Next
      sOutData = Mid(C_DIGIT_STRING, tmpRem + 1, 1) & sOutData
      InData = sQuot
   Loop While sQuot <> ""
   If OutRadix <= 36 Then sOutData = StrConv(sOutData, vbUpperCase)
   If Len(sOutData) < Length Then
      ChangeRadix = WorksheetFunction.Rept("0", Length - Len(sOutData)) & sOutData
   Else
      If Length < 0 Then
         Length = WorksheetFunction.RoundUp(WorksheetFunction.Log(InRadix ^ inlen) / WorksheetFunction.Log(OutRadix), 0)
         If Len(sOutData) < Length Then
            ChangeRadix = WorksheetFunction.Rept("0", Length - Len(sOutData)) & sOutData
            Exit Function
         End If
      End If
      ChangeRadix = sOutData
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_ChangeRadix()
   assert "12", ChangeRadix(InData:="1100", InRadix:=2, OutRadix:=10)
   assert "0012", ChangeRadix(InData:="1100", InRadix:=2, OutRadix:=10, Length:=4)
   assert "12", ChangeRadix("1100", 2, 10)
   assert "43981", ChangeRadix("abcd", 16, 10)
   assert "1111111111111111", ChangeRadix("ffff", 16, 2)
   assert "00000000000000001111111111111111", ChangeRadix("ffff", 16, 2, 32)
   assert "0011", ChangeRadix("11", 10, 10, 4)
End Sub
