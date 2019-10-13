' VBA: Import ../Wrapper/_GetPrivateProfileString.bas

Function GetStringFromIniFile(ByVal sFile As String, ByVal sName As String, ByVal sKey As String, Optional ByVal iLen As Long = 255, Optional ByVal sDefault As String = "") As String
   Dim iRet As Long
   Dim sVal As String

   GetStringFromIniFile = sDefault
   sVal = Space$(iLen + 1)
   iRet = GetPrivateProfileString(sName, sKey, sDefault, sVal, iLen, sFile)

   If iRet > 0 Then
      If InStr(sVal, Chr$(0)) > 0 Then
         GetStringFromIniFile = Left$(sVal, InStr(sVal, Chr$(0)) - 1)
      End If
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetStringFromIniFile()
   assert "<NONE>", GetStringFromIniFile(sFile:=ThisWorkbook.Path & "/none.ini", sName:="NON", sKey:="KEY", iLen:=255, sDefault:="<NONE>")
   assert "<NONE>", GetStringFromIniFile(sFile:=ThisWorkbook.Path & "/test.ini", sName:="NON", sKey:="KEY", iLen:=255, sDefault:="<NONE>")
   assert "<NONE>", GetStringFromIniFile(sFile:=ThisWorkbook.Path & "/test.ini", sName:="SECT", sKey:="NON", iLen:=255, sDefault:="<NONE>")
   assert "VAL", GetStringFromIniFile(sFile:=ThisWorkbook.Path & "/test.ini", sName:="SECT", sKey:="KEY", iLen:=255, sDefault:="<NONE>")
End Sub
