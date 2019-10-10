Private Function EditCodeModule_ConvPrivateScope( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean

   Const C_TYPE = "Private"
   If Not propBean("OPTIONS")("Private") Then EditCodeModule_ConvPrivateScope = True: Exit Function
   Dim regResults As Object

   If propState("CODE_STATE") = "<DECLARE>" Then
      If RegExe(sLine, "^(\s*)(?:Public\s)?(\s*Declare\s.*)", regResults) Then
         sLine = regResults(1) & C_TYPE & " " & regResults(2)
      ElseIf RegExe(sLine, "^(\s*)(?:Public\s)?(\s*Enum\s.*)", regResults) Then
         sLine = regResults(1) & C_TYPE & " " & regResults(2)
      ElseIf RegExe(sLine, "^(\s*)(?:Public\s)?(\s*Type\s.*)", regResults) Then
         sLine = regResults(1) & C_TYPE & " " & regResults(2)
      ElseIf RegExe(sLine, "^(\s*)(?:Public\s)?(\s*Const\s.*)", regResults) Then
         sLine = regResults(1) & C_TYPE & " " & regResults(2)
      ElseIf RegExe(sLine, "^(\s*)Public\s(.*)", regResults) Then
         sLine = regResults(1) & C_TYPE & " " & regResults(2)
      End If
   ElseIf propState("CODE_STATE") Like "<PROC:*:True>" Then
      If propState("PrivateScope") <> propState("CODE_STATE") Then
         If RegExe(sLine, "^(\s*)(?:Public\s)?(\s*(?:Sub|Function|Property).*)", regResults) Then
            sLine = regResults(1) & C_TYPE & " " & regResults(2)
         End If
         propState("PrivateScope") = propState("CODE_STATE")
      End If
   ElseIf propState("CODE_STATE") = "<INIT>" Or propState("CODE_STATE") = "<END>" Then
      propState("PrivateScope") = ""
   End If

   EditCodeModule_ConvPrivateScope = True
End Function
