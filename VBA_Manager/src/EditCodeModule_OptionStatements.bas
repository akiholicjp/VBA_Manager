Private Function EditCodeModule_OptionStatements( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean

   If propState("CODE_STATE") <> "<DECLARE>" Then EditCodeModule_OptionStatements = True: Exit Function
   Dim regResults As Object
   Dim s As String: s = Trim(sLine)
   Dim oOptions As Object: Set oOptions = propMod("OPTIONS")
   If RegExe(s, "^Option\s+Base\s+(0|1)", regResults) Then
      If Not oOptions.Exists("Base") Then
         oOptions("Base") = regResults(1)
      ElseIf oOptions("Base") = "" Then
         oOptions("Base") = regResults(1)
      ElseIf regResults(1) = oOptions("Base") Then
         bRemove = True
      Else
         msgError = msgError & "Option BaseéwíËÇ™ñµèÇÇµÇƒÇ¢Ç‹Ç∑: " & propBean("FILE") & vbCrLf
         EditCodeModule_OptionStatements = False
         Exit Function
      End If
   ElseIf RegExe(s, "^Option\s+Compare\s+(Binary|Text|Database)", regResults) Then
      If Not oOptions.Exists("Compare") Then
         oOptions("Compare") = regResults(1)
      ElseIf oOptions("Compare") = "" Then
         oOptions("Compare") = regResults(1)
      ElseIf regResults(1) = oOptions("Compare") Then
         bRemove = True
      Else
         msgError = msgError & "Option CompareéwíËÇ™ñµèÇÇµÇƒÇ¢Ç‹Ç∑" & propBean("FILE") & vbCrLf
         EditCodeModule_OptionStatements = False
         Exit Function
      End If
   ElseIf RegTest(s, "^Option\s+Explicit") Then
      If Not oOptions.Exists("Explicit") Then
         oOptions("Explicit") = True
      ElseIf Not oOptions("Explicit") Then
         oOptions("Explicit") = True
      Else
         bRemove = True
      End If
   ElseIf RegTest(s, "^Option\s+Private") Then
      If Not oOptions.Exists("Private") Then
         oOptions("Private") = True
      ElseIf Not oOptions("Private") Then
         oOptions("Private") = True
      Else
         bRemove = True
      End If
   End If
   EditCodeModule_OptionStatements = True
End Function
