Private Function EditCodeModule_Import( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean

   Dim s As String: s = Trim(sComment)
   Dim regResults As Object

   If s <> "" Then
      If RegExe(s, "^'?\s*VBA\s*:\s*Import\s+(.*)", regResults) Then
         If Not ParseBeanImport(regResults(1), propBean("OPTIONS"), True, propMod, propGlobal, msgError) Then
            msgError = msgError & "BeanÉÇÉWÉÖÅ[ÉãÇÃí«â¡Ç…é∏îsÇµÇ‹ÇµÇΩ: " & regResults(1) & vbCrLf
            EditCodeModule_Import = False
            Exit Function
         End If
         bRemove = True
      End If
   End If
   EditCodeModule_Import = True
End Function
