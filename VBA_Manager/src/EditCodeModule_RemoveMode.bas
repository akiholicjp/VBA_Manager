Private Function EditCodeModule_RemoveMode( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean

   Dim s As String: s = Trim(sComment)
   Dim sKeyword As String: sKeyword = opt
   Dim sKey As String: sKey = "REMOVE_MODE_" & sKeyword
   IF Not propState.Exists(sKey) Then propState(sKey) = False

   If s <> "" Then
      If RegTest(s, "^'?\s*[=-_\s]*VBA\s*:\s*" & sKeyword & "\s*:\s*Begin[=-_\s]*") Then propState(sKey) = True
      ElseIf RegTest(s, "^'?\s*[=-_\s]*VBA\s*:\s*" & sKeyword & "\s*:\s*End[=-_\s]*") Then propState(sKey) = False
   End If
   If propState(sKey) And (Not propState(sKey) Or Not propGlobal(sKeyword)) Then
      bRemove = True
   End If
   EditCodeModule_RemoveMode = True
End Function
