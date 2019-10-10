Private Function EditCodeModule_Global( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean

   Dim s As String: s = Trim(sComment)
   Dim sGVarName As String
   Dim regResults As Object
   Dim oGlobalList As Object

   If s <> "" Then
      If RegExe(s, "^'?\s*(?:VBA\s*:)?\s*Global\s+(Set\s+)?([^\s]+)\s+As\s+([^\s]+)\s*(=.*)", regResults) Then
         If Not propGlobal.Exists("GLOBAL_LIST") Then propGlobal.Add Key:="GLOBAL_LIST", Item:=NewDic()
         Set oGlobalList = propGlobal("GLOBAL_LIST")

         sGVarName = regResults(2)

         If oGlobalList.Exists(sGVarName) Then
            msgError = msgError & "GlobalíËã`Ç™èdï°ÇµÇƒÇ¢Ç‹Ç∑: " & sGVarName & "@" & propBean("FILE") & vbCrLf
            EditCodeModule_Global = False
            Exit Function
         End If

         oGlobalList.Add Key:=sGVarName, Item:=DicProp( _
            "NAME", sGVarName, _
            "FILE", propBean("FILE"), _
            "MODULE", propMod("NAME"), _
            "SET", (regResults(1) Like "Set*"), _
            "TYPE", regResults(3), _
            "INIT", Trim(regResults(4)) _
         )
      End If
   End If
   EditCodeModule_Global = True
End Function

Private Function EditCodeModule_GlobalInsert( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean

   Dim dic As Variant
   Dim sFile As String
   Dim s As String
   Dim sBlank As String

   If RegTest(sComment, "^'[-=\s]+VBA\s*:\s*Global Variable Definition[-=\s]*$") Then
      s = sLine
      sBlank = vbCrLf & String(Len(sLine), " ")
      sFile = ""
      If Not propGlobal.Exists("GLOBAL_LIST") Then EditCodeModule_GlobalInsert = True: Exit Function
      If propGlobal("GLOBAL_LIST").Count = 0 Then EditCodeModule_GlobalInsert = True: Exit Function

      For Each dic In propGlobal("GLOBAL_LIST").Items
         If propGlobal("APPEND_INFO") Then
            If sFile <> dic("FILE") Then
               s = s & sBlank & "'####" & dic("FILE") & "####:###INSERTED###"
               sFile = dic("FILE")
            End If
         End If
         s = s & sBlank & "Public " & dic("NAME") & " As " & dic("TYPE")
      Next dic
      sLine = s
   ElseIf RegTest(sComment, "^'[-=\s]+VBA\s*:\s*Global Variable Initialize[-=\s]*$") Then
      s = sLine
      sBlank = vbCrLf & String(Len(sLine), " ")
      sFile = ""
      If Not propGlobal.Exists("GLOBAL_LIST") Then EditCodeModule_GlobalInsert = True: Exit Function
      If propGlobal("GLOBAL_LIST").Count = 0 Then EditCodeModule_GlobalInsert = True: Exit Function

      For Each dic In propGlobal("GLOBAL_LIST").Items
         If propGlobal("APPEND_INFO") Then
            If sFile <> dic("FILE") Then
               s = s & sBlank & "'####" & dic("FILE") & "####:###INSERTED###"
               sFile = dic("FILE")
            End If
         End If
         If dic("SET") Then
            s = s & sBlank & "Set " & dic("NAME") & " " & dic("INIT")
         Else
            s = s & sBlank & dic("NAME") & " " & dic("INIT")
         End If
      Next dic
      sLine = s
   End If
   EditCodeModule_GlobalInsert = True
End Function
