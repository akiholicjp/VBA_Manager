' VBA: Import beans/DicProp.bas
Private Function EditCodeModule_AutoMacro( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean

   Dim s As String: s = Trim(sComment)
   Dim vType As Variant
   Dim sProcName As String
   Dim regResults As Object
   Dim oMacroList As Object

   If s <> "" Then
      If propState("CODE_STATE") Like "<PROC:*:True>" Then
         sProcName = SplitEx(propState("CODE_STATE"), Delim:=":")(2)

         For Each vType In Array("Auto_Open", "Auto_Close", "Auto_Activate", "Auto_Deactivate")
            If RegExe(s, "^'\s*VBA\s*:\s*" & vType & "\s*", regResults) Then
               If Not propGlobal.Exists("AutoMacro_" & vType) Then propGlobal.Add Key:="AutoMacro_" & vType, Item:=DicProp()
               Set oMacroList = propGlobal("AutoMacro_" & vType)

               If oMacroList.Exists(sProcName) Then
                  msgError = msgError & "Auto_OpeníËã`Ç™èdï°ÇµÇƒÇ¢Ç‹Ç∑: " & sProcName & "@" & propBean("FILE") & vbCrLf
                  EditCodeModule_AutoMacro = False
                  Exit Function
               End If

               oMacroList.Add Key:=sProcName, Item:=propMod("NAME") & "." & sProcName
            End If
         Next vType
      End If
   End If

   EditCodeModule_AutoMacro = True
End Function

Private Function EditCodeModule_AutoMacroInsert( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean

   Dim s As String
   Dim sBlank As String
   Dim v As Variant
   Dim regResults As Object

   If RegExe(sComment, "^'[-=\s]*VBA\s*:\s*Run Auto Macro\s*:\s*(\w+)[-=\s]*$", regResults) Then
      s = sLine
      sBlank = vbCrLf & String(Len(sLine), " ")
      If Not propGlobal.Exists("AutoMacro_" & regResults(1)) Then EditCodeModule_AutoMacroInsert = True: Exit Function
      If propGlobal("AutoMacro_" & regResults(1)).Count = 0 Then EditCodeModule_AutoMacroInsert = True: Exit Function

      For Each v In propGlobal("AutoMacro_" & regResults(1)).Items
         s = s & sBlank & "Call " & v & "()"
      Next v
      sLine = s
   End If
   EditCodeModule_AutoMacroInsert = True
End Function
