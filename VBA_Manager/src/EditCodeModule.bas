Private Function EditCodeModule(ByRef oMod As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean
   Dim i As Long
   Dim sLine As String
   Dim iComment As Long, bContinue As Boolean
   Dim sComment As String
   Dim bRet As Boolean
   Dim bRemove As Boolean
   Dim iRemoveBeg As Long
   Dim iRemoveCnt As Long
   Dim propState As Object
   Dim iDeclLine As Long
   Dim iProcBodyLine As Long
   Dim sProcName As String
   Dim iType As Long

   Set propState = CreateObject("Scripting.Dictionary")

   bRet = True

   With oMod.CodeModule
      iComment = 0
      bContinue = False
      i = 1

      iRemoveBeg = 0
      iRemoveCnt = 0

      iDeclLine = VBA_GetCountOfDeclarationLines(oMod)
      propState("CODE_LINE") = 0
      propState("CODE_STATE") = "<INIT>"
      bRet = bRet And EditCodeModule_RemoveMode("TEST", "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_RemoveMode("DEBUG", "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_Import(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_OptionStatements(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_Global(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_GlobalInsert(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_Autocontrols(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_AutocontrolsInsert(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_AutoMacro(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_AutoMacroInsert(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_ConvPrivateScope(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)

      Do While i <= .CountOfLines
         bRemove = False
         sLine = .Lines(i, 1)
         If bContinue And iComment > 0 Then iComment = 1 Else iComment = VBA_GetCommentOfLine(sLine)
         If iComment > 0 Then sComment = Mid(sLine, iComment) Else sComment = ""
         If iComment > 0 Then sLine = Mid(sLine, 1, iComment - 1)

         propState("CODE_LINE") = i
         If i <= iDeclLine Then
            propState("CODE_STATE") = "<DECLARE>"
         ElseIf i > iDeclLine Then
            sProcName = .ProcOfLine(i, iType)
            iProcBodyLine = .ProcBodyLine(sProcName, iType)
            If iProcBodyLine > 0 Then
               propState("CODE_STATE") = "<PROC:" & sProcName & ":" & CStr(i >= iProcBodyLine) & ">"
            Else
               propState("CODE_STATE") = "<PROC:" & sProcName & ":>"
            End If
         End If

         bRet = bRet And EditCodeModule_RemoveMode("TEST", sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo Continue
         bRet = bRet And EditCodeModule_RemoveMode("DEBUG", sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo Continue
         bRet = bRet And EditCodeModule_Import(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo Continue
         bRet = bRet And EditCodeModule_OptionStatements(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo Continue
         bRet = bRet And EditCodeModule_Global(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo Continue
         bRet = bRet And EditCodeModule_GlobalInsert(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo Continue
         bRet = bRet And EditCodeModule_Autocontrols(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo Continue
         bRet = bRet And EditCodeModule_AutocontrolsInsert(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo Continue
         bRet = bRet And EditCodeModule_AutoMacro(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo Continue
         bRet = bRet And EditCodeModule_AutoMacroInsert(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo Continue
         bRet = bRet And EditCodeModule_ConvPrivateScope(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo Continue

         If propGlobal("IGNORE_COMMENT") Then
            If propGlobal("APPEND_INFO") And sComment Like "*:[#][#][#]INSERTED[#][#][#]" Then
               sComment = Replace(sComment, ":###INSERTED###", "")
            Else
               sComment = ""
            End If
         End If

         If Trim(sLine & sComment) = "" And propGlobal("IGNORE_BLANK") Then
            bRemove = True
            GoTo Continue
         End If

         If .Lines(i, 1) <> sLine & sComment Then .ReplaceLine i, sLine & sComment

Continue:
         If Not bRet Then Exit Do
         If bRemove Then
            If iRemoveBeg = 0 Then
               iRemoveBeg = i
               iRemoveCnt = 1
            Else
               iRemoveCnt = iRemoveCnt + 1
            End If
         Else
            If iRemoveBeg <> 0 Then
               .DeleteLines iRemoveBeg, iRemoveCnt
               i = i - iRemoveCnt
               iRemoveBeg = 0
               iRemoveCnt = 0
               iDeclLine = VBA_GetCountOfDeclarationLines(oMod)
            End If
         End If
         bContinue = (Right(Trim(.Lines(i, 1)), 1) = "_")
         i = i + 1
      Loop
      If iRemoveBeg <> 0 Then .DeleteLines iRemoveBeg, iRemoveCnt

      propState("CODE_STATE") = "<END>"
      propState("CODE_LINE") = i
      bRet = bRet And EditCodeModule_RemoveMode("TEST", "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_RemoveMode("DEBUG", "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_Import(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_OptionStatements(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_Global(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_GlobalInsert(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_Autocontrols(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_AutocontrolsInsert(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_AutoMacro(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_AutoMacroInsert(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_ConvPrivateScope(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
   End With
   EditCodeModule = bRet
End Function
