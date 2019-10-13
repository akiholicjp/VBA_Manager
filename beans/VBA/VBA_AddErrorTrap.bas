'<TODO>
'SSS_ERR_TRAP_CALLなどが処理されないようにする必要あり。
'現状は、"*ERR_TRAP_CALL*" というProcは処理しないように作り込んでいる。

Private Sub VBA_AddErrorTrap(oMod As Object)
   Dim i As Long
   Dim sProcName As String
   Dim sLine As String
   Dim iProcStartLine As Long
   Dim bProc As Boolean
   Dim regResults As Object
   Dim bEndProcExist As Boolean
   Dim sProcType As String
   Dim bErrTrapExist As Boolean
   Dim bMultiProcDef As Boolean
   Dim bIgnoreProc As Boolean

   With oMod.CodeModule
      bProc = False
      For i = 1 To .CountOfDeclarationLines
         sLine = .Lines(i, 1)
         If RegTest(sLine, "^\s*'.*\s*VBA\s*:\s*Auto Added Error Trap Code") Then Exit Sub
      Next i
      Call .InsertLines(1, "'!=!=!=!=! VBA: Auto Added Error Trap Code !=!=!=!=!")
      Call .InsertLines(2, "Private Const SSS_MOD_NAME = """ & oMod.Name & """")

      i = .CountOfDeclarationLines + 1
      Do While i <= .CountOfLines
         If .ProcOfLine(i, 0) <> vbNullString Then
            If sProcName <> .ProcOfLine(i, 0) Then
               sProcName = .ProcOfLine(i, 0)
               iProcStartLine = .ProcStartLine(sProcName, 0)
               bEndProcExist = False
               bErrTrapExist = False
               If sProcName Like "*ERR_TRAP_CALL*" Then
                  bIgnoreProc = True
               Else
                  bIgnoreProc = False
               End If
            End If
            If bIgnoreProc Then GoTo Continue
            If iProcStartLine <= i Then
               sLine = .Lines(i, 1)
               If RegTest(sLine, "^\s*(Private\s+|Public\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property)\s*.*") Or bMultiProcDef Then
                  If sLine Like "* _" Then
                     bMultiProcDef = True
                  Else
                     Call .InsertLines(i + 1, "Const SSS_PROC_NAME = """ & sProcName & """: On Error GoTo SSS_ERR_TRAP")
                     i = i + 1
                     bMultiProcDef = False
                     bProc = True
                  End If
                  GoTo Continue
               ElseIf RegExe(sLine, "^\s*End\s+(Sub|Function|Property)\s*", regResults) Then
                  sProcType = regResults(1)
                  If Not bErrTrapExist Then
                     If bEndProcExist Then
                        Call .InsertLines(i, "Exit " & sProcType)
                     Else
                        Call .InsertLines(i, "SSS_END_PROC: Exit " & sProcType)
                     End If
                     Call .InsertLines(i + 1, "SSS_ERR_TRAP: Call SSS_ERR_TRAP_CALL(SSS_PROC_NAME, SSS_MOD_NAME): Resume SSS_END_PROC")
                     i = i + 2
                  End If
                  bProc = False
                  GoTo Continue
               ElseIf RegTest(sLine, "^\s*SSS_END_PROC\s*:.*") Then
                  bEndProcExist = True
               ElseIf RegTest(sLine, "^\s*SSS_ERR_TRAP\s*:") Then
                  bErrTrapExist = True
               End If
            End If
         End If
Continue:
         i = i + 1
      Loop
   End With
End Sub
