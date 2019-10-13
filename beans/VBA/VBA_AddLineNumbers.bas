' VBA: Import ../RegTest.bas

Private Sub VBA_AddLineNumbers(oMod As Object)
   Dim i As Long, iOff As Long
   Dim sProcName As String
   Dim sLine As String
   Dim iProcStartLine As Long
   Dim iProcEndLine As Long
   Dim bProc As Boolean
   Dim bMultiLine As Boolean

   With oMod.CodeModule
      Select Case oMod.Type
      Case 1: iOff = 10001 ' Module.Standard
      Case 2: iOff = 10009 ' Module.Class
      Case 3: iOff = 10015 ' Module.Forms
      Case 100: iOff = 10009 ' Module.Document
      Case Else: iOff = 10000
      End Select
      bProc = False
      bMultiLine = False
      For i = 1 To .CountOfDeclarationLines
         If RegTest(.Lines(i, 1), "^\s*'.*Auto Added Line Numbers") Then Exit Sub
      Next i
      Call .InsertLines(1, "'!=!=!=!=! Auto Added Line Numbers !=!=!=!=!")
      For i = .CountOfDeclarationLines + 1 To .CountOfLines
         If .ProcOfLine(i, 0) <> vbNullString Then
            If sProcName <> .ProcOfLine(i, 0) Then
               sProcName = .ProcOfLine(i, 0)
               iProcStartLine = .ProcStartLine(sProcName, 0)
               iProcEndLine = .ProcCountLines(sProcName, 0) + iProcStartLine - 1
            End If
            If iProcStartLine <= i And i <= iProcEndLine Then
               sLine = .Lines(i, 1)

               If RegTest(sLine, "^\s*(Private\s+|Public\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property)\s*") Then
                  bProc = True
                  If RegTest(sLine, "_\s*$") Then bMultiLine = True
                  GoTo Continue
               ElseIf RegTest(sLine, "^\s*End\s+(Sub|Function|Property)\s*") Then
                  bProc = False
                  If RegTest(sLine, "_\s*$") Then bMultiLine = True
                  GoTo Continue
               End If

               If bProc Then
                  If bMultiLine Then
                     .ReplaceLine i, Space(5) & " " & sLine
                     If Not RegTest(sLine, "_\s*$") Then bMultiLine = False
                     GoTo Continue
                  End If

                  If RegTest(sLine, "_\s*$") Then bMultiLine = True

                  If RegTest(sLine, "^\s*(Case)\s*") Then
                     .ReplaceLine i, Space(5) & " " & sLine
                     GoTo Continue
                  ElseIf RegTest(sLine, "^\s*#") Then
                     .ReplaceLine i, Space(5) & " " & sLine
                     GoTo Continue
                  End If
                  .ReplaceLine i, CStr(i + iOff) & Space(5 - Len(CStr(i + iOff))) & " " & sLine
               End If
            End If
         End If
Continue:
      Next i
   End With
End Sub

