' VBA: Import ../RegTest.bas

Public Sub VBA_RemoveLineNumbers(oMod As Object)
   Dim i As Long
   Dim sProcName As String
   Dim sLine As String
   Dim iProcStartLine As Long
   Dim iProcEndLine As Long
   Dim bProc As Boolean
   Dim bTarget As Boolean

   With oMod.CodeModule
      bProc = False
      bTarget = False
      For i = 1 To .CountOfDeclarationLines
         If RegTest(.Lines(i, 1), "^\s*'.*Auto Added Line Numbers") Then
            Call .DeleteLines(i)
            bTarget = True
            Exit For
         End If
      Next i
      If Not bTarget Then Exit Sub
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
                  GoTo Continue
               ElseIf RegTest(sLine, "^\s*End\s+(Sub|Function|Property)\s*") Then
                  bProc = False
                  GoTo Continue
               End If
               If bProc Then
                  .ReplaceLine i, Mid(sLine, 7)
               End If
            End If
         End If
Continue:
      Next i
   End With
End Sub
