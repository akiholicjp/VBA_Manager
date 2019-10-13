Function VBA_RemoveProcFromCodeModule(ByRef oMod As Object, ByVal sNameLike As String) As Boolean
   Dim sProcName As String
   Dim iLine As Long
   Dim iLineFirst As Long
   Dim iLineEnd As Long

   With oMod.CodeModule
      sProcName = ""
      iLine = 1
      iLineFirst = 1
      iLineEnd = 1
      Do While iLine <= .CountOfLines
         If sProcName <> .ProcOfLine(iLine, 0) Then
            If sProcName Like sNameLike Then
               .DeleteLines StartLine:=iLineFirst, Count:=iLineEnd - iLineFirst + 1
               iLine = iLineFirst
            Else
               iLineFirst = iLine
               iLine = iLine + 1
            End If
            sProcName = .ProcOfLine(iLine, 0)
         Else
            iLineEnd = iLine
            iLine = iLine + 1
         End If
      Loop
      If sProcName Like sNameLike Then
         .DeleteLines StartLine:=iLineFirst, Count:=iLineEnd - iLineFirst + 1
      End If
   End With
   VBA_RemoveProcFromCodeModule = True
End Function
