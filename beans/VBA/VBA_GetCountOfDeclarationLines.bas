Function VBA_GetCountOfDeclarationLines(ByRef oMod As Object) As Long
   Dim iDeclPos As Long
   Dim sProcName As String
   Dim i As Long
   Dim iType As Long

   With oMod.CodeModule
      iDeclPos = .CountOfDeclarationLines
      If .CountOfLines - iDeclPos > 0 Then
         sProcName = .ProcOfLine(iDeclPos + 1, iType)
         For i = .ProcStartLine(sProcName, iType) To .ProcBodyLine(sProcName, iType) - 1
            If (Trim(.Lines(i, 1)) Like "[#]End*") Then iDeclPos = i
         Next i
      End If
   End With
   VBA_GetCountOfDeclarationLines = iDeclPos
End Function
