Function VBA_GetDicProc(ByRef oMod As Object) As Object
   Dim lp_line As Long
   Dim prcName As String
   Dim sName As String
   Dim dicProc As Object
   Dim fp As Long
   Dim iType As Long

   sName = VBA_GetModuleName(oMod, IgnoreBlankDcm:=True)
   If sName <> "" Then
      Set dicProc = CreateObject("Scripting.Dictionary")
   End If
   With oMod.CodeModule
      prcName = ""
      For lp_line = 1 To .CountOfLines
         If prcName <> .ProcOfLine(lp_line, iType) Then
            prcName = .ProcOfLine(lp_line, iType)
            dicProc.Add Key:=prcName, Item:=Array(lp_line, Trim(.Lines(.ProcBodyLine(prcName, iType), 1)))
         End If
      Next lp_line
   End With
   Set VBA_GetDicProc = dicProc
End Function
