Function NewRegExp(Optional ByVal bGlobal As Boolean = True, Optional ByVal MultiLine As Boolean = True, Optional ByVal IgnoreCase As Boolean = False, Optional ByVal sPattern As String = "") As Object
   Set NewRegExp = CreateObject("VBScript.RegExp")
   With NewRegExp
      .Global = bGlobal
      .MultiLine = MultiLine
      .IgnoreCase = IgnoreCase
      .Pattern = sPattern
   End With
End Function
