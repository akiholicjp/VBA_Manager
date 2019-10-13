' VBA: Import ../GetFSO.bas
' VBA: Import VBA_ImportDocumentCodeModule.bas

Function VBA_ImportCodeModule(ByRef oComps As Object, ByVal sPath As String, ByRef msgError As String) As Object
   Set VBA_ImportCodeModule = Nothing
   Dim oComp As Object
   Select Case GetFSO().GetExtensionName(sPath)
   Case "bas", "cls", "frm"
      Set oComp = oComps.Import(sPath)
      If oComp Is Nothing Then Exit Function
   Case "dcm"
      Set oComp = VBA_ImportDocumentCodeModule(oComps, sPath, msgError)
      If oComp Is Nothing Then Exit Function
   Case Else
      msgError = msgError & "VBAモジュールファイルの拡張子が不正です: " & sPath & vbCrLf
      Exit Function
   End Select
   Set VBA_ImportCodeModule = oComp
End Function
