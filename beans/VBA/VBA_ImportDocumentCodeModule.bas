' VBA: Import ../GetFSO.bas

Function VBA_ImportDocumentCodeModule(ByRef oComps As Object, ByVal sPath As String, ByRef msgError As String) As Object
   Set VBA_ImportDocumentCodeModule = Nothing
   Dim o As Object
   Dim sName As String
   Dim bExist As Boolean
   sName = GetFSO().GetBaseName(sPath)
   bExist = False
   For Each o In oComps
      If o.Name = sName Then
         bExist = True
         Exit For
      End If
   Next o
   If Not bExist Then
      msgError = msgError & "インポート先のDocumentが見つかりません: " & sName & vbCrLf
      Exit Function
   End If
   With o.CodeModule
      .DeleteLines StartLine:=1, Count:=.CountOfLines
      .AddFromFile sPath
      .DeleteLines StartLine:=1, Count:=4
   End With
   Set VBA_ImportDocumentCodeModule = o
End Function

