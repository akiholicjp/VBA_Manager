' VBA: Import Wrapper/_CoCreateGuid.bas
' VBA: Import IIf.bas

Function GetGUID() As String
   Dim udtGUID As GUID
   If (CoCreateGuid(udtGUID) = 0) Then
   GetGUID = _
      String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & "-" & _
      String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & "-" & _
      String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & "-" & _
      IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
      IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & "-" & _
      IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
      IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
      IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
      IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
      IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
      IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetGUID()
   Dim s As String
   s = GetGUID()
   assert Len(s), 36
   assertTrue s <> GetGUID()
End Sub
