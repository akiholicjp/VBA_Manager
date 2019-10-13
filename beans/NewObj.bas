' VBA: Import Wrapper/_rtcCallByName.bas

Function NewObj(ByVal o As Object, ParamArray params() As Variant) As Object
   Dim iNum As Long
   Dim args() As Variant
   Dim i As Long

   iNum = UBound(params)
   If iNum < 0 Then
      o.Init
      Set NewObj = o
      Exit Function
   End If

   ReDim args(iNum)
   For i = 0 To iNum
      If IsObject(params(i)) Then Set args(i) = params(i) Else args(i) = params(i)
   Next
   rtcCallByName o, StrPtr("Init"), VbMethod, args
   Set NewObj = o
End Function
