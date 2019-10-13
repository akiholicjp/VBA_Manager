#If Win64 Then
Public Function FARPROC(pfn As LongPtr) As LongPtr    ''AddressOf演算子の戻り値を戻す関数
   FARPROC = pfn
End Function
#Else
Public Function FARPROC(pfn As Long) As Long    ''AddressOf演算子の戻り値を戻す関数
   FARPROC = pfn
End Function
#End If
