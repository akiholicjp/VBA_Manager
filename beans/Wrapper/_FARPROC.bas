#If Win64 Then
Public Function FARPROC(pfn As LongPtr) As LongPtr    ''AddressOf���Z�q�̖߂�l��߂��֐�
   FARPROC = pfn
End Function
#Else
Public Function FARPROC(pfn As Long) As Long    ''AddressOf���Z�q�̖߂�l��߂��֐�
   FARPROC = pfn
End Function
#End If
