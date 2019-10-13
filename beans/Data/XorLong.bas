Function XorLong(ByVal a As Long, ByVal b As Long) As Long
   XorLong = a Xor b
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_XorLong()
   assert &H0000, XorLong(&H1234, &H1234)
   assert &H1000, XorLong(&H1234, &H0234)
End Sub
