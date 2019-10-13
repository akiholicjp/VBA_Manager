' ������̋ߎ������擾�i���I�v��@�ɂ�郌�[�x���V���^�C�������̎Z�o�j
' ����
'   str1 ��r�Ώە�����1
'   str2 ��r�Ώە�����2
' �߂�l
'   ���[�x���V���^�C������
Function GetLevenShteinDistance(ByVal str1 As String, ByVal str2 As String) As Double
   Dim i As Long, j As Long
   Dim len1 As Long
   Dim len2 As Long
   Dim distance() As Long
   Dim cost As Long

   len1 = Len(str1)
   len2 = Len(str2)
   ReDim distance(len1, len2)

   For i = 0 To len1
      distance(i, 0) = i
   Next

   For j = 0 To len2
      distance(0, j) = j
   Next

   For i = 1 To len1
      For j = 1 To len2
         If asc(Mid$(str1, i, 1)) = asc(Mid$(str2, j, 1)) Then cost = 0 Else cost = 1

         If (distance(i - 1, j) + 1) < (distance(i, j - 1) + 1) Then
            distance(i, j) = distance(i - 1, j) + 1
         Else
            distance(i, j) = distance(i, j - 1) + 1
         End If
         If (distance(i - 1, j - 1) + cost) < distance(i, j) Then
            distance(i, j) = distance(i - 1, j - 1) + cost
         End If
      Next
   Next

   GetLevenShteinDistance = distance(len1, len2)
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetLevenShteinDistance()
   assert 0, GetLevenShteinDistance("abc", "abc")
   assert 3, GetLevenShteinDistance("abc", "")
   assert 3, GetLevenShteinDistance("", "abc")
   assert 2, GetLevenShteinDistance("abc", "cba")
End Sub