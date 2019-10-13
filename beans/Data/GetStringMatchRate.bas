' VBA: Import GetLevenShteinDistance.bas

' ������̋ߎ������擾�i���I�v��@�ɂ�郌�[�x���V���^�C�������̎Z�o�j
' ����
'   str1 ��r�Ώە�����1
'   str2 ��r�Ώە�����2
' �߂�l
'   �ߎ����i0.0�`1.0�j
Function GetStringMatchRate(ByVal str1 As String, ByVal str2 As String) As Double
   Dim len1 As Long
   Dim len2 As Long

   len1 = Len(str1)
   len2 = Len(str2)
   If (len1 = 0) And (len2 = 0) Then
      GetStringMatchRate = 1
      Exit Function
   ElseIf (len1 = 0) Or (len2 = 0) Then
      GetStringMatchRate = 0
      Exit Function
   End If

   '���[�x���V���^�C������
   Dim levenshtein As Double
   levenshtein = GetLevenShteinDistance(str1, str2)

   '������̋ߎ���
   If len1 > len2 Then
      GetStringMatchRate = (len1 - levenshtein) / len1
   Else
      GetStringMatchRate = (len2 - levenshtein) / len2
   End If
End Function
