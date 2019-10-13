' VBA: Import GetLevenShteinDistance.bas

' 文字列の近似率を取得（動的計画法によるレーベンシュタイン距離の算出）
' 引数
'   str1 比較対象文字列1
'   str2 比較対象文字列2
' 戻り値
'   近似率（0.0～1.0）
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

   'レーベンシュタイン距離
   Dim levenshtein As Double
   levenshtein = GetLevenShteinDistance(str1, str2)

   '文字列の近似率
   If len1 > len2 Then
      GetStringMatchRate = (len1 - levenshtein) / len1
   Else
      GetStringMatchRate = (len2 - levenshtein) / len2
   End If
End Function
