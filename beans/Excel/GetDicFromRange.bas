' VBA: Import ../Data/ToCollection.bas
Function GetDicFromRange(oRange As Range, ByVal key_columns As Variant, ByVal item_columns As Variant, Optional ByVal begin_row As Long = -1, Optional ByVal end_row As Long = -1, Optional ByVal bKeepCollection As Boolean = False, Optional ByRef dic As Object) As Object
   Dim oDic As Object
   Dim iR As Long
   Dim sKey As String
   Dim vKeyCol As Variant, vItemCol As Variant
   Dim oCol As Object
   Dim i As Long

   Set key_columns = ToCollection(key_columns)
   If dic Is Nothing Then Set dic = CreateObject("Scripting.Dictionary")
   With oRange
      If end_row < 0 Then
         If .Rows.Count > .SpecialCells(xlCellTypeLastCell).Row Then end_row = .SpecialCells(xlCellTypeLastCell).Row Else end_row = .Rows.Count
      End If
      If begin_row < 0 Then begin_row = 1

      For iR = begin_row To end_row
         Set oDic = dic
         For i = 1 To key_columns.Count: vKeyCol = key_columns(i)
            sKey = Trim(.Cells(iR, vKeyCol).Text)
            If sKey <> "" Then
               If i = key_columns.Count Then
                  If oDic.Exists(sKey) Then
                     ' èdï°ÇµÇΩèÍçáÇÕñ≥éã
                     sKey = ""
                  Else
                     oDic.Add Key:=sKey, Item:=GetDicFromRangeSub(oRange, iR, item_columns, bKeepCollection)
                  End If
                  Exit For
               Else
                  If Not oDic.Exists(sKey) Then oDic.Add Key:=sKey, Item:=CreateObject("Scripting.Dictionary")
                  Set oDic = oDic(sKey)
               End If
            End If
         Next i
      Next iR
   End With
   Set GetDicFromRange = dic
End Function

Private Function GetDicFromRangeSub(oRange As Range, ByVal iR As Long, ByRef item_columns As Variant, ByVal bKeepCollection As Boolean) As Variant
   Dim i As Long
   Dim o As Object
   Dim vKey As Variant
   If TypeName(item_columns) = "Dictionary" Then
      If item_columns.Count = 1 And Not bKeepCollection Then
         GetDicFromRangeSub = Trim(oRange.Cells(iR, item_columns.Items(1)).Text)
      Else
         Set o = CreateObject("Scripting.Dictionary")
         With oRange
            For Each vKey In item_columns.Keys
               o.Add Key:=vKey, Item:=Trim(.Cells(iR, item_columns(vKey)).Text)
            Next vKey
         End With
         Set GetDicFromRangeSub = o
      End If
   Else
      Set item_columns = ToCollection(item_columns)
      If item_columns.Count = 1 And Not bKeepCollection Then
         GetDicFromRangeSub = Trim(oRange.Cells(iR, item_columns(1)).Text)
      Else
         Set o = New Collection
         With oRange
            For i = 1 To item_columns.Count
               o.Add Trim(.Cells(iR, item_columns(i)).Text)
            Next i
         End With
         Set GetDicFromRangeSub = o
      End If
   End If
End Function

' =================== VBA: TEST: Begin ===================

' VBA: Import ../DicProp.bas

Public Sub xUnitTest_beans_GetDicFromRange_Sheet()
   Dim o1 As Object
   Dim o2 As Object
   Dim o3 As Object
   Dim o4 As Object
   Dim o5 As Object
   Set o1 = GetDicFromRange( _
      ThisWorkbook.Worksheets("GetDictionaryArrayFromWorksheet").Cells, _
      key_columns:=Array(3), _
      item_columns:=Array(4, 5, 7), _
      begin_row:=2, _
      end_row:=5, _
      bKeepCollection:=True)
   assert "{""LABEL"":[""aaa"",""bbb"",""ddd""],""A"":[""1"",""11"",""31""],""B"":[""2"",""12"",""32""],""C"":[""3"",""13"",""33""]}", Dump(o1)

   Call GetDicFromRange( _
      ThisWorkbook.Worksheets("GetDictionaryArrayFromWorksheet").Cells, _
      key_columns:=Array(3), _
      item_columns:=Array(4, 5, 7), _
      begin_row:=3, _
      end_row:=5, _
      bKeepCollection:=True, _
      dic:=o2)
   assert "{""A"":[""1"",""11"",""31""],""B"":[""2"",""12"",""32""],""C"":[""3"",""13"",""33""]}", Dump(o2)

   Set o3 = GetDicFromRange( _
      ThisWorkbook.Worksheets("GetDictionaryArrayFromWorksheet").Cells, _
      key_columns:=3, _
      item_columns:=5, _
      begin_row:=2, _
      bKeepCollection:=False, _
      end_row:=5)
   assert "{""LABEL"":""bbb"",""A"":""11"",""B"":""12"",""C"":""13""}", Dump(o3)

   Call GetDicFromRange( _
      ThisWorkbook.Worksheets("GetDictionaryArrayFromWorksheet").Cells, _
      key_columns:=3, _
      item_columns:=5, _
      begin_row:=3, _
      end_row:=5, _
      bKeepCollection:=False, _
      dic:=o4)
   assert "{""A"":""11"",""B"":""12"",""C"":""13""}", Dump(o4)

   Set o5 = GetDicFromRange( _
      ThisWorkbook.Worksheets("GetDictionaryArrayFromWorksheet").Cells, _
      key_columns:=Array(3), _
      item_columns:=DicProp("K1", 4, "K2", 5, "K3", 7), _
      begin_row:=2, _
      end_row:=5, _
      bKeepCollection:=True)
   assert "{""LABEL"":{""K1"":""aaa"",""K2"":""bbb"",""K3"":""ddd""},""A"":{""K1"":""1"",""K2"":""11"",""K3"":""31""},""B"":{""K1"":""2"",""K2"":""12"",""K3"":""32""},""C"":{""K1"":""3"",""K2"":""13"",""K3"":""33""}}", Dump(o5)
End Sub

Public Sub xUnitTest_beans_GetDicFromRange_Range()
   Dim o1 As Object
   Dim o2 As Object
   Dim o3 As Object
   Dim o4 As Object
   Dim o5 As Object
   Set o1 = GetDicFromRange( _
      ThisWorkbook.Worksheets("GetDictionaryArrayFromWorksheet").[C2:H19], _
      key_columns:=Array(1), _
      item_columns:=Array(2, 3, 5), _
      begin_row:=1, _
      end_row:=4, _
      bKeepCollection:=True)
   assert "{""LABEL"":[""aaa"",""bbb"",""ddd""],""A"":[""1"",""11"",""31""],""B"":[""2"",""12"",""32""],""C"":[""3"",""13"",""33""]}", Dump(o1)

   Call GetDicFromRange( _
      ThisWorkbook.Worksheets("GetDictionaryArrayFromWorksheet").[C2:H19], _
      key_columns:=Array(1), _
      item_columns:=Array(2, 3, 5), _
      begin_row:=2, _
      end_row:=4, _
      bKeepCollection:=True, _
      dic:=o2)
   assert "{""A"":[""1"",""11"",""31""],""B"":[""2"",""12"",""32""],""C"":[""3"",""13"",""33""]}", Dump(o2)

   Set o3 = GetDicFromRange( _
      ThisWorkbook.Worksheets("GetDictionaryArrayFromWorksheet").[C2:H19], _
      key_columns:=1, _
      item_columns:=3, _
      begin_row:=1, _
      bKeepCollection:=False, _
      end_row:=4)
   assert "{""LABEL"":""bbb"",""A"":""11"",""B"":""12"",""C"":""13""}", Dump(o3)

   Call GetDicFromRange( _
      ThisWorkbook.Worksheets("GetDictionaryArrayFromWorksheet").[C2:H19], _
      key_columns:=1, _
      item_columns:=3, _
      begin_row:=2, _
      end_row:=4, _
      bKeepCollection:=False, _
      dic:=o4)
   assert "{""A"":""11"",""B"":""12"",""C"":""13""}", Dump(o4)

   Set o5 = GetDicFromRange( _
      ThisWorkbook.Worksheets("GetDictionaryArrayFromWorksheet").[C2:H19], _
      key_columns:=Array(1), _
      item_columns:=DicProp("K1", 2, "K2", 3, "K3", 5), _
      begin_row:=1, _
      end_row:=4, _
      bKeepCollection:=True)
   assert "{""LABEL"":{""K1"":""aaa"",""K2"":""bbb"",""K3"":""ddd""},""A"":{""K1"":""1"",""K2"":""11"",""K3"":""31""},""B"":{""K1"":""2"",""K2"":""12"",""K3"":""32""},""C"":{""K1"":""3"",""K2"":""13"",""K3"":""33""}}", Dump(o5)
End Sub

' VBA: Import GetTitleColumns.bas
Public Sub xUnitTest_beans_GetDicFromRange_Title()
   Dim o As Object
   Dim oRng As Range
   Set oRng = ThisWorkbook.Worksheets("GetDictionaryArrayFromWorksheet").[C2:H19]
   Set o = GetDicFromRange( _
      oRng, _
      key_columns:=Array(1), _
      item_columns:=GetTitleColumns(oRng, column_From:=2, column_End:=4, idx_Row:=1), _
      begin_row:=2, _
      end_row:=3, _
      bKeepCollection:=True)
   assert "{""A"":{""aaa"":""1"",""bbb"":""11"",""ccc"":""21""},""B"":{""aaa"":""2"",""bbb"":""12"",""ccc"":""22""}}", Dump(o)

   Set oRng = ThisWorkbook.Worksheets("GetDictionaryArrayFromWorksheet").[C2:F4]
   Set o = GetDicFromRange( _
      oRng, _
      key_columns:=1, _
      item_columns:=GetTitleColumns(oRng), _
      begin_row:=2 _
   )
   assert "{""A"":{""LABEL"":""A"",""aaa"":""1"",""bbb"":""11"",""ccc"":""21""},""B"":{""LABEL"":""B"",""aaa"":""2"",""bbb"":""12"",""ccc"":""22""}}", Dump(o)
End Sub

