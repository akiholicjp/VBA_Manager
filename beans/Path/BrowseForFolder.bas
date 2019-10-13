Function BrowseForFolder(Optional sTitle As String = "フォルダの選択", Optional vRootFolder As Variant) As String
   Dim objFolder As Object
   Set objFolder = CreateObject("Shell.Application").BrowseForFolder(hWndAccessApp, sTitle, &H10, vRootFolder)
   If Not objFolder Is Nothing Then
      BrowseForFolder = objFolder.Self.Path
      Set objFolder = Nothing
   End If
End Function

' BIF_RETURNONLYFSDIRS	&H1	●ファイルシステムのフォルダー
' BIF_DONTGOBELOWDOMAIN	&H2	△ネットワークフォルダを含めない
' BIF_STATUSTEXT	&H4	×ステータステキストを設定
' * 要Callback（API SHBrowseForFolder仕様）
' BIF_RETURNFSANCESTORS	&H8	△ルート選択不可
' BIF_EDITBOX	&H10	●名称ボックス表示
' BIF_VALIDATE	&H20	×選択アイテムの妥当性チェック
' * 要Callback（API SHBrowseForFolder仕様）
' BIF_NEWDIALOGSTYLE	&H40	×新しいスタイル表示
' BIF_BROWSEINCLUDEURLS	&H80	△URLを対象にできる
' BIF_UAHINT	&H100	△ヒントを表示（文言が変化しないので効果は期待できず）
' BIF_NONEWFOLDERBUTTON	&H200	●「新しいフォルダの作成」を表示しない
' BIF_NOTRANSLATETARGETS	&H400	△ショートカットのターゲットのPIDLを返します…詳細不明
' BIF_BROWSEFORCOMPUTER	&H1000	△ネットワークが対象
' CSIDL_NETWORK (18) 併用
' BIF_BROWSEFORPRINTER	&H2000	△プリンターを対象
' CSIDL_PRINTERS (4) 併用
' プリンターを選択してもエラーです（意味なし）

' BIF_BROWSEINCLUDEFILES	&H4000	●ファイルも表示する
' ファイルを選択してもエラーです（意味なし）

' BIF_SHAREABLE	&H8000	△共有可能なリソースを表示できる
' （なぜか仕様未掲載のZIP,LZHなど圧縮ファイル展開もする）
' 通常は↓こう展開するが

' この設定だと、↓こう展開する

' BIF_BROWSEFILEJUNCTIONS	&H10000	△ZIP,LZHなど圧縮書庫ファイルも表示＆展開＆選択が可能

' その他の圧縮形式 7z や Rar は表示されず
' BIF_BROWSEINCLUDEFILES を組み合わせても選択できません（エラー発生）
