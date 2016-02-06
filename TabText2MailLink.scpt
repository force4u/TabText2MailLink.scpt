(*
TabText2MailLink
タブ区切りテキストからHTMLを生成します。
2016026　初回作成

*)



------------------------------------ ダブルクリックの始まり
on run
	---プロンプトの文言改行が使えます\nを入れます
	set theWithPrompt to "タブテキストをメールリンクに変換"
	---ファイル選択ダイアログのデフォルトのディレクトリ
	tell application "Finder"
		---ダイアログの開き先を指定
		set theDefaultLocation to (container of (path to me)) as alias
	end tell
	---Uniform Type Identifier指定
	set theFileTypeList to "public.plain-text,com.apple.traditional-mac-plain-text" as text
	---↑のファイルタイプをリスト形式に整形する
	---Uniform Type Identifier用に区切り文字をカンマに
	set AppleScript's text item delimiters to {","}
	---Uniform Type Identifierをリストに格納する
	set theFileTypeList to every text item of theFileTypeList
	---ダイアログを出して選択されたファイルは「open」に渡す
	open (choose file default location theDefaultLocation ¬
		with prompt theWithPrompt ¬
		of type theFileTypeList ¬
		invisibles true ¬
		with multiple selections allowed without showing package contents)
end run


on open objOpenDrop
	---Openするファイルをエイリアスで取得
	set theFileAlias to objOpenDrop as alias
	---選択したファイルのUNIXパスを取得
	set theUnixPass to POSIX path of theFileAlias as text
	---ファイルをUTF8形式で読み込み
	set theData to (do shell script "cat '" & theUnixPass & "'") as «class utf8»
	---区切り文字をMac改行に変換
	set AppleScript's text item delimiters to {"\r"}
	---改行毎にリスト形式に格納
	set retListVal to (every text item of theData) as list
	-----行数カウンタ初期化
	set numListLine to count of retListVal
	---読み込み開始行を初期化
	set numLine to 1 as number
	--タブインデックス用のカウンタ初期化
	set numTabIndex to 1 as number
	---データ格納先を初期化
	set theLineData1 to "" as text
	set theLineData2 to "" as text
	set theLineData3 to "" as text
	---HTMLヘッダー
	set theHtmlData to "<!-- ここからEmailアドレステーブル -->\r"
	---CSSを設定
	set theHtmlData to theHtmlData & "<style type=\"text/css\" media=\"screen\">div.bordertable{width: 80%;} .bordertable table , .bordertable th , .bordertable td {font-size: 12px;border: solid 1px #666666;padding: 3px;border-collapse: collapse;overflow: inherit;word-wrap: break-word;} .bordertable th.mailname{font-size: small;text-align: left;} .bordertable a.emailadd{font-family: monospace;font-style: normal;　word-break: keep-all;} .bordertable p{font-size: small;text-align: center;}</style>" as text
	---CSSの反映用のDIVの開始
	set theHtmlData to theHtmlData & "<div class=\"bordertable\">" as text
	set theHtmlData to theHtmlData & "<p> LiveMailを利用している人はLiveMailをクリック</p>\r" as text
	set theHtmlData to theHtmlData & "<table>" as text
	set theHtmlData to theHtmlData & "<caption>メールアドレス一覧｜個人情報につき取り扱いには注意しましょう</caption>" as text
	---テーブルの見出
	set theHtmlData to theHtmlData & "<thead title=\"見出し\"><tr><th>氏名</th><th>所属</th><th colspan=\"2\">メール</th></tr></thead>" as text
	---tbody
	set theHtmlData to theHtmlData & "<tbody title=\"メールアドレス一覧\">" as text
	---繰り返しのはじまり
	repeat numListLine times
		---最初の行を読み込む
		set theDataListLine to (item numLine of retListVal) as text
		---区切り文字をタブに指定
		set AppleScript's text item delimiters to {"\t"}
		---タブ毎にリスト形式で取得
		set theDataListLineData to every text item of theDataListLine as list
		------ここは使わないけど　リストの項目数を取得
		set numListLineCnt to (count of theDataListLineData) as number
		-------各項目を取得
		set theLineData1 to (item 1 of theDataListLineData) as text ---名前
		set theLineData2 to (item 2 of theDataListLineData) as text ---所属
		set theLineData3 to (item 3 of theDataListLineData) as text ---メール
		---項目３をメールアドレスとする
		set theOrgEmail to theLineData3 as text
		---リンク用の名前を設定（所属+名前+さん）
		set theLinkName to ("【" & theLineData2 & "】" & theLineData1 & "さん") as text
		
		---///Macと最近のWindows用にURLエンコード
		set theLinkNameMac to my doUrlEncode(theLinkName) as text
		---URLエンコードのスペース表記+を取る
		set theLinkNameMac to my doReplace(theLinkNameMac, "+", "") as text
		---最初と最後の<>
		set theLinkNameMac to (theLinkNameMac & "%3C" & theOrgEmail & "%3E") as text
		
		---///WindowsLiveMai用のSJISエンコード
		set theLinkNameWin to my doSJISencode(theLinkName) as text
		---最初と最後の<>
		set theLinkNameWin to (theLinkNameWin & "%3C" & theOrgEmail & "%3E") as text
		
		---///body入りのリンク作成
		set theLinkBodyName to ("TO:\r【" & theLineData2 & "】" & theLineData1 & "さん") as text
		---Macと最近のWindows用にURLエンコード
		set theLinkBodyNameMac to my doUrlEncode(theLinkBodyName) as text
		
		---HTMLにする
		set theHtmlData to theHtmlData & "<tr>" as text
		set theHtmlData to theHtmlData & "<th title=\"" & theLineData1 & "さんのメールアドレスリンク\" class=\"mailname\">" as text
		set theHtmlData to (theHtmlData & "<a rel=\"nofollow\" href=\"mailto:" & theLinkNameMac & "?body=" & theLinkBodyNameMac & "\" tabindex=\"" & (numTabIndex) as text) & "\" title=\"" & theLineData1 & "さん\">" & theLineData1 & "</a></th>" as text
		set theHtmlData to theHtmlData & "<td title=\"所属\">" & theLineData2 & "</td>" as text
		set theHtmlData to theHtmlData & "<td title=\"メールリンク\">&nbsp;" as text
		set theHtmlData to (theHtmlData & " <a rel=\"nofollow\" href=\"mailto:" & theLinkNameMac & "\" title=\"" & theLineData1 & "さん+メールアドレス。一般的なメールソフトではこちらをクリック\" tabindex=\"" & (numTabIndex + 1) as text) & "\">MailLink</a>" as text
		set theHtmlData to theHtmlData & "&nbsp;|&nbsp;" as text
		set theHtmlData to (theHtmlData & " <a rel=\"nofollow\" href=\"mailto:" & theLinkNameWin & "\" title=\"" & theLineData1 & "さん+メールアドレス。WindowsLiveメールを利用している方はこちらをクリック\" tabindex=\"" & (numTabIndex + 2) as text) & "\">LiveMail</a>" as text
		set theHtmlData to theHtmlData & "&nbsp;</td>" as text
		set theHtmlData to theHtmlData & "<td title=\"メールアドレス\">" as text
		set theHtmlData to (theHtmlData & "<a rel=\"nofollow\" href=\"mailto:" & theOrgEmail & "\" title=\"メールアドレスだけのリンク\" tabindex=\"" & (numTabIndex + 3) as text) & "\" class=\"emailadd\">" & theOrgEmail & "</a></td>" as text
		set theHtmlData to theHtmlData & "</tr>" as text
		set theHtmlData to theHtmlData & "\r" as text
		---カウントアップ
		set numLine to numLine + 1 as number
		set numTabIndex to numTabIndex + 4 as number
		---データ初期化
		set theLineData1 to "" as text
		set theLineData2 to "" as text
		set theLineData3 to "" as text
		--リピートの終わり
	end repeat
	---HTMLのフッター処理
	
	set theHtmlData to theHtmlData & "</tbody>\r"
	set theHtmlData to theHtmlData & "</table>\r</div>\r"
	set theHtmlData to theHtmlData & "<!-- ここまでEmailアドレステーブル -->\r"
	---結果をテキストエディタに表示する
	tell application "TextEdit" to launch
	tell application "TextEdit"
		make new document with properties {text:theHtmlData}
	end tell
	tell application "TextEdit" to activate
	
end open

--------//Mac用リンクエンコードサブルーチン
on doUrlEncode(str)
	set scpt to "php -r 'echo urlencode(\"" & str & "\");'"
	return do shell script scpt
end doUrlEncode


--------//Windows用リンクエンコードサブルーチン
on doSJISencode(str)
	set scpt to ("php -r 'echo urlencode(mb_convert_encoding(\"" & str & "\"" & ",\"SJIS\", \"auto\"));'") as text
	return do shell script scpt
end doSJISencode



--------//文字の置き換えのサブルーチン
to doReplace(theText, orgStr, newStr)
	set oldDelim to AppleScript's text item delimiters
	set AppleScript's text item delimiters to orgStr
	set tmpList to every text item of theText
	set AppleScript's text item delimiters to newStr
	set tmpStr to tmpList as text
	set AppleScript's text item delimiters to oldDelim
	return tmpStr
end doReplace
