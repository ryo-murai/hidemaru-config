'// バージョンアップお知らせスクリプト
'// 有限会社サイトー企画
'// 2011/09/16 Ver1.00
'// 2011/09/22 Ver1.01
'// 	ベータフラグを追加して、ベータ版からの呼び出しを識別するように修正
'// 2011/12/12 Ver1.02
'// 	Twitterの設定によるアクセス方法の違いに対処するように修正
'// 2012/01/24 Ver1.03
'// 	アルファベットが混在するバージョン情報が処理対象になるとエラーになる不具合の修正
'// 2012/10/16 Ver1.04
'// 	Twitterの仕様変更なのか参照していたRSSファイルがなくなってしまったのでAPIのタイムラインを取得するように修正
'// 	自動モード時にメッセージ内のソフトウェア名が秀丸エディタ固定になっていた不具合の修正
'// 	Twitter情報取得時のエラーチェックを強化
'// 2012/10/19 Ver1.05
'// 	自社サーバ依存型に完全に作り変える
'// 	5番目のパラメータ「動作フラグ」は無効にした
'// 2013/03/06 Ver1.06
'// 	コミュニテックスサーバへのアクセス制限を廃止した
'// 	エラー発生時のメッセージを簡素化
'// 2013/03/07 Ver1.07
'// 	チェックサーバの順番を入れ替えた
'// 		「過去ログサーバ」－＞「コミュニテック」を「コミュニテックス」－＞「過去ログサーバ」
'// 	回数制限時のメッセージにサイトへの誘導が無かった不具合の修正
'//
'// hmvc.vbs "ソフトウェア名" "バージョン" "ダウンロードURL" "ベータフラグ"

	Option Explicit

	Dim mobjWShell
	Dim mstrErrorMessage, mstrTitle, mstrMes, mstrS, mstrGetURL
	Dim marrSoftwareTable(20, 3), marrParam, marrItemParam
	Dim mstrRegKey

	Const REG_KEY_LASTCHECKTIME = "\HMVC1"
	Const REG_KEY_CSERVERCHECKTIME = "\HMVC2"

	WScript.Timeout = 0	'// タイムアウトを無効化
	'// ソフトウェア名とチェックファイル名の設定
	'// 新しいソフトウェアの場合は、ここに「ソフトウェア名」と「ファイル名」「レジストリキー」を追加します。

	'// 秀丸エディタ////////////////////////////////////////////////
	marrSoftwareTable(0, 0) = "秀丸エディタ"
	marrSoftwareTable(0, 1) = "hidemaru.ver"
	marrSoftwareTable(0, 2) = "HKCU\Software\Hidemaruo\Hidemaru"
	'///////////////////////////////////////////////////////////////
	'// 秀丸メール//////////////////////////////////////////////////
	marrSoftwareTable(1, 0) = "秀丸メール"
	marrSoftwareTable(1, 1) = "tk.ver"
	marrSoftwareTable(1, 2) = "HKCU\Software\Hidemaruo\TuruKame"
	'///////////////////////////////////////////////////////////////
	'// Hidemarnet Explorer/////////////////////////////////////////
	marrSoftwareTable(2, 0) = "Hidemarnet Explorer"
	marrSoftwareTable(2, 1) = "hmnetex.ver"
	marrSoftwareTable(2, 2) = "HKCU\Software\Hidemaruo\hmnetex"
	'///////////////////////////////////////////////////////////////
	'// Hidemarnet Explorer with FTPS///////////////////////////////
	marrSoftwareTable(3, 0) = "Hidemarnet Explorer with FTPS"
	marrSoftwareTable(3, 1) = "hmnetexnet.ver"
	marrSoftwareTable(3, 2) = "HKCU\Software\Hidemaruo\hmnetex"
	'///////////////////////////////////////////////////////////////
	'// 秀丸パブリッシャー//////////////////////////////////////////
	marrSoftwareTable(4, 0) = "秀丸パブリッシャー"
	marrSoftwareTable(4, 1) = "hmpv.ver"
	marrSoftwareTable(4, 2) = "HKCU\Software\Hidemaruo\Hmpv"
	'///////////////////////////////////////////////////////////////
	'// 秀丸リマインダ//////////////////////////////////////////////
	marrSoftwareTable(5, 0) = "秀丸リマインダ"
	marrSoftwareTable(5, 1) = "hmrem.ver"
	marrSoftwareTable(5, 2) = "HKCU\Software\Hidemaruo\HmReminder"
	'///////////////////////////////////////////////////////////////
	'// 秀丸ファイラーClassic///////////////////////////////////////
	marrSoftwareTable(6, 0) = "秀丸ファイラーClassic"
	marrSoftwareTable(6, 1) = "hmfilerclassic.ver"
	marrSoftwareTable(6, 2) = "HKCU\Software\Hidemaruo\HmFilerClassic"
	'///////////////////////////////////////////////////////////////
	'// 秀丸スタートメニュー////////////////////////////////////////
	marrSoftwareTable(7, 0) = "秀丸スタートメニュー"
	marrSoftwareTable(7, 1) = "hmstartmenu.ver"
	marrSoftwareTable(7, 2) = "HKCU\Software\Hidemaruo\HmStartMenu"
	'///////////////////////////////////////////////////////////////

	On Error Resume Next
	'// 起動パラメータのチェック
	If Not CheckParam(marrParam, mstrErrorMessage) Then
		'// 不具合があればエラーとする
		Call MsgBox(mstrErrorMessage, vbCritical Or vbOKOnly, "サイトー企画  バージョンチェッカー")
		WScript.Quit(-1)
	End If
	'// ダイアログ用タイトルの設定
	mstrTitle = marrParam(0) & " バージョンチェッカー"
	'// Shellオブジェクトの作成
	Set mobjWShell = WScript.CreateObject("WScript.Shell")
	If mobjWShell Is Nothing Then
		'// 作成できなければエラーとする
		Call DefaultErrorMessage(False)
		WScript.Quit(-1)
	End If
	'// 使用するレジストリキーを確定させてしまう
	mstrRegKey = GetRegKey(marrParam(0))
	If mstrRegKey = "" Then
		Call DefaultErrorMessage(True)
		WScript.Quit(-1)
	End If
	'// 本日1回目の起動かどうか調べる
	If Not IsTodayCheck(marrParam(0)) Then
		'// 複数回目だったらメッセージを表示して終わる
		mstrMes = "今日は既に、「最新バージョンの確認」コマンドを実行済みのようです。" & vbCrLf
		mstrMes = mstrMes & "「最新バージョンの確認」コマンドのご利用は、各ソフトウェアごとに１日１回まででお願いします。" & vbCrLf & vbCrLf
		mstrMes = mstrMes & marrParam(0) & "の最新バージョンを紹介ページで確認しますか？"
		If vbYes = MsgBox(mstrMes, vbQuestion Or vbYesNo, mstrTitle) Then Call mobjWShell.Run(marrParam(2))
		WScript.Quit(0)
	End If
	'// バージョン情報ファイルを取得するURLを取得する
	mstrGetURL = GetCheckURL(marrParam(0), False)
	If mstrGetURL = "" Then
		'// 取得できなればエラーとする
		Call DefaultErrorMessage(True)
		WScript.Quit(-1)
	End If
	'// サーバにアクセスしてバージョン情報ファイルを取得する
	mstrS = GetHTTPFile(mstrGetURL, marrParam(0))
	'// 取得状態をチェックする
	If mstrS = "" Then
		'// なにも取得できなかった場合はエラーが発生している
		'// コミュニテックスサーバのURLを取得してみる
		mstrGetURL = GetCheckURL(marrParam(0), True)
		'// 取得できたらリトライする
		If Not mstrGetURL = "" Then
			'// サーバにアクセスしてバージョン情報ファイルを取得する
			mstrS = GetHTTPFile(mstrGetURL, marrParam(0))
			If mstrS = "" Then
				'// それでも駄目ならしょうがない
				Call DefaultErrorMessage(True)
				WScript.Quit(-1)
			End If
		Else
			Call DefaultErrorMessage(True)
			WScript.Quit(-1)
		End If
	End If
	'// 取得した内容からバージョンチェックを行う
	If Not CheckVersion(mstrS, marrParam(0), marrParam(1), marrParam(3), marrItemParam) Then
		Call DefaultErrorMessage(True)
		WScript.Quit(-1)
	End If

	'// 取得した情報から各種メッセージを表示します。
	On Error Resume Next
	If Not marrItemParam(0) = "" Then
		rem Set mobjWShell = WScript.CreateObject("WScript.Shell")
		mstrMes = "新しいバージョンが利用可能です。" & vbCrLf & vbCrLf
		mstrMes = mstrMes & Mid(marrItemParam(0), 1, 4) & "年" & Mid(marrItemParam(0), 6, 2) & "月" & Mid(marrItemParam(0), 9) & "日に、"
		If marrItemParam(2) Then
			mstrMes = mstrMes & "Ver" & marrItemParam(1) & " の正式版が公開されています。" & vbCrLf & vbCrLf
		Else
			mstrMes = mstrMes & "Ver" & marrItemParam(1) & " が公開されています。" & vbCrLf & vbCrLf
		End If
		mstrMes = mstrMes & "ダウンロードページを表示しますか？"
		If vbYes = MsgBox(mstrMes, vbQuestion Or vbYesNo, mstrTitle) Then Call mobjWShell.Run(marrParam(2))
	Else
		If CBool(marrParam(3)) Then
			mstrMes = "ご利用の" & marrParam(0) & "は、Ver" & marrParam(1) & "のベータ版です。"  & vbCrLf
			mstrMes = mstrMes & "「最新バージョンの確認」で確認できるのは、正式版として公開されたバージョンのみで、ベータ版の最新バージョンを確認することは出来ません。" & vbCrLf & vbCrLf
			mstrMes = mstrMes & marrParam(0) & "の最新バージョンを紹介ページで確認しますか？"
			If vbYes = MsgBox(mstrMes, vbQuestion Or vbYesNo, mstrTitle) Then Call mobjWShell.Run(marrParam(2))
			WScript.Quit(-1)
		Else
			If CBool(marrItemParam(3)) Then
				mstrMes = "ご利用のバージョンは、Ver" & marrParam(1) & "です。" & vbCrLf & vbCrLf
				mstrMes = mstrMes & "サーバから取得した最新のバージョン情報は、Ver" & marrItemParam(1) & "です。" & vbCrLf
				mstrMes = mstrMes & marrParam(0) & "の最新バージョンを紹介ページで確認しますか？"
				If vbYes = MsgBox(mstrMes, vbQuestion Or vbYesNo, mstrTitle) Then Call mobjWShell.Run(marrParam(2))
				WScript.Quit(-1)
			Else
				mstrMes = "ご利用の" & marrParam(0) & "は、最新バージョンです。" & vbCrLf & vbCrLf
				mstrMes = mstrMes & "ご利用のバージョン = Ver" & marrParam(1) & vbCrLf
				mstrMes = mstrMes & "サーバから取得した最新のバージョン = Ver" & marrItemParam(1)
				Call MsgBox(mstrMes, vbInformation Or vbOkOnly, mstrTitle)
			End If
		End If
	End If
	On Error GoTo 0
	WScript.Quit(0)

	Function CheckVersion(ByVal Souce, ByVal SoftwareTitle, ByVal NowVersion, ByVal BetaFlg, ByRef NewItemParam)
		Dim strTitle, strDate, strVer, arrItem, strNowVersion, strNewVersion
		ReDim NewItemParam(4)
		arrItem = Split(Souce)
		strDate = arrItem(0)
		strVer = arrItem(1)
		strNewVersion = ConvVersion(strVer)
		strNowVersion = ConvVersion(NowVersion)
		If CBool(BetaFlg) Then
			If CSng(strNowVersion * 100) = CSng(strNewVersion * 100) Then
				NewItemParam(0) = strDate
				NewItemParam(1) = strVer
				NewItemParam(2) = True
				NewItemParam(3) = False
				CheckVersion = True
				Exit Function
			ElseIf CSng(strNowVersion * 100) < CSng(strNewVersion * 100) Then
				NewItemParam(0) = strDate
				NewItemParam(1) = strVer
				NewItemParam(2) = False
				NewItemParam(3) = False
				CheckVersion = True
				Exit Function
			Else
				NewItemParam(0) = ""
				NewItemParam(1) = strVer
				NewItemParam(2) = False
				NewItemParam(3) = True
				CheckVersion = True
				Exit Function
			End If
		Else
			If CSng(strNowVersion * 100) = CSng(strNewVersion * 100) Then
				NewItemParam(0) = ""
				NewItemParam(1) = strVer
				NewItemParam(2) = False
				NewItemParam(3) = False
				CheckVersion = True
				Exit Function
			ElseIf CSng(strNowVersion * 100) < CSng(strNewVersion * 100) Then
				NewItemParam(0) = strDate
				NewItemParam(1) = strVer
				NewItemParam(2) = False
				NewItemParam(3) = False
				CheckVersion = True
				Exit Function
			Else
				NewItemParam(0) = ""
				NewItemParam(1) = strVer
				NewItemParam(2) = False
				NewItemParam(3) = True
				CheckVersion = True
				Exit Function
			End If
		End If
		CheckVersion = False
	End Function

	Function ConvVersion(ByVal Version)
		'// アルファベット付のバージョンかどうか調べて
		'// アルファベットがついている場合は変換して返します。
		Dim intI, intX
		Dim strS, strTmp
		strTmp = ""
		For intI = 1 To Len(Version)
			strS = Mid(Version, intI, 1)
			intX = Asc(UCase(strS))
			If intX >= 65 And intX <= 73 Then
				strTmp = strTmp & "0" & CStr(intX - 64)
			ElseIf intX >= 74 And intX <= 90 Then
				strTmp = strTmp & CStr(intX - 64)
			Else
				strTmp = strTmp & strS
			End If
		Next
		ConVVersion = strTmp
	End Function

	Function CheckParam(ByRef Param, ByRef ErrorMessage)
		'// 起動パラメータのチェックを行います。
		Dim intI, bolFlg
		ErrorMessage = ""
		CheckParam = False
		Select Case WScript.Arguments.Count
		Case 2, 3, 4, 5
		Case Else
			ErrorMessage = "内部エラー: スクリプトに渡されたパラメータが間違ってます。" & vbCrLf
			ErrorMessage = ErrorMessage & "一般にこのエラーはスクリプトの呼び出し元がおかしい可能性があります。" & vbCrLf
			Exit Function
		End Select
		ReDim Param(WScript.Arguments.Count)
		For intI = 0 To WScript.Arguments.Count - 1
			Param(intI) = WScript.Arguments(intI)
		Next
		If Param(0) = "" Then
			ErrorMessage = "内部エラー: スクリプトに渡されたパラメータ（ソフトウェア名）が間違ってます。" & vbCrLf
			ErrorMessage = ErrorMessage & "一般にこのエラーはスクリプトの呼び出し元がおかしい可能性があります。" & vbCrLf
			Exit Function
		End If
		'// ソフトウェア名のチェック
		For intI = 0 To UBound(marrSoftwareTable)
			If LCase(Param(0)) = LCase(marrSoftwareTable(intI, 0)) Then
				bolFlg = True
				Exit For
			End If
		Next
		If Not bolFlg Then
			ErrorMessage = "内部エラー: スクリプトに渡されたパラメータ（ソフトウェア名）が間違ってます。" & vbCrLf
			ErrorMessage = ErrorMessage & "一般にこのエラーは、スクリプトの呼び出し元がパラメータに指定したソフトウェア名がおかしいか、"
			ErrorMessage = ErrorMessage & "スクリプト内のソフトウェア情報がおかしい可能性があります。" & vbCrLf
			Exit Function
		End If
		If 0 = InStr(1, Param(1), ".", 1) Then
			Select Case Len(Param(1))
			Case 2	: Param(1) = Mid(Param(1), 1, 1) & "." & Mid(Param(1), 2, 1)
			Case 3	: Param(1) = Mid(Param(1), 1, 1) & "." & Mid(Param(1), 2)
			Case 4	: Param(1) = Mid(Param(1), 1, 2) & "." & Mid(Param(1), 3)
			End Select
		End If
		'// 指定されたバージョン番号のチェック
		If Not CheckParamVersion(Param(1)) Then
			'// 指定されたバージョン番号が想定されたものではない
			ErrorMessage = "内部エラー: スクリプトに渡されたパラメータ（バージョン情報）が間違ってます。" & vbCrLf
			ErrorMessage = ErrorMessage & "一般にこのエラーはスクリプトの呼び出し元がおかしい可能性があります。" & vbCrLf
			Exit Function
		End If
		Select Case UBound(Param)
		Case 2
			ReDim Preserve Param(5)
			Param(2) = "http://hide.maruo.co.jp/"
			Param(3) = False
			Param(4) = False
		Case 3
			ReDim Preserve Param(5)
			Param(3) = False
			Param(4) = False
		Case 4
			ReDim Preserve Param(5)
			Param(4) = False
		End Select
		If Param(2) = "" Then Param(2) = "http://hide.maruo.co.jp/"
		CheckParam = True
	End Function

	Function CheckParamVersion(Byval Version)
		'// 指定のバージョン情報を検証します。
		On Error Resume Next
		Dim objRE	: Set objRE = CreateObject("VBScript.RegExp")
		'// バージョン番号のチェック
		Select Case Instr(1, Version, ".", 1)
		Case 2		'// 1.00～9.00
			objRE.Pattern = "^[0-9].[0-9][0-9]$"
			If Not objRE.Test(Version) Then
				objRE.Pattern = "^[0-9].[0-9][0-9][a-z]$"
				If Not objRE.Test(Version) Then
					Set objRE = Nothing
					CheckParamVersion = False
					Exit Function
				End If
			End If
		Case 3		'// 10.00～99.00
			objRE.Pattern = "^[0-9][0-9].[0-9][0-9]$"
			If Not objRE.Test(Version) Then
				objRE.Pattern = "^[0-9][0-9].[0-9][0-9][a-z]$"
				If Not objRE.Test(Version) Then
					Set objRE = Nothing
					CheckParamVersion = False
					Exit Function
				End If
			End If
		Case Else
			Set objRE = Nothing
			CheckParamVersion = False
			Exit Function
		End Select
		Set objRE = Nothing
		CheckParamVersion = True
		On Error Goto 0
	End Function

	Function GetHTTPFile(ByVal URL, ByVal SoftwareTitle)
		'// 指定URLのファイルを取得します。
		Dim objMSXML, strRes
		On Error Resume Next
		Set objMSXML = CreateObject("Msxml2.XMLHTTP")
		If objMSXML Is Nothing Then GetHTTPFile = "" : Exit Function
		With objMSXML
			Call .Open("GET", URL, False)
			.setRequestHeader "Pragma", "no-cache"
			.setRequestHeader "Cache-Control", "no-cache"
			.setRequestHeader "If-Modified-Since", "Thu, 01 Jun 1970 00:00:00 GMT"
			Call .Send
			If Err.Number <> 0 Or Not .Status = 200 Then GetHTTPFile = "" : Exit Function
			strRes = .ResponseText
			'// 最新のバージョン情報が正常に取得できているか調べる
			'// 何も無ければ取得時に何らかのエラーが発生した
			If strRes = "" Or Not CheckSouce(strRes) Then GetHTTPFile = "" : Exit Function
			GetHTTPFile = strRes
		End With
		Set objMSXML = Nothing
		'// 正常にアクセスできたら日時を記録する
		Call mobjWShell.RegWrite(mstrRegKey & REG_KEY_LASTCHECKTIME, Now, "REG_SZ")
		On Error GoTo 0
	End Function

	Function GetCheckURL(ByVal SoftwareTitle, ByVal Flg)
		'// ベージョン情報ファイルを取得するURLを確定します。
		Dim strFileName, strRegKey, strLastCheckTime
		Dim intI
		'// ソフトウェア名から取得するファイル名を確定します。
		For intI = 0 To UBound(marrSoftwareTable) - 1
			If LCase(SoftwareTitle) = LCase(marrSoftwareTable(intI, 0)) Then
				strFileName = marrSoftwareTable(intI, 1)
				Exit For
			End If
		Next
		'// ファイル名が取得できなければエラーとするしかない
		If strFileName = "" Then GetCheckURL = "" : Exit Function
		On Error Resume Next
		If Flg Then
			'// 過去ログサーバ
			GetCheckURL = "http://maruo.dyndns.org:81/software/" & strFileName
		Else
			'// コミュニテックスサーバ
			GetCheckURL = "http://www2.maruo.co.jp/_software/" & strFileName
		End If
		On Error Goto 0
	End Function

	Function CheckSouce(Souce)
		'// 取得した文字列が正常かどうか調べます。
		On Error Resume Next
		Dim objRE
		Set objRE = CreateObject("VBScript.RegExp")
		objRE.Pattern = "^20[0-9][0-9]/[0-1][0-9]/[0-3][0-9] [0-9].*$"
		If objRE.Test(Souce) Then CheckSouce = True Else CheckSouce = False
		Set objRE = Nothing
		On Error Goto 0
	End Function

	Function IsTodayCheck(SoftwareTitle)
		'// 同日に複数回チェックしていないか調べる
		'// していればFalse、していなければTrueを返す
		'// 時間は見ていないので、日付が変わればOKとする
		Dim strRegKey, strLastCheckTime
		On Error Resume Next
		strRegKey = mstrRegKey & REG_KEY_LASTCHECKTIME
		strLastCheckTime = mobjWShell.RegRead(strRegKey)
		If Not Err.Number = 0 Then
			strLastCheckTime = ""
			Err.Clear
		End If
		If Not strLastCheckTime = "" Then
			If 0 = DateDiff("d", strLastCheckTime, Now) Then
				IsTodayCheck = False
				Exit Function
			End If
		End If
		IsTodayCheck = True
		On Error Goto 0
	End Function

	Function GetRegKey(ByVal SoftwareTitle)
		'// ソフトウェア名から使用するレジストリキーを取得する
		Dim intI
		For intI = 0 To UBound(marrSoftwareTable) - 1
			If LCase(SoftwareTitle) = LCase(marrSoftwareTable(intI, 0)) Then
				GetRegKey = marrSoftwareTable(intI, 2)
				Exit Function
			End If
		Next
		GetRegKey = ""
	End Function

	Sub DefaultErrorMessage(ByVal Flg)
		'// 標準のエラーメッセージ処理
		Dim strMes
		If CBool(marrParam(3)) Then
			strMes = "ご利用の" & marrParam(0) & "は、Ver" & marrParam(1) & "のベータ版です。"  & vbCrLf
		Else
			strMes = "ご利用の" & marrParam(0) & "は、Ver" & marrParam(1) & "です。"  & vbCrLf
		End If
		strMes = strMes & "最新のバージョン情報を確認することが出来ません。" & vbCrLf & vbCrLf
		If Flg Then
			strMes = strMes & marrParam(0) & "の最新バージョンを紹介ページで確認しますか？"
			If vbYes = MsgBox(strMes, vbQuestion Or vbYesNo, mstrTitle) Then Call mobjWShell.Run(marrParam(2))
		Else
			strMes = strMes & marrParam(0) & "の紹介ページでご確認ください。"
			Call MsgBox(strMes, vbInformation Or vbOkOnly, mstrTitle)
		End If
		WScript.Quit(-1)
	End Sub
