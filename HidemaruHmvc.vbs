'// �o�[�W�����A�b�v���m�点�X�N���v�g
'// �L����ЃT�C�g�[���
'// 2011/09/16 Ver1.00
'// 2011/09/22 Ver1.01
'// 	�x�[�^�t���O��ǉ����āA�x�[�^�ł���̌Ăяo�������ʂ���悤�ɏC��
'// 2011/12/12 Ver1.02
'// 	Twitter�̐ݒ�ɂ��A�N�Z�X���@�̈Ⴂ�ɑΏ�����悤�ɏC��
'// 2012/01/24 Ver1.03
'// 	�A���t�@�x�b�g�����݂���o�[�W������񂪏����ΏۂɂȂ�ƃG���[�ɂȂ�s��̏C��
'// 2012/10/16 Ver1.04
'// 	Twitter�̎d�l�ύX�Ȃ̂��Q�Ƃ��Ă���RSS�t�@�C�����Ȃ��Ȃ��Ă��܂����̂�API�̃^�C�����C�����擾����悤�ɏC��
'// 	�������[�h���Ƀ��b�Z�[�W���̃\�t�g�E�F�A�����G�ۃG�f�B�^�Œ�ɂȂ��Ă����s��̏C��
'// 	Twitter���擾���̃G���[�`�F�b�N������
'// 2012/10/19 Ver1.05
'// 	���ЃT�[�o�ˑ��^�Ɋ��S�ɍ��ς���
'// 	5�Ԗڂ̃p�����[�^�u����t���O�v�͖����ɂ���
'// 2013/03/06 Ver1.06
'// 	�R�~���j�e�b�N�X�T�[�o�ւ̃A�N�Z�X������p�~����
'// 	�G���[�������̃��b�Z�[�W���ȑf��
'// 2013/03/07 Ver1.07
'// 	�`�F�b�N�T�[�o�̏��Ԃ����ւ���
'// 		�u�ߋ����O�T�[�o�v�|���u�R�~���j�e�b�N�v���u�R�~���j�e�b�N�X�v�|���u�ߋ����O�T�[�o�v
'// 	�񐔐������̃��b�Z�[�W�ɃT�C�g�ւ̗U�������������s��̏C��
'//
'// hmvc.vbs "�\�t�g�E�F�A��" "�o�[�W����" "�_�E�����[�hURL" "�x�[�^�t���O"

	Option Explicit

	Dim mobjWShell
	Dim mstrErrorMessage, mstrTitle, mstrMes, mstrS, mstrGetURL
	Dim marrSoftwareTable(20, 3), marrParam, marrItemParam
	Dim mstrRegKey

	Const REG_KEY_LASTCHECKTIME = "\HMVC1"
	Const REG_KEY_CSERVERCHECKTIME = "\HMVC2"

	WScript.Timeout = 0	'// �^�C���A�E�g�𖳌���
	'// �\�t�g�E�F�A���ƃ`�F�b�N�t�@�C�����̐ݒ�
	'// �V�����\�t�g�E�F�A�̏ꍇ�́A�����Ɂu�\�t�g�E�F�A���v�Ɓu�t�@�C�����v�u���W�X�g���L�[�v��ǉ����܂��B

	'// �G�ۃG�f�B�^////////////////////////////////////////////////
	marrSoftwareTable(0, 0) = "�G�ۃG�f�B�^"
	marrSoftwareTable(0, 1) = "hidemaru.ver"
	marrSoftwareTable(0, 2) = "HKCU\Software\Hidemaruo\Hidemaru"
	'///////////////////////////////////////////////////////////////
	'// �G�ۃ��[��//////////////////////////////////////////////////
	marrSoftwareTable(1, 0) = "�G�ۃ��[��"
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
	'// �G�ۃp�u���b�V���[//////////////////////////////////////////
	marrSoftwareTable(4, 0) = "�G�ۃp�u���b�V���["
	marrSoftwareTable(4, 1) = "hmpv.ver"
	marrSoftwareTable(4, 2) = "HKCU\Software\Hidemaruo\Hmpv"
	'///////////////////////////////////////////////////////////////
	'// �G�ۃ��}�C���_//////////////////////////////////////////////
	marrSoftwareTable(5, 0) = "�G�ۃ��}�C���_"
	marrSoftwareTable(5, 1) = "hmrem.ver"
	marrSoftwareTable(5, 2) = "HKCU\Software\Hidemaruo\HmReminder"
	'///////////////////////////////////////////////////////////////
	'// �G�ۃt�@�C���[Classic///////////////////////////////////////
	marrSoftwareTable(6, 0) = "�G�ۃt�@�C���[Classic"
	marrSoftwareTable(6, 1) = "hmfilerclassic.ver"
	marrSoftwareTable(6, 2) = "HKCU\Software\Hidemaruo\HmFilerClassic"
	'///////////////////////////////////////////////////////////////
	'// �G�ۃX�^�[�g���j���[////////////////////////////////////////
	marrSoftwareTable(7, 0) = "�G�ۃX�^�[�g���j���["
	marrSoftwareTable(7, 1) = "hmstartmenu.ver"
	marrSoftwareTable(7, 2) = "HKCU\Software\Hidemaruo\HmStartMenu"
	'///////////////////////////////////////////////////////////////

	On Error Resume Next
	'// �N���p�����[�^�̃`�F�b�N
	If Not CheckParam(marrParam, mstrErrorMessage) Then
		'// �s�������΃G���[�Ƃ���
		Call MsgBox(mstrErrorMessage, vbCritical Or vbOKOnly, "�T�C�g�[���  �o�[�W�����`�F�b�J�[")
		WScript.Quit(-1)
	End If
	'// �_�C�A���O�p�^�C�g���̐ݒ�
	mstrTitle = marrParam(0) & " �o�[�W�����`�F�b�J�["
	'// Shell�I�u�W�F�N�g�̍쐬
	Set mobjWShell = WScript.CreateObject("WScript.Shell")
	If mobjWShell Is Nothing Then
		'// �쐬�ł��Ȃ���΃G���[�Ƃ���
		Call DefaultErrorMessage(False)
		WScript.Quit(-1)
	End If
	'// �g�p���郌�W�X�g���L�[���m�肳���Ă��܂�
	mstrRegKey = GetRegKey(marrParam(0))
	If mstrRegKey = "" Then
		Call DefaultErrorMessage(True)
		WScript.Quit(-1)
	End If
	'// �{��1��ڂ̋N�����ǂ������ׂ�
	If Not IsTodayCheck(marrParam(0)) Then
		'// ������ڂ������烁�b�Z�[�W��\�����ďI���
		mstrMes = "�����͊��ɁA�u�ŐV�o�[�W�����̊m�F�v�R�}���h�����s�ς݂̂悤�ł��B" & vbCrLf
		mstrMes = mstrMes & "�u�ŐV�o�[�W�����̊m�F�v�R�}���h�̂����p�́A�e�\�t�g�E�F�A���ƂɂP���P��܂łł��肢���܂��B" & vbCrLf & vbCrLf
		mstrMes = mstrMes & marrParam(0) & "�̍ŐV�o�[�W�������Љ�y�[�W�Ŋm�F���܂����H"
		If vbYes = MsgBox(mstrMes, vbQuestion Or vbYesNo, mstrTitle) Then Call mobjWShell.Run(marrParam(2))
		WScript.Quit(0)
	End If
	'// �o�[�W�������t�@�C�����擾����URL���擾����
	mstrGetURL = GetCheckURL(marrParam(0), False)
	If mstrGetURL = "" Then
		'// �擾�ł��Ȃ�΃G���[�Ƃ���
		Call DefaultErrorMessage(True)
		WScript.Quit(-1)
	End If
	'// �T�[�o�ɃA�N�Z�X���ăo�[�W�������t�@�C�����擾����
	mstrS = GetHTTPFile(mstrGetURL, marrParam(0))
	'// �擾��Ԃ��`�F�b�N����
	If mstrS = "" Then
		'// �Ȃɂ��擾�ł��Ȃ������ꍇ�̓G���[���������Ă���
		'// �R�~���j�e�b�N�X�T�[�o��URL���擾���Ă݂�
		mstrGetURL = GetCheckURL(marrParam(0), True)
		'// �擾�ł����烊�g���C����
		If Not mstrGetURL = "" Then
			'// �T�[�o�ɃA�N�Z�X���ăo�[�W�������t�@�C�����擾����
			mstrS = GetHTTPFile(mstrGetURL, marrParam(0))
			If mstrS = "" Then
				'// ����ł��ʖڂȂ炵�傤���Ȃ�
				Call DefaultErrorMessage(True)
				WScript.Quit(-1)
			End If
		Else
			Call DefaultErrorMessage(True)
			WScript.Quit(-1)
		End If
	End If
	'// �擾�������e����o�[�W�����`�F�b�N���s��
	If Not CheckVersion(mstrS, marrParam(0), marrParam(1), marrParam(3), marrItemParam) Then
		Call DefaultErrorMessage(True)
		WScript.Quit(-1)
	End If

	'// �擾������񂩂�e�탁�b�Z�[�W��\�����܂��B
	On Error Resume Next
	If Not marrItemParam(0) = "" Then
		rem Set mobjWShell = WScript.CreateObject("WScript.Shell")
		mstrMes = "�V�����o�[�W���������p�\�ł��B" & vbCrLf & vbCrLf
		mstrMes = mstrMes & Mid(marrItemParam(0), 1, 4) & "�N" & Mid(marrItemParam(0), 6, 2) & "��" & Mid(marrItemParam(0), 9) & "���ɁA"
		If marrItemParam(2) Then
			mstrMes = mstrMes & "Ver" & marrItemParam(1) & " �̐����ł����J����Ă��܂��B" & vbCrLf & vbCrLf
		Else
			mstrMes = mstrMes & "Ver" & marrItemParam(1) & " �����J����Ă��܂��B" & vbCrLf & vbCrLf
		End If
		mstrMes = mstrMes & "�_�E�����[�h�y�[�W��\�����܂����H"
		If vbYes = MsgBox(mstrMes, vbQuestion Or vbYesNo, mstrTitle) Then Call mobjWShell.Run(marrParam(2))
	Else
		If CBool(marrParam(3)) Then
			mstrMes = "�����p��" & marrParam(0) & "�́AVer" & marrParam(1) & "�̃x�[�^�łł��B"  & vbCrLf
			mstrMes = mstrMes & "�u�ŐV�o�[�W�����̊m�F�v�Ŋm�F�ł���̂́A�����łƂ��Č��J���ꂽ�o�[�W�����݂̂ŁA�x�[�^�ł̍ŐV�o�[�W�������m�F���邱�Ƃ͏o���܂���B" & vbCrLf & vbCrLf
			mstrMes = mstrMes & marrParam(0) & "�̍ŐV�o�[�W�������Љ�y�[�W�Ŋm�F���܂����H"
			If vbYes = MsgBox(mstrMes, vbQuestion Or vbYesNo, mstrTitle) Then Call mobjWShell.Run(marrParam(2))
			WScript.Quit(-1)
		Else
			If CBool(marrItemParam(3)) Then
				mstrMes = "�����p�̃o�[�W�����́AVer" & marrParam(1) & "�ł��B" & vbCrLf & vbCrLf
				mstrMes = mstrMes & "�T�[�o����擾�����ŐV�̃o�[�W�������́AVer" & marrItemParam(1) & "�ł��B" & vbCrLf
				mstrMes = mstrMes & marrParam(0) & "�̍ŐV�o�[�W�������Љ�y�[�W�Ŋm�F���܂����H"
				If vbYes = MsgBox(mstrMes, vbQuestion Or vbYesNo, mstrTitle) Then Call mobjWShell.Run(marrParam(2))
				WScript.Quit(-1)
			Else
				mstrMes = "�����p��" & marrParam(0) & "�́A�ŐV�o�[�W�����ł��B" & vbCrLf & vbCrLf
				mstrMes = mstrMes & "�����p�̃o�[�W���� = Ver" & marrParam(1) & vbCrLf
				mstrMes = mstrMes & "�T�[�o����擾�����ŐV�̃o�[�W���� = Ver" & marrItemParam(1)
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
		'// �A���t�@�x�b�g�t�̃o�[�W�������ǂ������ׂ�
		'// �A���t�@�x�b�g�����Ă���ꍇ�͕ϊ����ĕԂ��܂��B
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
		'// �N���p�����[�^�̃`�F�b�N���s���܂��B
		Dim intI, bolFlg
		ErrorMessage = ""
		CheckParam = False
		Select Case WScript.Arguments.Count
		Case 2, 3, 4, 5
		Case Else
			ErrorMessage = "�����G���[: �X�N���v�g�ɓn���ꂽ�p�����[�^���Ԉ���Ă܂��B" & vbCrLf
			ErrorMessage = ErrorMessage & "��ʂɂ��̃G���[�̓X�N���v�g�̌Ăяo���������������\��������܂��B" & vbCrLf
			Exit Function
		End Select
		ReDim Param(WScript.Arguments.Count)
		For intI = 0 To WScript.Arguments.Count - 1
			Param(intI) = WScript.Arguments(intI)
		Next
		If Param(0) = "" Then
			ErrorMessage = "�����G���[: �X�N���v�g�ɓn���ꂽ�p�����[�^�i�\�t�g�E�F�A���j���Ԉ���Ă܂��B" & vbCrLf
			ErrorMessage = ErrorMessage & "��ʂɂ��̃G���[�̓X�N���v�g�̌Ăяo���������������\��������܂��B" & vbCrLf
			Exit Function
		End If
		'// �\�t�g�E�F�A���̃`�F�b�N
		For intI = 0 To UBound(marrSoftwareTable)
			If LCase(Param(0)) = LCase(marrSoftwareTable(intI, 0)) Then
				bolFlg = True
				Exit For
			End If
		Next
		If Not bolFlg Then
			ErrorMessage = "�����G���[: �X�N���v�g�ɓn���ꂽ�p�����[�^�i�\�t�g�E�F�A���j���Ԉ���Ă܂��B" & vbCrLf
			ErrorMessage = ErrorMessage & "��ʂɂ��̃G���[�́A�X�N���v�g�̌Ăяo�������p�����[�^�Ɏw�肵���\�t�g�E�F�A���������������A"
			ErrorMessage = ErrorMessage & "�X�N���v�g���̃\�t�g�E�F�A��񂪂��������\��������܂��B" & vbCrLf
			Exit Function
		End If
		If 0 = InStr(1, Param(1), ".", 1) Then
			Select Case Len(Param(1))
			Case 2	: Param(1) = Mid(Param(1), 1, 1) & "." & Mid(Param(1), 2, 1)
			Case 3	: Param(1) = Mid(Param(1), 1, 1) & "." & Mid(Param(1), 2)
			Case 4	: Param(1) = Mid(Param(1), 1, 2) & "." & Mid(Param(1), 3)
			End Select
		End If
		'// �w�肳�ꂽ�o�[�W�����ԍ��̃`�F�b�N
		If Not CheckParamVersion(Param(1)) Then
			'// �w�肳�ꂽ�o�[�W�����ԍ����z�肳�ꂽ���̂ł͂Ȃ�
			ErrorMessage = "�����G���[: �X�N���v�g�ɓn���ꂽ�p�����[�^�i�o�[�W�������j���Ԉ���Ă܂��B" & vbCrLf
			ErrorMessage = ErrorMessage & "��ʂɂ��̃G���[�̓X�N���v�g�̌Ăяo���������������\��������܂��B" & vbCrLf
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
		'// �w��̃o�[�W�����������؂��܂��B
		On Error Resume Next
		Dim objRE	: Set objRE = CreateObject("VBScript.RegExp")
		'// �o�[�W�����ԍ��̃`�F�b�N
		Select Case Instr(1, Version, ".", 1)
		Case 2		'// 1.00�`9.00
			objRE.Pattern = "^[0-9].[0-9][0-9]$"
			If Not objRE.Test(Version) Then
				objRE.Pattern = "^[0-9].[0-9][0-9][a-z]$"
				If Not objRE.Test(Version) Then
					Set objRE = Nothing
					CheckParamVersion = False
					Exit Function
				End If
			End If
		Case 3		'// 10.00�`99.00
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
		'// �w��URL�̃t�@�C�����擾���܂��B
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
			'// �ŐV�̃o�[�W������񂪐���Ɏ擾�ł��Ă��邩���ׂ�
			'// ����������Ύ擾���ɉ��炩�̃G���[����������
			If strRes = "" Or Not CheckSouce(strRes) Then GetHTTPFile = "" : Exit Function
			GetHTTPFile = strRes
		End With
		Set objMSXML = Nothing
		'// ����ɃA�N�Z�X�ł�����������L�^����
		Call mobjWShell.RegWrite(mstrRegKey & REG_KEY_LASTCHECKTIME, Now, "REG_SZ")
		On Error GoTo 0
	End Function

	Function GetCheckURL(ByVal SoftwareTitle, ByVal Flg)
		'// �x�[�W�������t�@�C�����擾����URL���m�肵�܂��B
		Dim strFileName, strRegKey, strLastCheckTime
		Dim intI
		'// �\�t�g�E�F�A������擾����t�@�C�������m�肵�܂��B
		For intI = 0 To UBound(marrSoftwareTable) - 1
			If LCase(SoftwareTitle) = LCase(marrSoftwareTable(intI, 0)) Then
				strFileName = marrSoftwareTable(intI, 1)
				Exit For
			End If
		Next
		'// �t�@�C�������擾�ł��Ȃ���΃G���[�Ƃ��邵���Ȃ�
		If strFileName = "" Then GetCheckURL = "" : Exit Function
		On Error Resume Next
		If Flg Then
			'// �ߋ����O�T�[�o
			GetCheckURL = "http://maruo.dyndns.org:81/software/" & strFileName
		Else
			'// �R�~���j�e�b�N�X�T�[�o
			GetCheckURL = "http://www2.maruo.co.jp/_software/" & strFileName
		End If
		On Error Goto 0
	End Function

	Function CheckSouce(Souce)
		'// �擾���������񂪐��킩�ǂ������ׂ܂��B
		On Error Resume Next
		Dim objRE
		Set objRE = CreateObject("VBScript.RegExp")
		objRE.Pattern = "^20[0-9][0-9]/[0-1][0-9]/[0-3][0-9] [0-9].*$"
		If objRE.Test(Souce) Then CheckSouce = True Else CheckSouce = False
		Set objRE = Nothing
		On Error Goto 0
	End Function

	Function IsTodayCheck(SoftwareTitle)
		'// �����ɕ�����`�F�b�N���Ă��Ȃ������ׂ�
		'// ���Ă����False�A���Ă��Ȃ����True��Ԃ�
		'// ���Ԃ͌��Ă��Ȃ��̂ŁA���t���ς���OK�Ƃ���
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
		'// �\�t�g�E�F�A������g�p���郌�W�X�g���L�[���擾����
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
		'// �W���̃G���[���b�Z�[�W����
		Dim strMes
		If CBool(marrParam(3)) Then
			strMes = "�����p��" & marrParam(0) & "�́AVer" & marrParam(1) & "�̃x�[�^�łł��B"  & vbCrLf
		Else
			strMes = "�����p��" & marrParam(0) & "�́AVer" & marrParam(1) & "�ł��B"  & vbCrLf
		End If
		strMes = strMes & "�ŐV�̃o�[�W���������m�F���邱�Ƃ��o���܂���B" & vbCrLf & vbCrLf
		If Flg Then
			strMes = strMes & marrParam(0) & "�̍ŐV�o�[�W�������Љ�y�[�W�Ŋm�F���܂����H"
			If vbYes = MsgBox(strMes, vbQuestion Or vbYesNo, mstrTitle) Then Call mobjWShell.Run(marrParam(2))
		Else
			strMes = strMes & marrParam(0) & "�̏Љ�y�[�W�ł��m�F���������B"
			Call MsgBox(strMes, vbInformation Or vbOkOnly, mstrTitle)
		End If
		WScript.Quit(-1)
	End Sub
