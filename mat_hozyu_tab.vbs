'文字列比較・補充
'文字列は10000文字を限度とする。

dim fso
set fso = createObject("Scripting.FileSystemObject")
pass = fso.getParentFolderName(WScript.ScriptFullName)

 'ファイル名設定
strVal = now() 
 if Mid(strVal,13,1) = ":" then
 yy = Mid(strVal,3,2)
 mm = Mid(strVal,6,2)
 dd = Mid(strVal,9,2)
 hh = "0" & Mid(strVal,12,1)
 min = Mid(strVal,14,2)
 ss = Mid(strVal,17,2)
 fmei = yy & mm & dd & "_" & hh & min & ss & ".txt"
 else
 yy = Mid(strVal,3,2)
 mm = Mid(strVal,6,2)
 dd = Mid(strVal,9,2)
 hh = Mid(strVal,12,2)
 min = Mid(strVal,15,2)
 ss = Mid(strVal,18,2)
 fmei = yy & mm & dd & "_" & hh & min & ss & ".txt"
 end if

fmei2 = inputbox("ファイル名は下記で作成します。","hozyu.vbs",("hozyutab_" & fmei))

msttab = inputbox("何列目（何タブ目）の文字列を比較するか指定してください。" & vbcr & "マスタファイル","mat_hozyu.vbs_tab",(1))
trntab = inputbox("何列目（何タブ目）の文字列を比較するか指定してください。" & vbcr & "トランファイル","mat_hozyu.vbs_tab",(1))

mfmei = "mst.txt"
tfmei = "trn.txt"

'マスターファイル読み込み
set objfs = createobject("scripting.filesystemobject")
set objtext = objfs.opentextfile(pass & "\" & mfmei)

'出力ファイル
set d1 = createobject("scripting.filesystemobject")
set e1 = d1.getfolder(pass)
set f1 = e1.createtextfile(fmei2)

'①mstファイルを読み終わるまで繰り返す
Do While objtext.AtEndOfStream <> True
tabcnt = 1
dataline = objtext.readline
mojiretu = ""
mojicnt = 0
mojisu = len(dataline)
tabcnt = 0
flghai = ""
	'②datalineが読み込み終わるか、mojicnt > 10000まで、以下を繰り返す（文字列の最後に来るまで）
	for i = 1 to 10000 
		moji = Mid(dataline,i,1)
		mojicnt = mojicnt + 1
		if mojicnt = mojisu then
			i = 10001
		else
		end if
	
		if msttab = 1 and tabcnt = 0 then
			flghai = "1"
		else
			'flghai = ""
		end if

		if moji = "	" then
			tabcnt = tabcnt + 1
			'msgbox("_" & msttab & "_")
			tabcnthi = tabcnt + 1
			'msgbox("_" & tabcnthi & "_")
			if CInt(msttab) = CInt(tabcnthi) then 
				'msgbox("tab")
				flghai = "1"
			else
				'msgbox("tab2")
				flghai = ""
			end if
		else
			if flghai = "1" then
			mojiretu = mojiretu & moji
			else
			end if
		end if



	'②ここまで
	next	
	'msgbox(mojiretu)	


		'マスターファイル読み込み
		set objfs = createobject("scripting.filesystemobject")
		set objtext2 = objfs.opentextfile(pass & "\" & tfmei)

		'③trnファイルを読み終わるまで繰り返す
		Do While objtext2.AtEndOfStream <> True
		dataline2 = objtext2.readline
		tabcntt = 1
		mojiretut = ""
		mojicntt = 0
		mojisut = len(dataline2)
		tabcntt = 0
		flghait = ""
'		msgbox(mojisut)
'		msgbox(dataline2)

		'④dataline2が読み込み終わるか、mojicntt > 10000まで、以下を繰り返す（文字列の最後に来るまで）
		for i = 1 to 10000 
			mojit = Mid(dataline2,i,1)
			'msgbox(mojit)
			mojicntt = mojicntt + 1
			if CInt(mojicntt) = CInt(mojisut) then
				i = 10001
			else
			end if
	
			if trntab = 1 and tabcntt = 0 then
				flghait = "1"
			else
				'flghait = ""
			end if
	
			'msgbox("_flghait_" & flghait)
			if mojit = "	" then
				tabcntt = tabcntt + 1
				tabcnthit = tabcntt + 1

'			msgbox("_" & trntab & "_")
'			msgbox("_" & tabcnthit & "_")
	
				if CInt(trntab) = CInt(tabcnthit) then 
					'msgbox("tab")
					flghait = "1"
				else
					flghait = ""
				end if
			else
				if flghait = "1" then
				mojiretut = mojiretut & mojit
					'msgbox("__" & mojiretut)
				else
				end if
			end if

		'msgbox(flghait & "_" & mojiretut)

		'④ここまで
		next	
'	msgbox(mojiretut)	



		if mojiretu = mojiretut then
			dataline = dataline & "	" & dataline2
'			mojiretu = mojiretu & "	" & mojiretut
		else
		end if
	

		'③ここまで
		loop


	f1.writeline(dataline)

'①ここまで
loop



