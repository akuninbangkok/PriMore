'PriMore V4.22 (2009/02/26)
'Copyright (C) 2008-2010 komikoni All Rights Reserved.
'2013/03/21 クリップボード未使用化(達也)
'2013/07/27 (V4.30) Modify of syncronization between each spool out by akuninbangkok

'****************************************************************
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject("WScript.Shell")

if WScript.Arguments.Unnamed.Count=1 Then
  call Parent_Process
ElseIf WScript.Arguments.Named.Count=1  AND _
       WScript.Arguments.Named.Exists("PSFilePath") Then
	call Child_Process
'ElseIf WScript.Arguments.Named.Count=1  AND _
'       WScript.Arguments.Named.Exists("SetClipBoard") Then
'	call SetClipBoard_Process
Else
	a=msgbox("Arguments Error",64)
	msgbox("named="&WScript.Arguments.Named.Count) & vbcrlf & _
	"unnamed="&WScript.Arguments.Unnamed.Count
	WScript.Quit
End If
'****************************************************************
Sub Parent_Process()

'引数のファイルをリネイム
Set objFilePSFile = objFSO.GetFile(WScript.Arguments.Unnamed.Item(0))
objFilePSFile.name = objFSO.GetTempName()

'実体PSファイルのフォルダ
tempFolder=objFSO.GetSpecialFolder(2).Path & "\"

'PriMoreListパス
ListPath=tempFolder & "PriMoreList.ps"

'PSFile情報読み取り
Set objTextPSFile = objFSO.OpenTextFile(objFilePSFile.Path,ForReading)
Y_Top=""
'  Do Until strInput="%%EndComments" Or objFile2.AtEndOfStream
Do Until objTextPSFile.AtEndOfStream
	strInput = objTextPSFile.ReadLine()
	If left(strInput,8)= "%%Title:" Then
		title = Mid(strInput, 10)
		if left(title,1)="<" and right(title,1)=">" Then
			title = HexDecode(mid(title,2,len(title)-2)) '< > both trim
		elseif left(title,1)="(" AND right(title,1) = ")" Then
			title = OctDecode(mid(title,2,len(title)-2)) '( ) both trim
		end if
'暫定Word,PowerPoint対策START
		If left(title,17)= "Microsoft Word - " Then
			title = Mid(title, 18)
		ElseIf left(Title,23) = "Microsoft PowerPoint - " Then
			title = Mid(title, 24)
		End If

'for replace "Page xxxxx"
'		Set objRE = new RegExp
'		objRE.IgnoreCase = True
'		objRE.pattern = "\s+page\s+\d+"
'		Title = objRE.Replace(Title,"")
'		Set objRE = Nothing

'ファイル名部分だけにし
'ファイル名に使えない「? : , ; * " < > |」を「 _ 」に置換する。
		EscapeFilename = objFSO.GetBaseName(title)
		EscapeFilename = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace( _
				EscapeFilename,	"?","_"), _
						":","_"), _
						",","_"), _
						";","_"), _
						"*","_"), _
						"""","_"), _
						"<","_"), _
						">","_"), _
						"|","_")

		CommentTitle = ""
		PffDocTitle =PDFDocEncoding(title)
		PdfMarkTitle = "[/Title (" & PffDocTitle &  ") /DOCINFO pdfmark"
	ElseIf  left(strInput,6)= "%%For:" Then
		CommentFor = strInput
	ElseIf  left(strInput,8)= "%%Pages:" And Mid(strInput,10,1)<>"(" Then  '(atend)
		Pages = Mid(strInput, 10)
	ElseIf  left(strInput,18)= "%%PageBoundingBox:" And Y_Top="" Then
		Y_Top = split(Mid(strInput, 20))(3)
	End If
Loop
objTextPSFile.close

' PrimoPDFが起動していない上でリストファイルが存在している場合、最大三十秒待つ
For iCount = 0 to 59
	If objShell.AppActivate("PrimoPDF by Nitro PDF Software") Then
		Exit for
	End If
	if objFSO.FileExists(ListPath) = false Then
		Exit for
	End If
	WScript.Sleep 500
Next

If objShell.AppActivate("PrimoPDF by Nitro PDF Software") Then
	'メインファイルが存在する場合 （PrimoPDFが起動している場合）
	'List.psに変更後のtempファイルの情報を追加する
	Set objTextList = objFSO.OpenTextFile(ListPath,ForAppending,True)
	'リストファイルにファイル名をエスケイプ(\⇒\\)して出力
	objTextList.WriteLine "[ /Count 0 /Page Pages /View [/XYZ 0 " & Y_Top & " null] /Title (" & PffDocTitle & ") /OUT pdfmark"
	objTextList.WriteLine "_begin_job_"
	'リストファイルにPSファイルのタイトルをしおりで出力
	objTextList.WriteLine "(" & Replace(objFilePSFile.path,"\","\\") & ")run"
	objTextList.WriteLine "__end__job_"
	objTextList.WriteLine "/Pages Pages "& Pages &" add def"

Else
'メインファイルが存在しない場合（PrimoPDFが起動していない場合）
'メインファイルを作成→リストファイルを作成→リストファイルにPSファイルのタイトルをしおりで出力し、PrimoPDFを起動のため子プロセスを起動

'PriMoreMainパス
MainPath=tempFolder & EscapeFilename & ".ps"
	'メインファイルを作成
	Set objTextMain = objFSO.OpenTextFile(MainPath,ForWriting,True)
	objTextMain.WriteLine "%!PS-Adobe-3.0"
	'メインファイルにPSファイルのタイトルをコメント形式とPDFMARK形式で出力
	objTextMain.WriteLine CommentFor
	objTextMain.WriteLine CommentTitle

	'(メインファイルにしおりを出力)
	'メインファイルにリストファイル名をエスケイプ(\⇒\\)して出力
	objTextMain.WriteLine "(" & Replace(ListPath,"\","\\") & ")run"
	objTextMain.WriteLine PdfMarkTitle
'	objTextMain.WriteLine "userdict /pdfmark systemdict /cleartomark get put"
	objTextMain.Close

	'リストファイルを作成
	Set objTextList = objFSO.OpenTextFile(ListPath,ForWriting,True)
	objTextList.WriteLine "%!PS-Adobe-3.0"
	objTextList.WriteLine "% Written by Helge Blischke, see"
	objTextList.WriteLine "% http://groups.google.com/groups?ic=1&selm=3964A684.49D%40srz-berlin.de"
	objTextList.WriteLine "/_begin_job_"
	objTextList.WriteLine "{"
	objTextList.WriteLine "        /tweak_save save def"
	objTextList.WriteLine "        /tweak_dc countdictstack def"
	objTextList.WriteLine "        /tweak_oc count 1 sub def"
	objTextList.WriteLine "        userdict begin"
	objTextList.WriteLine "}bind def"
	objTextList.WriteLine "/__end__job_"
	objTextList.WriteLine "{"
	objTextList.WriteLine "        count tweak_oc sub{pop}repeat"
	objTextList.WriteLine "        countdictstack tweak_dc sub{end}repeat"
	objTextList.WriteLine "        tweak_save restore"
	objTextList.WriteLine "}bind def"
	'リストファイルにファイル名をエスケイプ(\⇒\\)して出力
	objTextList.WriteLine "[ /Count 0 /Page 1 /View [/XYZ 0 " & Y_Top & " null] /Title (" & PffDocTitle & ") /OUT pdfmark"
	objTextList.WriteLine "_begin_job_"
	'リストファイルにPSファイルのタイトルをしおりで出力
	objTextList.WriteLine "(" & Replace(objFilePSFile.path,"\","\\") & ")run"
	objTextList.WriteLine "__end__job_"
	objTextList.WriteLine "/Pages 1 "& Pages &" add def"
	objTextList.Close

	'子プロセス(自分自身)を'メインファイルを名前つきパラメータで渡して起動(戻りなし)
	execstmt= """" & Wscript.path & "\cscript.exe"" " & """" & WScript.ScriptFullName & """ /PSFilePath:""" & MainPath & """"
	' msgbox execstmt
	Set objExec = objShell.Exec(execstmt)

End If

End Sub
''****************************************************************
'Sub SetClipBoard_Process()
''IEを使えばクリップボード操作は出来るが、セキュリティ上IEの起動を
''禁止している場合が有るのでVBSだけで行う。
''別プロセスでVBSのINPUTBOXにクリップボードに入れたい文字列を表示し
''呼出元のプロセスでSendKey操作によりクリップボードにセット
''クリップボードにセットが行えるのは1024Byteまでで、改行、タブ等は使用できない
''非表示には出来ない為、遠くに表示。
''AA=INPUTBOX("","PriMore_SetClipBoard",WScript.Arguments.Named("SetClipBoard"),-10000,-10000)
'End Sub
''****************************************************************
Sub Child_Process()

	'メインファイルのパス
	PSFilePath=WScript.Arguments.Named("PSFilePath")

	'PrimoPDFのフォルダ
	PrimoFolder=objFSO.GetParentFolderName(WScript.ScriptFullName)& "\"
	'PrimoPDFのパス
	PrimoPath =PrimoFolder & "PrimoPDF.exe"

	'実体PSファイルのフォルダ
	'tempFolder=objFSO.GetParentFolderName(PSFilePath) & "\"
	tempFolder=objFSO.GetSpecialFolder(2).Path & "\"
	'PriMoreListパス
	ListPath=tempFolder & "PriMoreList.ps"

	'メインファイルを引数にPrimoPDFを起動(戻り確認あり)

	execstmt= """" & PrimoPath & """ """ & PSFilePath & """"
	Set objExec = objShell.Exec(execstmt)

	'PrimoPDFが終了するまで待つ
	Do While objExec.Status = 0
	  WScript.Sleep 100
	Loop
	'念のため二秒待つ
	WScript.Sleep 2000

	'リストファイル内で指定されているファイルを削除
	Set objTextList = objFSO.OpenTextFile(ListPath,ForReading)
	for i=1 to 16
		objTextList.SkipLine
	next
	Do Until objTextList.AtEndOfStream
		objTextList.SkipLine
		objTextList.SkipLine
		strInput = objTextList.ReadLine()
		strInput = Replace(MID(strInput,2,len(strInput)-5),"\\","\")
		objFSO.DeleteFile strInput
		objTextList.SkipLine
		objTextList.SkipLine
	Loop
	objTextList.Close
	'メインファイル、リストファイルを削除
	objFSO.DeleteFile PSFilePath
	objFSO.DeleteFile ListPath

End Sub
'****************************************************************
Function OctDecode(Source)
On Error Resume Next
sTmp=""
iCount = 1
lSrcLen=Len(Source)
Do Until iCount > lSrcLen
	If Mid(Source,iCount,1)="\" Then
		If Mid(Source,iCount +1,1)="\" Or _
		Mid(Source,iCount +1,1)="(" Or _
		Mid(Source,iCount +1,1)=")" Then
			sHex=Hex(asc(Mid(Source,iCount +1,1)))
			iCount = iCount + 2
		Else
			sHex=Hex("&O"&Mid(Source,iCount +1,3))
			If Len(sHex) <2 Then
				sHex ="0" & sHex
			End If
			iCount = iCount + 4
		End If
	Else
        	sHex=Hex(Asc(Mid(Source,iCount,1)))
		If len(sHex) <2 Then
			sHex ="0" & sHex
		End If
		iCount = iCount + 1
	End If
	sTmp=sTmp & sHex
Loop
OctDecode = HexDecode(sTmp)
End Function
'****************************************************************
Function HexDecode(Source)
On Error Resume Next
sTmp=""
iCount = 1
lSrcLen=Len(Source)
Do Until iCount > lSrcLen
	sHex = Mid(Source,iCount,2)
	iCount = iCount + 2
	iAsc = CByte("&H" & sHex)
	If (&H00 <= iAsc And iAsc <= &H80) Or _
	(&HA0 <= iAsc And iAsc <= &HDF) Then
	'1Byte char
		sChr=Chr(iAsc)
	ElseIf (&H81 <= iAsc And iAsc <= &H9F) Or _
	(&HE0 <= iAsc And iAsc <= &HFF) Then
	'2byte char
		sHex2 = Mid(Source,iCount,2)
		iCount = iCount + 2
		sChr=Chr(CInt("&H" & sHex & sHex2))
	End If
	sTmp=sTmp & sChr
Loop
HexDecode = sTmp
End Function
'****************************************************************
Function PDFDocEncoding(Source)
Result=""
For i=1 to len(Source)
	Wbyte= hex(AscW(Mid(Source,i,1)))
	If len(Wbyte) <4 Then
		Wbyte =String(4-len(Wbyte),"0") & Wbyte
	End If

	HiByte= "&h" & Mid(Wbyte,1,2)
	LoByte= "&h" & Mid(Wbyte,3,2)

	Res = Res & PDFDocEncodingByte(HiByte) & _
		PDFDocEncodingByte(LoByte)
Next

If Res<>"" Then
	PDFDocEncoding ="\376\377" & Res
Else
	PDFDocEncoding =Res
End If

End Function
'****************************************************************
Function PDFDocEncodingByte(Source)
select case Source
	case &h5c , &h28 , &h29
		PDFDocEncodingByte="\" & ChrW(Source)
	case else
		if Source >= &h20 And _
			Source <= &h7f Then
			PDFDocEncodingByte=ChrW(Source)
		else
			Octstr=Oct(Source)
			PDFDocEncodingByte= "\" & String(3-len(Octstr),"0") & Octstr
		end if
End Select
End Function
'****************************************************************
