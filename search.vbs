Set WshShell = WScript.CreateObject("WScript.Shell")
dim fs:set fs= CreateObject("Scripting.FileSystemObject")
Const ForReading = 1, ForWriting = 2, ForAppending = 8
' random
Dim rd1, rd2 :rd1 = 1  : rd2 = 37
' 'xoa file
' fs.DeleteFile GetNameFolder(strFileData) & "/find.bat"
MyArrayTagHTML = Array("<a href="""" target=""" & "_blank" & """></a>", "<iframe src =""""></iframe>", "<img src="""" alt=""" & "image" & """>"," them tu do")

Dim strSearch, strFileData, filetxt, strConnect

Do
	'bo qua loi
	On Error Resume Next
	'tim kiem chinh
 	strSearch=InputBox(ThuTrongTuan() & " " & Date() & " " & Time() & vbCrLf & "- --- --" & vbCrLf & "Connect " & CheckFile(strFileData) & vbCrLf & "- --- --" & vbCrLf & "[SEARCH]" & vbCrLf & strSearch,strFileData,GoiY(strFileData))
	'thiet lap
	if strSearch = ".connect" Then
		strFileData = InputBox("LINK > FILE DATA","CONNECT...")
		'tao file bat de bao mat du lieu
		if fs.FileExists(strFileData) Then
			' thuc thi file bat
			RunBatFile GetNameFolder(strFileData) & "/BV.bat", "attrib +s +h +r ", strFileData
		else
			x=msgbox("KHONG TIM THAY DU LIEU!" ,16, "ERROR!")
		End if
		
	'tim kiem thong tin
	Elseif strSearch <> "" And strSearch <> ".add" And strSearch <> ".offline" And fs.FileExists(strFileData) And strSearch <> ".restart" And strSearch <> ".drop" And strSearch <> ".call" And strSearch <> ".source" And strSearch <> ".bat"  Then
		'kiem tra su ton tai cua thu muc
		if fs.FolderExists(GetNameFolder(strFileData)) Then
			'dung 1 giay
			WScript.Sleep(1000)
			
			'ghi du lieu vao 1 file.bat o trong thu muc da kiem tra
			GhiFile GetNameFolder(strFileData) & "/find.bat", ForWriting, "C:" & vbCrLf &  "find """ & strSearch & """ " & """" & strFileData & """ > " & GetNameFolder(strFileData) & "/KQ.txt"
			
			'dung 1 giay
			WScript.Sleep(1000)
			
			'chay file tim kiem du lieu
			WshShell.Run GetNameFolder(strFileData) & "/find.bat"
			
			'dung 1 giay
			WScript.Sleep(1000)
			
			'ghi du lieu styles vao file KQ, html code
			GhiFile GetNameFolder(strFileData) & "/KQ.txt",ForAppending, "<style> body {background-color: #000; color: #fff; animation: colorBackground 100s infinite; background-repeat: no-repeat;background-attachment: fixed; background-size: cover; margin-left: 100px; margin-right: 100px; margin-top: 50px} a{color: aqua;}  @keyframes colorBackground {0% {background-image: url(./BG/" & GetRd(rd1,rd2) & ".gif);}20% {background-image: url(./BG/" & GetRd(rd1,rd2) & ".gif);}40% {background-image: url(./BG/" & GetRd(rd1,rd2) & ".gif);}60% {background-image: url(./BG/" & GetRd(rd1,rd2) & ".gif);}80% {background-image: url(./BG/" & GetRd(rd1,rd2) & ".gif);}100% {background-image: url(./BG/" & GetRd(rd1,rd2) & ".gif);}} </style><title>" & strSearch & "</title>"
			
			'dung 1 giay
			WScript.Sleep(1000)
			
			' tao file html vs phan head
			GhiFile GetNameFolder(strFileData) & "/KQ.html",ForWriting, "<link rel=""icon"" type=""image/x-icon"" href=""./icon.ico"">" & DocFile(GetNameFolder(strFileData) & "/KQ.txt")

			'dung 1 giay
			WScript.Sleep(1000)

			'chay file html
			WshShell.Run GetNameFolder(strFileData) & "/KQ.html"
			
			'dung 2 giay
			WScript.Sleep(2000)

			' hight line noi dung tim kiem bang send key
			WshShell.SendKeys "^{f}"

			'ghi nua giay
			WScript.Sleep 1000
			' ghi du lieu tim kiem de hight line
			WshShell.SendKeys strSearch

			'ghi nua 2 giay
			WScript.Sleep 2000
		End if
	'them du lieu vao file data
	Elseif strSearch = ".add" And fs.FileExists(strFileData) Then
		' chon the html
		strTagHTML = InputBox("[0] " & MyArrayTagHTML(0) & vbCrLf & "[1] " & MyArrayTagHTML(1) & vbCrLf & "[2] " & MyArrayTagHTML(2) & vbCrLf & "[3] " & MyArrayTagHTML(3) & vbCrLf & "[4] them tu file",CheckFile(strFileData),"0 > 4")
		' chon tu 1 den 4
		if strTagHTML > -1 And strTagHTML < 5 Then
			strNewAdd = InputBox(MyArrayTagHTML(strTagHTML),CheckFile(strFileData),MyArrayTagHTML(strTagHTML))
			if strTagHTML = 4 Then
				strNewAdd = InputBox("FILE DU LIEU" & vbCrLf & strFileData,CheckFile(strFileData),"FILE DATA")
				'kiem tra file co ton tai
				if fs.FileExists(strNewAdd) Then
					'thuc thi file bat
					RunBatFile GetNameFolder(strFileData) & "/BV.bat","attrib -s -h -r ",strFileData
					
					' dung 5 giay
					WScript.Sleep(5000)

					'ghi du lieu vao file da thiet lap
					GhiFile strFileData,ForAppending,DocFile(strNewAdd)
					
					'thuc thi file bat
					RunBatFile GetNameFolder(strFileData) & "/BV.bat","attrib +s +h +r ",strFileData
					
					' hien thong bao
					msgbox "THEM THANH CONG."
				else
					x=msgbox("KHONG TIM THAY DU LIEU!" ,16, "ERROR!")
				End if
			'them tu do 
			Elseif strNewAdd <> "" Then
				'thuc thi file bat
				RunBatFile GetNameFolder(strFileData) & "/BV.bat","attrib -s -h -r ",strFileData
				
				' dung 5 giay
				WScript.Sleep(5000)
				
				'ghi du lieu vao file da thiet lap
				GhiFile strFileData,ForAppending,"<p><h3>" & strNewAdd & "</h3></p>"
				
				'thong bao
				msgbox DocFile( GetNameFolder(strFileData) & "/TB.txt")
				
				'thuc thi file bat
				RunBatFile GetNameFolder(strFileData) & "/BV.bat","attrib +s +h +r ",strFileData
			End if
		End if
		
	' mo file de chinh sua
	Elseif strSearch = ".drop" And fs.FileExists(strFileData) Then
		'thuc thi file bat
		RunBatFile GetNameFolder(strFileData) & "/BV.bat","attrib -s -h -r ",strFileData

		'mo file
		WshShell.Run strFileData
	' mo duong dan
	Elseif strSearch = ".call" Then
		strCall = Inputbox("DUONG DAN DUOC TIN TUONG!","LINK..")
		if filesys.FileExists(strCall) Or foldersys.FolderExists(strCall) Or InStr(strCall, "https://") > 0  Then
			WshShell.Run strCall
		Else
			x=msgbox("KHONG TIM THAY DUONG DAN TREN!" & vbCrLf & strCall & "?",16, "THONG BAO.")
		End if
	' mo thu muc goc
	Elseif strSearch = ".source" Then
		'kiem tra thu muc co ton tai
		if fs.FolderExists(fs.GetParentFolderName(strFileData)) Then
			WshShell.Run fs.GetParentFolderName(strFileData)
		Else
			msgbox "I DON'T KNOW!"
		End if
	' thuc thi cau lenh cmd
	Elseif strSearch = ".bat" And fs.FileExists(strFileData) Then
		'ghi lenh
		strBat = InputBox("CMD ,BAT",GetNameFolder(strFileData) & "/file.bat")
		
		if strBat <> "" Then
			'thuc thi file bat
			RunBatFile GetNameFolder(strFileData) & "/file.bat",strBat,""
		End if
		
	End if

Loop Until strSearch = ".offline" Or strSearch = ".restart"

' khoi dong lai he thong
if strSearch = ".restart" Then
	WshShell.Run WScript.ScriptFullName
End if


'lay ten thu muc
Function GetNameFolder(strNameFile)
	if fs.FileExists(strNameFile) Then
		GetNameFolder = fs.GetParentFolderName(strNameFile)
	End if
	
End Function

'Kiem tra trang thai ket noi
Function CheckFile(strNameFile)
	'kiem tra file
	if fs.FileExists(strNameFile) Then
		CheckFile = "[ ONLINE. ]"
	else
		CheckFile = "[ OFFLINE... ]"
	End if
	
End Function

'ham hien thi ngay trong tuan
Function ThuTrongTuan()
	Dim intDay
	intDay = Weekday(Date())
	result = WeekdayName(intDay)
	ThuTrongTuan = result
End Function

'doc du lieu trong file
Function DocFile(strNameFile)
	Dim inFile: Set inFile = fs.OpenTextFile(strNameFile)
	DocFile = inFile.ReadAll
	inFile.Close
End Function

' thuc thi file.bat
Function RunBatFile(strNameFile, strCodeCMD, strLinkRun)
	' dong file
	Set filetxt = fs.OpenTextFile(strNameFile,2, True)
	filetxt.WriteLine(strCodeCMD & strLinkRun)
	filetxt.Close
	' dung 1 giay
	WScript.Sleep(1000)
	' chay file bv
	WshShell.Run strNameFile
End Function

'goi y cho nguoi dung
Function GoiY(strNameFile)
	'kiem tra thu muc co ton tai
	if fs.FileExists(strNameFile) Then
		GoiY = ".add"
	else 
		GoiY = ".connect"
	End if
End Function

'ghi file
Function GhiFile(strNameFile,cstForWhat,strContent)
    Set filetxt = fs.OpenTextFile(strNameFile,cstForWhat, True)
    filetxt.WriteLine(strContent)
    filetxt.Close
    GhiFile = strContent
End Function

'Random
Function GetRd(min, max)
	Randomize
	GetRd = (Int((max-min+1)*Rnd+min))
End Function
