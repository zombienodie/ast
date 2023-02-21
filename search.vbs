Set WshShell = WScript.CreateObject("WScript.Shell")
dim fs:set fs= CreateObject("Scripting.FileSystemObject")
' 'xoa file
' fs.DeleteFile GetNameFolder(strFileData) & "/find.bat"

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
			'ghi du lieu vao 1 file.bat o trong thu muc da kiem tra
			Set filetxt = fs.OpenTextFile(GetNameFolder(strFileData) & "/find.bat",2, True)
			filetxt.WriteLine("C:" & vbCrLf &  "find """ & strSearch & """ " & """" & strFileData & """ > " & GetNameFolder(strFileData) & "/KQ.html")
			filetxt.Close
			'dung 1 giay
			WScript.Sleep(1000)
			'chay file tim kiem du lieu
			WshShell.Run GetNameFolder(strFileData) & "/find.bat"
			'dung 1 giay
			WScript.Sleep(1000)
			'ghi du lieu styles vao file KQ
			Set filetxt = fs.OpenTextFile(GetNameFolder(strFileData) & "/KQ.html",8, True)
			filetxt.WriteLine("<style> body {background-color: #000; color: #fff}a{color: aqua;} </style><title>" & strSearch & "</title>")
			filetxt.Close
			'dung 1 giay
			WScript.Sleep(1000)
			'chay file html
			WshShell.Run GetNameFolder(strFileData) & "/KQ.html"
		End if
	'them du lieu vao file data
	Elseif strSearch = ".add" And fs.FileExists(strFileData) Then
		strNewAdd = InputBox("1 them tu file" & vbCrLf & strFileData,CheckFile(strFileData),"<a href="""" target=""" & "_blank" & """></a>")
		' gui thong bao neu du lieu la rong
		if strNewAdd = "" Then
			x=msgbox("KHONG TIM THAY NOI DUNG!" ,16, "NULL!")
		Elseif strNewAdd = "1" Then
			strNewAdd = InputBox("FILE DU LIEU" & vbCrLf & strFileData,CheckFile(strFileData),"FILE DATA")
			'kiem tra file co ton tai
			if fs.FileExists(strNewAdd) Then
				'thuc thi file bat
				RunBatFile GetNameFolder(strFileData) & "/BV.bat","attrib -s -h -r ",strFileData
				
				' dung 5 giay
				WScript.Sleep(5000)

				'ghi du lieu vao file da thiet lap
				Set filetxt = fs.OpenTextFile(strFileData,8,True)
				filetxt.WriteLine(DocFile(strNewAdd))
				filetxt.Close
				
				'thuc thi file bat
				RunBatFile GetNameFolder(strFileData) & "/BV.bat","attrib +s +h +r ",strFileData
				
				' hien thong bao
				msgbox "THEM THANH CONG."
			else
				x=msgbox("KHONG TIM THAY DU LIEU!" ,16, "ERROR!")
			End if
		else
			'thuc thi file bat
			RunBatFile GetNameFolder(strFileData) & "/BV.bat","attrib -s -h -r ",strFileData
			
			' dung 5 giay
			WScript.Sleep(5000)
			
			'ghi du lieu vao file da thiet lap
			Set filetxt = fs.OpenTextFile(strFileData,8,True)
			filetxt.WriteLine("<p>" & strNewAdd & "</p>")
			filetxt.Close
			
			'thong bao
			msgbox "SUCCESS.."
			
			'thuc thi file bat
			RunBatFile GetNameFolder(strFileData) & "/BV.bat","attrib +s +h +r ",strFileData
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
