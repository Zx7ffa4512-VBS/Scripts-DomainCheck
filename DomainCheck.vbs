RunWithCscript()

If WScript.Arguments.Count = 1 Then 
	Suf=WScript.Arguments(0)
Else 
	usage()
	WScript.StdOut.Write "Enter Suffix:"
	Suf=WScript.StdIn.Readline()
End If
WScript.Echo Suf
Do While Suf<>""
	WScript.StdOut.Write vbcrlf & "Domain:"
	Domain=WScript.StdIn.ReadLine()
	ret=HttpPost("http://www.now.cn/whois/nowcheck.net","query=" & Domain & "&domain%5B%5D=" & Suf)
	val=Trim(Replace(FindString(ret,"[\u4e00-\u9fa5]{4}(\s+\n\s+)?(?=\))"), Chr(10) ,""))
	'WScript.Echo "|"&val&"|"         '"(δ|��)��ע��(?=\))"    '"[\u4e00-\u9fa5]{4}(?=\))"    '"�ѱ�ע��(?=\))|δ��ע��(?=\s+)"
	If val="δ��ע��" Then
		WScript.Echo Domain & Suf & " ---> ��"
	ElseIf val="�ѱ�ע��" Then
		WScript.Echo Domain & Suf & " ---> X"
	Else 
		WScript.Echo "Error:" & val
	End If 
Loop 

Sub Usage()
	WScript.Echo String(79,"*")
	WScript.Echo "Usage:"
	WScript.Echo "cscript "&Chr(34)&WScript.ScriptFullName&Chr(34)&" [.com|.org]"
	WScript.Echo String(79,"*")&vbCrLf 
End Sub

'------------------------------------------------------------------------
'ǿ����cscript����
'------------------------------------------------------------------------
Sub RunWithCscript()
	If (LCase(Right(WScript.FullName,11))="wscript.exe") Then 
		Set objShell=WScript.CreateObject("wscript.shell")
		If WScript.Arguments.Count=0 Then 
			objShell.Run("cmd.exe /k cscript //nologo "&chr(34)&WScript.ScriptFullName&chr(34))
		Else
			Dim argTmp
			For Each arg In WScript.Arguments
				argTmp=argTmp&arg&" "
			Next 
			objShell.Run("cmd.exe /k cscript //nologo "&chr(34)&WScript.ScriptFullName&chr(34)&" "&argTmp)
		End If
		WScript.Quit
	End If
End Sub
Domain=WScript.Arguments(0)

'------------------------------------------------------------------------
'��sSource��sPartnƥ�䣬����ƥ�����ֵ��ÿ��һ��
'------------------------------------------------------------------------
Function FindString(sSource,sPartn)
	Dim RegEx,Match,Matches,ret
	Set RegEx=New RegExp
	RegEx.MultiLine = True
	RegEx.Pattern = sPartn
	RegEx.IgnoreCase=1
	RegEx.Global=1
	Set Matches=RegEx.Execute(sSource)
	For Each Match In Matches 
		ret = ret & Match.Value 
	Next
	FindString = ret
End Function
'------------------------------------------------------------------------
'post get���ð�
'------------------------------------------------------------------------
Function HttpPost(url,data)
	Dim http
	Set http=CreateObject("Msxml2.ServerXMLHTTP")
	http.setOption 2,13056	'����https����
	http.open "POST",url,False
	http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	http.send data
	http.waitForResponse 50
	Cs=JudgeCharset(http.responseBody)
	HttpPost = BytesToStr(http.responseBody,Cs)
End Function 


'------------------------------------------------------------------------
'�ж��ַ�����
'------------------------------------------------------------------------
Function JudgeCharset(sSource)
	Dim Str
	With CreateObject("adodb.stream")
		.Type = 1 : .Open
		.Write sSource : .Position = 0
		.Type = 2 : .Charset = "utf-8"
		Str = .ReadText : .Close
	End With
	
	Dim RegEx,Match,Matches,SubMatch,ret,ret2
	Set RegEx=New RegExp
	RegEx.MultiLine = True
	RegEx.Pattern = "Charset=\x22?(utf-8|unicode|gb2312|gbk)\x22?"
	RegEx.IgnoreCase=1
	RegEx.Global=1
	Set Matches=RegEx.Execute(Str)
	If Matches.Count<>0 Then JudgeCharset=Matches(0).Submatches(0)
End Function

'------------------------------------------------------------------------
'ת���õ� 
'------------------------------------------------------------------------
Function BytesToStr(Str,charset)
	If charset="" Then charset="utf-8"
	With CreateObject("adodb.stream")
		.Type = 1 : .Open
		.Write Str : .Position = 0
		.Type = 2 : .Charset = charset
		BytesToStr = .ReadText : .Close
	End With
End Function