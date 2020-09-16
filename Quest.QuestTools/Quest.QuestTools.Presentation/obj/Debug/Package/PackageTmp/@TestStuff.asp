<%

'Response.Write("Referrer:" & Request.ServerVariables("HTTP_REFERER"))

	'Response.Write(GetScriptNameOnlyV2("http://172.18.13.31:8081/xa7.asp?employeeID=4350&"))

Call DebugLog("test", 1)

	Function GetScriptNameOnlyV2(str_URL)
		Dim str_Ret
		On Error Resume Next
		Dim a_Parts: a_Parts = Split(str_URL, "/")

		str_Ret = a_Parts(UBound(a_Parts))

		a_Parts = Split(str_Ret & "?", "?")
		str_Ret = a_Parts(0)

		GetScriptNameOnlyV2 = str_Ret
		On Error Goto 0
	End Function

	Function DebugLog(str_Msg, i_Tab)
		Dim o_FSO
		Set o_FSO = Server.CreateObject("Scripting.FileSystemObject") 
		Const fsoForReading = 1
		Const fsoForWriting = 2
		Const fsoForAppending = 8

		Dim o_TextStream

		Set o_TextStream = o_FSO.OpenTextFile(Server.MapPath("_Logs\Debug.log"), fsoForAppending, true)

		o_TextStream.WriteLine(String(i_Tab, vbTab) & Now & ": " & str_Msg)
		o_TextStream.Close
		Set o_TextStream = Nothing
		Set o_FSO = Nothing
	End Function

	Function ReadFile(str_File)
		Dim str_Ret
		Dim o_FSO
		Set o_FSO = Server.CreateObject("Scripting.FileSystemObject") 
		Const fsoForReading = 1 

		If o_FSO.FileExists(str_File) Then
			Dim o_TextStream
			Set o_TextStream = o_FSO.OpenTextFile(str_File, fsoForReading)
			str_Ret = o_TextStream.ReadAll
			o_TextStream.Close
			Set o_TextStream = Nothing
		End If

		Set o_FSO = Nothing
		ReadFile = str_Ret
	End Function

%>