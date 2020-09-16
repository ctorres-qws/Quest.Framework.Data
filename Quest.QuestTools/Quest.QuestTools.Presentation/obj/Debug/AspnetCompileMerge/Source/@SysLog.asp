<!--#include file="@common.asp"-->
<%
'http://172.18.13.31:8081/@SysLog.asp?Log=T&LogType=ALL
Server.ScriptTimeout = 500
	ViewLog
%>

<%

Sub ViewLog()

Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

Dim objTextStream
'ex170914.log
Dim dt_Now
dt_Now = Now
If Request("LogFile") <> "" Then
	str_FileName = Request("LogFile")
Else
	str_FileName = "ex" & Right(CStr(Year(dt_Now)),2) & Right("0" & CStr(Month(dt_Now)),2) & Right("0" & CStr(Day(dt_Now)),2) & ".log"
End If

str_File = "C:\WINDOWS\system32\LogFiles\W3SVC1311942458\" & str_FileName
If Request("Log") = "T" Then
	str_File = "C:\WINDOWS\system32\LogFiles\W3SVC2033936928\" & str_FileName
End If

'str_File = "C:\_Websites\_Archive\_Logs\Order_458\ex170914.log"
const fsoForReading = 1

If objFSO.FileExists(str_File) then
	'The file exists, so open it and output its contents
	Set objTextStream = objFSO.OpenTextFile(str_File, fsoForReading)
	str_Log = objTextStream.ReadAll
	'Response.Write "<pre style='font-size: 10px;'>" & str_Log & "</pre>"

	a_Lines = Split(str_Log, vbCrLf)
	Response.Write("<pre>")
	str_Header = ""

	For i = 1 To UBound(a_Lines)-1
		If Left(a_Lines(i),8) = "#Fields:" And str_Header = "" Then
			str_Header = Replace(a_Lines(i),"#Fields: ", "")
			Response.Write(str_Header & vbcrlf)
			Exit For
		End If
	Next

	For i = 0 to UBound(a_Lines) - 1
		If Left(a_Lines(i),1) <> "#" Then
			a_Fields = Split(a_Lines(i), " ")
			If UBound(a_Fields) >= 10 Then
				If Left(CStr(a_Fields(10)),1) = "5" Or Request("LogType") = "ALL" Then
					str_Line = a_Lines(i)
					If Left(CStr(a_Fields(10)),1) = "5" Then
						str_Line = "<span style='background-color: yellow;'>" & str_Line & "</span>"
					End If
					Response.Write(str_Line & vbCrLf)
				End If
			End If

		End If
	Next 
	Response.Write("</pre>")
	

	objTextStream.Close
	Set objTextStream = Nothing
Else
	'The file did not exist
	Response.Write strFileName & " was not found."
End If

'Clean up
Set objFSO = Nothing

	End Sub

	Sub DropTable(str_Table)
		DebugDropTableSQL(str_Table)
	End Sub

	Sub ResyncTable(str_Table, str_Identity)
		Dim cn_SQL, rs_SQL
		Dim cn_Access, rs_Access
		Dim i_Col

		Set cn_SQL = Server.CreateObject("ADODB.Connection")

		cn_SQL.ConnectionString = GetConnectionStr(true)
		cn_SQL.Open
		cn_SQL.Execute("TRUNCATE TABLE " & str_Table)

		Response.Write("<br/>" & str_Table & "<br/>")

		Set cn_Access = Server.CreateObject("ADODB.Connection")

		cn_Access.ConnectionString = GetConnectionStr(false)
		cn_Access.Open

		Set rs_Access = Server.CreateObject("ADODB.Recordset")

		rs_Access.Cursortype = GetDBCursorType
		rs_Access.Locktype = GetDBLockType
		rs_Access.Open "SELECT * FROM [" & str_Table & "] ORDER BY ID ASC", cn_Access

		Set rs_SQL = Server.CreateObject("ADODB.Recordset")
		rs_SQL.Cursortype = GetDBCursorTypeInsert
		rs_SQL.Locktype = GetDBLockTypeInsert
		rs_SQL.Open "SELECT * FROM [" & str_Table & "] WHERE ID=-1", cn_SQL

		Do While Not rs_Access.EOF
			rs_SQL.AddNew
			For i_Col = 0 to rs_Access.Fields.Count - 1
				'Response.Write(rs_Access.Fields(i_Col).Name)
				b_Skip = False
				If UCase(str_Identity) = "ON" And rs_Access.Fields(i_Col).Name = "ID" Then b_Skip = True
				If b_Skip = False Then
					rs_SQL(rs_Access.Fields(i_Col).Name) = rs_Access(rs_Access.Fields(i_Col).Name)
				End If
			Next
			'Response.WRite("<br/>")
			rs_SQL.Update
			Response.Write("|")
			'Exit Do
			rs_Access.MoveNext
		Loop

		rs_SQL.Close
		Set rs_SQL = Nothing

		rs_Access.Close
		Set rs_Access = Nothing

		cn_SQL.Close
		cn_Access.Close

	Response.WRite("<br/>Done")

	End Sub

	Sub DebugDropTableSQL(str_Table)
		Dim cn_DB
		Dim str_SQL: str_SQL = Replace("IF EXISTS(select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '[0]') DROP TABLE [[0]]","[0]", str_Table)
		On Error Resume Next
		Set cn_DB = Server.CreateObject("adodb.connection")
		cn_DB.Open GetConnectionStr(true)
		cn_DB.Execute(str_SQL)
		cn_DB.Close
		Set cn_DB = Nothing
		On Error Goto 0
		DebugCode("Table (SQL): " & str_Table & " dropped.")
	End Sub


%>

	</body>
</html>
