<!--#include file="@common.asp"-->
<html>
	<head>
		<style>
			body { font-family: arial; }
		</style>
	</head>
	<body>

<div>ReSync Table</div>
		<form action="@ReSyncTable.asp" method="get">
			<input type="text" name="Table" value="<%'= Request("Table") %>">
			<input type="checkbox" name="Identity">Identity
			<input type="hidden" name="Action" value="RESYNC">
			<input type="submit" name="Submit" value="ReSync">
		</form>

<div>Drop Table</div>
		<form action="@ReSyncTable.asp" method="get">
			<select name="Processing">
				<option value="CUT">CUT</option>
				<option value="DMSAW">DMSAW</option>
			</select>
			<input type="text" name="Table">
			<input type="hidden" name="Action" value="DROP_TABLE">
			<input type="submit" name="Submit" value="Drop Table">
		</form>

<div>View Log</div>
		<form action="@ReSyncTable.asp" method="get">
			Log File (ex170914.log):<input type="text" name="LogFile" style="width: 100px;" value="<%= Request("LogFile") %>">&nbsp;Page:<input type="text" name="Page" style="width: 100px;" value="<%= Request("Page") %>">
			<select name="Log">
				<option value="O">Order</option>
				<option value="T" <% If Request("Log") = "T" Then Response.Write(" selected") End If %>>Tools</option>
			</select>
			<select name="LogType">
				<option value="ALL">All</option>
				<option value="ERR">Error</option>
			</select>
			<select name="Lines">
				<option value="ALL">All</option>
				<option value="50" selected>50</option>
				<option value="100">100</option>
				<option value="200">200</option>
			</select>
			<input type="hidden" name="Action" value="VIEW_LOG">
			<input type="submit" name="Submit" value="View Log">
		</form>
<%

Server.ScriptTimeout = 500

Dim str_Action, str_Table
str_Action = Request("Action")
str_Table = Request("Table")
Select Case(UCase(Request("Action")))
	Case "RESYNC"
		If (str_Table <> "") Then
			Call ResyncTable(str_Table, Request("Identity"))
		End If
	Case "DROP_TABLE"
		str_Table = Request("Processing") & "_" & str_Table
		If (str_Table <> "") Then
			Response.Write(str_Table)
			DropTable(str_Table)
		End If
	Case "VIEW_LOG"
		ViewLog
End Select
%>

<%

	Sub ViewLog()

Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

Dim objTextStream
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
	str_Header = ""

	Response.Write("<table width='100%'><tr><td style='background-color: #000000;'><table style='font-size: 11px; background-color: white;'  width='100%' cellspacing='1' cellpadding='0'>")
	If Request("Log") = "T" Then
		a_Show = Array(0,1,10,13,2,4,5,6)
	Else
		a_Show = Array(0,1,10,15,2,4,5,6)
	End If

	For i = 1 To UBound(a_Lines)-1
		If Left(a_Lines(i),8) = "#Fields:" And str_Header = "" Then
			str_Header = Replace(a_Lines(i),"#Fields: ", "")
			a_Fields = Split(str_Header, " ")

			If UBound(a_Fields) > 13 Then
				For j = 0 To UBound(a_Show)-1
					str_Line = str_Line & "<td>" & a_Fields(a_Show(j)) & "<td>"
				Next
			End If

			Response.Write("<tr style='background-color: #ececec;'>" & str_Line & "</tr>")

			Exit For
		End If
	Next

	i_Lines = 0

	For i = UBound(a_Lines) - 1 to 0 Step Size -1
		b_Skip = False
		str_Line = ""
		If IsNumeric(Request("Lines")) Then
			If i_Lines > CInt(Request("Lines")) Then Exit For
		End If
		If Left(a_Lines(i),1) <> "#" Then
			a_Fields = Split(a_Lines(i), " ")
			If UBound(a_Fields) >= 13 Then
				str_Status = a_Fields(10)
				str_Page = Replace(a_Fields(4),"/","")
				If Left(CStr(str_Status),1) = "5" Or Request("LogType") = "ALL" Then

					If str_Status = "404" Or UCase(str_Page) = "@SYNCREPORT.ASP" Or UCase(str_Page) = "@RESYNCTABLE.ASP" Or UCase(str_Page) = "RESYNCTABLE.ASP" Or UCase(Right(str_Page,4)) = ".JPG" Or UCase(Right(str_Page,4)) = ".GIF" Or (Request("Page") <> "" And UCase(Request("Page")) <> UCase(str_Page)) Then b_Skip = True



					If b_Skip = False Then
						i_Lines = i_Lines + 1

						For j = 0 To UBound(a_Show)-1
							str_Line = str_Line & "<td>" & a_Fields(a_Show(j)) & "<td>"
						Next

						If Left(CStr(a_Fields(10)),1) = "5" Then
							str_Line = "<tr style='background-color: yellow;'>" & str_Line & "</tr>"
						Else
							str_Line = "<tr>" & str_Line & "</tr>"
						End If
						Response.Write(str_Line & vbCrLf)
					End If
				End If
			End If

		End If
	Next 
	Response.Write("</table></td></tr></table>")
	

	objTextStream.Close
	Set objTextStream = Nothing
Else
	'The file did not exist
	Response.Write str_FileName & " was not found."
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
		cn_SQL.Execute("TRUNCATE TABLE [" & str_Table & "]")

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
