<!--#include file="@common.asp"-->
<html>
	<head>
		<style>
			body { font-family: arial; }
		</style>
	</head>
	<body>

<div>View Log</div>
		<form method="get" name="fMain">
<%
	dt_Now = Now
	str_Date = "ex" & Right(Year(dt_Now),2) & Right("0" & Month(dt_Now),2) & Right("0" & Day(dt_Now),2) & ".log"
%>
			Log File (<a href="javascript: void();" onclick="fMain.LogFile.value = '<%= str_Date %>'"><%= str_Date %></a>):<input type="text" name="LogFile" style="width: 100px;" value="<%= Request("LogFile") %>">&nbsp;Page:<input type="text" name="Page" style="width: 100px;" value="<%= Request("Page") %>">&nbsp;Time Taken:<input type="text" name="TimeTaken" style="width: 60px;" value="<%= Request("TimeTaken") %>">
			<select name="Log">
				<option value="O">Order</option>
				<option value="T" <% If Request("Log") = "T" Or Request("Log") = "" Then Response.Write(" selected") End If %>>Tools</option>
			</select>
			<select name="LogType">
				<option value="ALL">All</option>
				<option value="ERR"<% If Request("LogType") = "ERR" Then Response.Write(" selected") End If %>>Error</option>
			</select>
			<select name="Lines">
				<option value="ALL">All</option>
				<option value="50" <% If Request("Lines") = "50" Then Response.Write(" selected") End If %>>50</option>
				<option value="100" <% If Request("Lines") = "100" Then Response.Write(" selected") End If %>>100</option>
				<option value="200" <% If Request("Lines") = "200" Then Response.Write(" selected") End If %>>200</option>
			</select>
			<input type="hidden" name="Action" value="VIEW_LOG">
			<input type="submit" name="Submit" value="View Log">
		</form>
<%

Server.ScriptTimeout = 500
'C:\_Websites\_Archive\_Logs\Order_458\
'C:\_Websites\_Archive\_Logs\Tools_928\
Dim str_Action, str_Table
str_Action = Request("Action")
str_Table = Request("Table")
Select Case(UCase(Request("Action")))
	Case "VIEW_LOG"
		ViewLog
End Select
%>

<%

	Sub ViewLog()

Dim objTextStream
dt_Now = Now
If Request("LogFile") <> "" Then
	str_FileName = Request("LogFile")
Else
	str_FileName = gstr_IIS_Log_Prefix & "ex" & Right(CStr(Year(dt_Now)),2) & Right("0" & CStr(Month(dt_Now)),2) & Right("0" & CStr(Day(dt_Now)),2) & ".log"
End If

Dim str_LogPath_O, str_LogPath_Archive_O
Dim str_LogPath_T, str_LogPath_Archive_T

Dim o_FS
Set o_FS = Server.CreateObject("Scripting.FileSystemObject")

str_LogPath_O = gstr_IIS_Log_Folder_Order & str_FileName
str_LogPath_Archive_O = gstr_IIS_Log_Folder_Order_Archive & str_FileName

str_LogPath_T = gstr_IIS_Log_Folder_Tools & str_FileName
str_LogPath_Archive_T = gstr_IIS_Log_Folder_Tools_Archive & str_FileName

If Request("Log") = "T" Then
	If o_FS.FileExists(str_LogPath_T) Then
		o_FS.CopyFile str_LogPath_T, str_LogPath_Archive_T, True
	End If
	str_File = str_LogPath_Archive_T
Else
	If o_FS.FileExists(str_LogPath_O) Then
		o_FS.CopyFile str_LogPath_O, str_LogPath_Archive_O, True
	End If
	str_File = str_LogPath_Archive_O
End If
'str_File = "C:\_Websites\_Archive\_Logs\Order_458\ex170914.log"
const fsoForReading = 1

If o_FS.FileExists(str_File) then
	'The file exists, so open it and output its contents
	Set objTextStream = o_FS.OpenTextFile(str_File, fsoForReading)
	str_Log = objTextStream.ReadAll
	'Response.Write "<pre style='font-size: 10px;'>" & str_Log & "</pre>"

	a_Lines = Split(str_Log, vbCrLf)
	str_Header = ""

	Response.Write("<table width='100%'><tr><td style='background-color: #000000;'><table style='font-size: 11px; background-color: white;'  width='100%' cellspacing='1' cellpadding='0'>")
	If Request("Log") = "T" Then
		a_Show = Array(0,1,10,13,8,4,5,6)
	Else
		a_Show = Array(0,1,10,15,8,4,5,6)
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
				str_TimeTaken = a_Fields(a_Show(3))
				If Left(CStr(str_Status),1) = "5" Or Request("LogType") = "ALL" Then

					If str_Status = "404" Or UCase(str_Page) = "@SYNCREPORT.ASP" Or UCase(str_Page) = "@RESYNCTABLE.ASP" Or UCase(str_Page) = "RESYNCTABLE.ASP" Or UCase(str_Page) = "A_IIS.ASP" Or UCase(Right(str_Page,4)) = ".PNG" Or UCase(Right(str_Page,4)) = ".JPG" Or UCase(Right(str_Page,4)) = ".GIF" Or UCase(Right(str_Page,4)) = ".CSS" Or UCase(Right(str_Page,3)) = ".JS" Or (Request("Page") <> "" And UCase(Request("Page")) <> UCase(Left(str_Page,Len(Request("Page"))))) Then b_Skip = True

					If(Request("TimeTaken") <> "") Then 
						If CLng(str_TimeTaken) < CLng(Request("TimeTaken")) Then b_Skip = True
					End If

					If b_Skip = False Then
						i_Lines = i_Lines + 1

						For j = 0 To UBound(a_Show)-1
							If a_Show(j) = 1 Then
								str_Line = str_Line & "<td>" & FormatTime(a_Fields(a_Show(j))) & "<td>"
							Else 
								str_Line = str_Line & "<td>" & LCase(a_Fields(a_Show(j))) & "<td>"
							End If
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
	Response.Write str_File & " was not found."
End If

'Clean up
Set o_FS = Nothing

	End Sub

	Function FormatTime(str_Time)
		Dim str_Ret: str_Ret = ""
		Dim a_Parts: a_Parts = Split(str_Time, ":")
		str_Ret = a_Parts(0)-4 & ":" & a_Parts(1) & ":" & a_Parts(2)
		FormatTime = str_Ret
	End Function

%>

	</body>
</html>
