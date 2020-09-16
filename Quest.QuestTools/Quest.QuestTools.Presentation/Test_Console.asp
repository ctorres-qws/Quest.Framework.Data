<!--#include file="@common.asp"-->

<%

'Create  DSN Less connection to Access Database
'Create DBConnection Object
Set DBConnection = Server.CreateObject("adodb.connection")
DSN = GetConnectionStr(false) 'method in @common.asp
DebugMsg(DSN)
DBConnection.Open DSN
Response.Write("Ya<br/>")

Set DBConnection2 = Server.CreateObject("adodb.connection")
DSN = GetConnectionStrSecondary(false) 'method in @common.asp
DebugMsg(DSN)
DBConnection2.Open DSN
Response.Write("Ya<br/>")

DebugMsg(Server.Mappath("database/quest.mdb"))

%>