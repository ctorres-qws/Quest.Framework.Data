<!--#include file="@common.asp"-->
<%

Set DBConnection = Server.CreateObject("adodb.connection")
DSN = gstr_DB_Access_Prod
DBConnection.Open DSN
DBConnection.Close: Set DBConnection = Nothing
Response.Write("OK")
%>