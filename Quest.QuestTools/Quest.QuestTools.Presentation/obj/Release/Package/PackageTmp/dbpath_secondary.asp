<!--#include file="@common.asp"-->
<%

'Create  DSN Less connection to Access Database
'Create DBConnection Object
Set DBConnection = Server.CreateObject("adodb.connection")
DSN = GetConnectionStr(b_SQL_Server) 'method in @common.asp

DBConnection.Open DSN

Set DBConnection2 = Server.CreateObject("adodb.connection")
DSN = GetConnectionStrSecondary(b_SQL_Server) 'method in @common.asp

DBConnection2.Open DSN

%>
<!--#include file="CountryLocation.INC"-->