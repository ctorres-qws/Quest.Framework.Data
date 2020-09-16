<!--#include file="@common.asp"-->
<%
'Create  DSN Less connection to Access Database
'Create DBConnection Object
Set DBConnection = Server.CreateObject("adodb.connection")
DSN = "DRIVER={Microsoft Access Driver (*.mdb)}; "
DSN = DSN & "DBQ=F:\database\quest.mdb"
DSN = DSN & ";PWD=stewart"
DSN = GetConnectionStr(false)

'DebugMsg(DSN)

DBConnection.Open DSN

'Create  DSN Less connection to Access Database
'Create DBConnection Object
Set DBConnection2 = Server.CreateObject("adodb.connection")
DSN2 = "DRIVER={Microsoft Access Driver (*.mdb)}; "
DSN2 = DSN2 & "DBQ=F:\database\InventoryReports.mdb"
DSN2 = DSN2 & ";PWD=stewart"
DSN2 = GetConnectionStr(true)

DBConnection2.Open DSN2

%>

