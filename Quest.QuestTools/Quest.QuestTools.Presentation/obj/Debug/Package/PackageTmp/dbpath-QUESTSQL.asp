<%
'Connect to LASSARD for QUEST Database
Set DBConnection2 = Server.CreateObject("adodb.connection")
DSN2 = "DRIVER={SQL Server}; "
DSN2 = DSN2 & "server=qwtordb1;UID=qws-dev;PWD=welcome1;Database=QWS-dev"
DBConnection2.Open DSN2
%>