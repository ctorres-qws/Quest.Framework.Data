<%
'Create  DSN Less connection to Access Database
'Create DBConnection Object
Set DBConnection = Server.CreateObject("adodb.connection")
DSN = "DRIVER={Microsoft Access Driver (*.mdb)}; "
DSN = DSN & "DBQ=" & Server.Mappath("database2/quest.mdb")
DSN = DSN & ";PWD=stewart"
DBConnection.Open DSN


'Connect to LASSARD for PREF Database
Set DBConnection2 = Server.CreateObject("adodb.connection")
DSN2 = "DRIVER={SQL Server}; "
DSN2 = DSN2 & "server=qwtordb1;UID=qws-dev;PWD=welcome1;Database=Quest-dev"
DBConnection2.Open DSN2

'Expired Permission user WEB
'server=qwtordb1;DRIVER=SQL SERVER;DATABASE=Quest;UID=Anonymous;PWD=somepass
%>

