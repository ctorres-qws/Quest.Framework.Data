<%
'Create  DSN Less connection to Access Database
'Create DBConnection Object
Set DBConnection = Server.CreateObject("adodb.connection")
DSN = "DRIVER={Microsoft Access Driver (*.mdb)}; "
DSN = DSN & "DBQ=F:\database\quest.mdb"
DSN = DSN & ";PWD=stewart"
DBConnection.Open DSN

'Create  DSN Less connection to Access Database
'Create DBConnection Object
Set DBConnection2 = Server.CreateObject("adodb.connection")
DSN2 = "DRIVER={Microsoft Access Driver (*.mdb)}; "
DSN2 = DSN2 & "DBQ=F:\database\ArchiveLists.mdb"
DSN2 = DSN2 & ";PWD=stewart"
DBConnection2.Open DSN2

%>

