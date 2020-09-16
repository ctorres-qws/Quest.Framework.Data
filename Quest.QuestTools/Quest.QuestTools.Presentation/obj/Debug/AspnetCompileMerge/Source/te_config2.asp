<%
	'Using Redim causes a performance degredation
	'But it's OK since array size is small
	redim arrDbs(iTotalConnections)
	redim arrDesc(iTotalConnections)
	redim arrConn(iTotalConnections)

	'Define database locations
	arrDBs(0) = "database/teadmin.mdb"
	arrDBs(1) = "database/quest.mdb"
	
	'Define descriptions for corresponding connections
	arrDesc(0) = "Table Editor User Administration"
	arrDesc(1) = "Quest"


%><!--#include file="te_includes2.asp"-->