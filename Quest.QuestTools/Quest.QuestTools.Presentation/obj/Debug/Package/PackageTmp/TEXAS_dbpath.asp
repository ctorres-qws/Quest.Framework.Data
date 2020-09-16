
<%

'Create DBConnection Object
' Connect to Texas Database Removed from DBPath and @common in place of specific use.
' March 2019 - Not best practice but better for our network - this page currently unused
' May 2019 to reach Scan_Texas.mdb, set ScanMode = True 

if ScanMode = True then
	DSN_Texas = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=\\10.34.16.11\db\Scan_Texas_Dev.mdb;"
else
	DSN_Texas = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=\\10.34.16.11\db\Texas_Dev.mdb;"
end if

Set DBConnection_Texas = Server.CreateObject("adodb.connection")
DBConnection_Texas.ConnectionTimeout = 100
DBConnection_Texas.CommandTimeout = 100

On Error Resume Next
DBConnection_Texas.Open DSN_Texas


'Sample open of Texas DB
'Set rsTEST = Server.CreateObject("adodb.recordset")
'		strSQL3 = "SELECT top 1 ID, JobNumber FROM PRODMT1"
'		rsTEST.Cursortype = 2
'		rsTEST.Locktype = 3
'		rsTEST.Open strSQL3, DBConnection_Texas
'		
'		response.write rsTEST("JObNUmber")
'		
'		rsTEST.close
'		set rsTEST = nothing
		
'DBConnection_Texas.Close
'DbCloseAll DBConnection_Texas = nothing
%>
