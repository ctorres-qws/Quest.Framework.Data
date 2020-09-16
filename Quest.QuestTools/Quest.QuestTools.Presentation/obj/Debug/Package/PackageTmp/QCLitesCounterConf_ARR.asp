<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created March 5th, 2014 - by Michael Bernholtz --> 

<html xmlns="http://www.w3.org/1999/xhtml">
<head>

</head>
<body>
  	 <!--#include file="QCdbpath.asp"-->
<%

ReturnSite = "QCLitesCounter.asp"
qcid = request.querystring("qcid")
action = request.querystring("action")

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		'Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

Function Process(isSQLServer)

DBOpenQC DBConnection, isSQLServer

Set rsComplete = Server.CreateObject("adodb.recordset")
strSQL = "Select Lites from QC_MASTER_GLASS WHERE ID = " & QCID
rsComplete.Cursortype = 2
rsComplete.Locktype = 3
rsComplete.Open strSQL, DBConnection
Lites = RSComplete("Lites")
if Lites = "" or isnull(Lites) then
Lites = 0
end if
if action = "plus" then
RSComplete("Lites") = Lites + 1
end if
if action = "minus" then
RSComplete("Lites") = Lites - 1
end if

RSComplete.update
RSComplete.close
Set RSComplete = nothing

DbCloseAll

'DBConnection.close
'set DBConnection=nothing

End Function

Response.Redirect Returnsite
%>
</body>
</html>