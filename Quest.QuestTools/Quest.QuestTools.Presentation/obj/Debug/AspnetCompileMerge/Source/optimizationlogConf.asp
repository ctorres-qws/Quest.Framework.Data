<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Confirmation page for the entry into Optimization log from OptimizationLogForm.asp-->
<!-- Written fro Victor, by Michael Bernholtz, August 2014 -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Optimization Log</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>
  


<% 
'
''Create a Query
'    SQL = "Select * FROM Y_INV ORDER BY PART ASC"
''Get a Record Set
'    Set RS = DBConnection.Execute(SQL)
JOB = REQUEST.QueryString("Job")
FLOOR = REQUEST.QueryString("Floor")
GLASS = REQUEST.QueryString("Glass")
LITES = REQUEST.QueryString("Lites")
if LITES = "" then 
	LITES = 0
end if
GType = REQUEST.QueryString("InventoryType")
Bendfile = REQUEST.QueryString("BendFile")
Opfile = REQUEST.QueryString("OpFile")
OpDate = REQUEST.QueryString("OpDate")
if not isDate(OpDate) then 
	OpDate = Now
end if
OpTime = Time

IsError = False


Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

if JOB ="" OR FLOOR = "" OR GLASS = "" OR LITES = "" OR Opfile = "" OR OpDate = "" then
IsError= TRUE
else

Set rs = Server.CreateObject("adodb.recordset")
'strSQL = "INSERT INTO OptimizeLog ([Job], [Floor], [Glass], [Lites], [Type], [BendFile], [Opfile], [Opdate], [OpTime]) VALUES( '" & JOB & "', '" & FLOOR &  "', '" & GLASS & "', '" & LITES & "', '" & GTYPE & "', '" & Bendfile & "', '" & OPFILE & "', '" & OPDATE & "', '" & OPTIME & "')"
strSQL = "SELECT * FROM OptimizeLog WHERE ID=-1"
rs.Cursortype = GetDBCursorTypeInsert
rs.Locktype = GetDBLockTypeInsert
rs.Open strSQL, DBConnection

rs.AddNew
rs.fields("job") = JOB
rs.fields("floor") = FLOOR
rs.fields("Glass") = GLASS
rs.fields("Lites") = LITES
rs.fields("Type") = GTYPE
rs.fields("BendFile") = Bendfile
rs.fields("Opfile") = OPFILE
rs.fields("Opdate") = OPDATE
rs.fields("OpTime") = OPTIME
If GetID(isSQLServer,1) <> "" Then rs.Fields("ID") = GetID(isSQLServer,1)
rs.Update

Call StoreID1(isSQLServer, rs.Fields("ID"))

end if

DbCloseAll

End Function

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="OptimizationLogForm.asp" target="_self">Optimize Log</a>
    </div>


    
<ul id="Report" title="Added" selected="true">
	<%
	IF IsError = TRUE then
	%>
	<li>Invalid Entry: Please Fill in All fields.</li>
	<li>All Fields are Required for entry into Database</li>
	<%
	End if
	%>
	
    <li><% response.write "Job: " & JOB %></li>
	<li><% response.write "Floor: " & FLOOR %></li>
    <li><% response.write "Glass: " & GLASS %></li>
	<li><% response.write "Glass Type: " & GTYPE %></li>
	<li><% response.write "Optimization File: " & Opfile %></li>
	<li><% response.write "Number of Lites: " & LITES %></li>
	<li><% response.write "Bending File: " & bendfile %></li>
    <li><% response.write "Optimization Date: " & Opdate %></li>
 <li><a class="whiteButton" href="OptimizationLogForm.asp" target="_self">Return to Form</a></li>
</ul>

<% 

'rs.close
'set rs=nothing
'DBConnection.close
'set DBConnection=nothing
%>

</body>
</html>



