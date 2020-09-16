<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		<!-- Collects information from FindStock.asp -->
		<!--Tool to show Employee, time, and Activity for Job/Floor/Tag
		 <!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Manage Glass Types</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
	 <script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

<%
	JOB = request.querystring("JOB")
	FLOOR = request.querystring("FLOOR")
	TAG = request.querystring("TAG")
%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="Index.html#_Report" target="_self">Reports</a>
        </div>

<ul id="screen1" title="Stock by JOB FLOOR TAG" selected="true">
   
<%

Set rs = Server.CreateObject("adodb.recordset")
Counter = 0
Forel = 0
Willian = 0
If Floor = "" Then
	strSQL = "Select * FROM X_BARCODEGA WHERE JOB = '" & JOB & "' AND TAG Like '%" & TAG & "%' ORDER BY TAG ASC"
Else
	strSQL = "Select * FROM X_BARCODEGA WHERE JOB = '" & JOB & "' AND FLOOR = '" & FLOOR & "' AND TAG Like '%" & TAG & "%' ORDER BY TAG ASC"
End If
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Response.write "<li class='group'>Glass items found </li>"
Response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
Response.write "<li><table border='1' class='sortable'><tr><th>Job</th><th>Floor</th><th>Tag</th><th>Position</th><th>Glass Type</th><th>Glass Line</th><th>DATETIME</th></tr>"

If rs.eof Then
	Response.write "<LI> NO Items found with that JOB Floor and Tag </li>"
Else

	Do While Not rs.eof
		response.write "<tr>"
		response.write "<td>" & rs("Job") & "</td>"
		response.write "<td>" & rs("Floor") & "</td>"
		response.write "<td>" & rs("Tag") & "</td>"
		response.write "<td>" & rs("Position") & "</td>"
		response.write "<td>" & rs("Type") & "</td>"
		response.write "<td>" & rs("DEPT") & "</td>"
		response.write "<td>" & rs("DATETIME") & "</td>"
		response.write "</tr>"
		Counter = Counter + 1
		If UCASE (rs("DEPT")) = "FOREL" Then
			Forel = Forel + 1
		End If
		If UCASE (rs("DEPT")) = "WILLIAN" Then
			Willian = Willian + 1
		End If
		rs.movenext
	Loop
End If

response.write "</table> </li>"
%>
<li> Forel Count: <%response.write Forel %> </li>	
<li> Willian Count: <%response.write Willian %> </li>
<li> Total Window Count: <%response.write Counter %> </li>

	<a class="whiteButton" href="FindGlass.asp">Search Again</a>	
	</ul>

<%
rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>
</body>
</html>
