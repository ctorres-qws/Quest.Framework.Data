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
strSQL = "Select * FROM X_BARCODE WHERE JOB = '" & JOB & "' AND FLOOR = '" & FLOOR & "' AND TAG Like '%" & TAG & "%' ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

response.write "<li class='group'>Glass items found </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Job/Floor/Tag</th><th>DEPT</th><th>EMPLOYEE</th><th>DATETIME</th></th></tr>"

if rs.eof then
response.write "<LI> NO Items found with that JOB Floor and Tag </li>"
else


do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & rs("JOB") & rs("FLOOR") & rs("TAG") & "</td>"
	response.write "<td>" & rs("DEPT") & "</td>"
	response.write "<td>" & rs("EMPLOYEE") & "</td>"
	response.write "<td>" & rs("DATETIME") & "</td>"
	response.write "</tr>"
rs.movenext
loop
end if

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "Select * FROM X_GLAZING WHERE JOB = '" & JOB & "' AND FLOOR = '" & FLOOR & "' AND TAG Like '%" & TAG & "%'"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

if rs2.eof then
else
do while not rs2.eof
	response.write "<tr>"
	response.write "<td>" & rs2("JOB") & rs2("FLOOR") & rs2("TAG") & "</td>"
	response.write "<td>" & rs2("DEPT") & "</td>"
	response.write "<td>" & rs2("EMPLOYEE") & "</td>"
	response.write "<td>" & rs2("DATETIME") & "</td>"
	response.write "</tr>"
rs2.movenext
loop
end if

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "Select * FROM X_BARCODEGA WHERE JOB = '" & JOB & "' AND FLOOR = '" & FLOOR & "' AND TAG Like '%" & TAG & "%' ORDER BY ID ASC"
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection

if rs3.eof then
else
do while not rs3.eof
	response.write "<tr>"
	response.write "<td>" & rs3("JOB") & rs3("FLOOR") & rs3("TAG") & "</td>"
	response.write "<td>" & rs3("DEPT") & "</td>"
	response.write "<td> N/A </td>"
	response.write "<td>" & rs3("DATETIME") & "</td>"
	response.write "</tr>"
rs3.movenext
loop
end if

Set rs4 = Server.CreateObject("adodb.recordset")
strSQL4 = "Select * FROM X_BARCODEOV WHERE JOB = '" & JOB & "' AND FLOOR = '" & FLOOR & "' AND TAG Like '%" & TAG & "%' ORDER BY ID ASC"
rs4.Cursortype = 2
rs4.Locktype = 3
rs4.Open strSQL4, DBConnection

if rs4.eof then
else
do while not rs4.eof
	response.write "<tr>"
	response.write "<td>" & rs4("JOB") & rs4("FLOOR") & rs4("TAG") & "</td>"
	response.write "<td>" & rs4("DEPT") & "</td>"
	response.write "<td>" & rs4("EMPLOYEE") & "</td>"
	response.write "<td>" & rs4("DATETIME") & "</td>"
	response.write "</tr>"
rs4.movenext
loop
end if


response.write "</table> </li>"




	
%>	
	<a class="whiteButton" href="FindStock.asp">Search Again</a>	
	</ul>

	
                
<%             
rs.close
set rs=nothing
rs2.close
set rs2=nothing
rs3.close
set rs3=nothing
rs4.close
set rs4=nothing
DBConnection.close
set DBConnection=nothing           
%>
</body>
</html>
