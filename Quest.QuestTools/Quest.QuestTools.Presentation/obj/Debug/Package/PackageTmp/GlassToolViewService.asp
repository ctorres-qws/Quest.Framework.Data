<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Report form of all active Service Glass-->
<!-- Reuqested by Gurveen and designed by Michael Bernholtz, December 2015 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Progress</title>
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

  <script src="sorttable.js"></script>

<style type="text/css">
	ul{
    margin: 0;
    padding: 0;
   }
   </style>
 
    </head>
<body>

<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Z_GLASSDB WHERE Department = 'Service' ORDER BY ID DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
%>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
    </div>

      <form id="Complete" name="Complete"  method="GET" target="_self" selected="true" >  
        <ul id="Profiles" title="Glass Report - Service" >
		<li>Choose Completed Window items that have been SHIPPED 
	<input type='submit' value = 'Remove Selected Items' onClick="Complete.action='glassRemoveConf.asp?ticket=Service'; Complete.submit()"></li>
	<li>Choose Entire POs that have been SHIPPED 
	<input type='submit' value = 'Remove Entire PO' class = "blueButton" onClick="Complete.action='glassRemovePOConf.asp?ticket=Service'; Complete.submit()"></li>
	
<% 
response.write "<li class='group'>Glass Progress</li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable' ><thead><tr><th>Complete?</th><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Orderby</th><th>PO</th><th>1 Mat</th><th>1 Spac</th><th>2 Mat</th><th>Notes</th><th>Status</th><th>Details</th><th>Timeline</th></tr></thead><tbody>"
do while not rs.eof

	if isdate(rs("optimadate")) and not isdate(rs("shipdate"))  then
		response.write "<tr>"

		response.write "<td><input type='checkbox' name='GID' value='" & RS("ID")& "'></td>"
		Response.write "<td>" & rs.fields("id") & "</td> "
		Response.write "<td>" & rs.fields("JOB") & "</td> " ' Job
		Response.Write "<td>" & rs.fields("FLOOR") & "</td> " ' Floor
		Response.write "<td>" & rs.fields("TAG") & "</td> " ' Tag
		Response.write "<td>" & rs.fields("ORDERBY") & "</td> " ' Ordered By
		Response.write "<td>" & rs.fields("PO") & "</td> " ' Po Number
		Response.write "<td>" & rs.fields("1 MAT") & "</td> " ' 1 MATERIAL
		Response.write "<td>" & rs.fields("1 SPAC") & "</td> " ' 1 SPACER
		Response.write "<td>" & rs.fields("2 MAT") & "</td> " ' 2 MATERIAL
		Response.write "<td>" & rs.fields("NOTES") & "</td> " ' Notes
%>
<!--#include file="GlassStatus.inc"-->
<%
		Response.write "<td>" & Status & "</td> " ' Notes

		Response.write "<td><a href='GlassManageForm.asp?gid=" & rs.fields("ID") & "&ticket=waitingser' target='_self' >Manage Glass</a> </td>" 
		Response.write "<td><a href='GlassManageTimeLineForm.asp?gid=" & rs.fields("ID") & "&ticket=waitingser' target='_self' >Manage Time Line</a> </td>" 
		Response.write "</tr>"

	end if
	rs.movenext
loop
response.write "</tbody></table></li>"

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>

    </ul>
     </form>

</body>
</html>
