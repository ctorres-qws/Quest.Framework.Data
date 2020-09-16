<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created March 5th, 2014 - by Michael Bernholtz --> 
<!--Edit Page for Glass Items-->
<!-- Submits to page GlassManageForm.asp with a GID -->
		 <!--#include file="dbpath.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Manage Glass</title>
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
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Choose Glass Item to Edit</h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
    </div>
<%
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT top 10000 * FROM Z_GlassDB WHERE [HIDE] IS NULL ORDER BY ID DESC"
		rs.Cursortype = GetDBCursorType
		rs.Locktype = GetDBLockType
		rs.Open strSQL, DBConnection

		Response.Write " <ul id='Glass' title='Glass Inventory' selected='true'> "
		response.Write "<li> Click on the Headers of each column to sort Ascending/Descending</li>"
		response.write "<li><table border='1' class='sortable'><tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Order By</th><th>PO #</th><th>1 Mat Order #</th><th>2 Mat Order #</th><th>External Glass</th><th>Spacer</th><th>Internal Glass</th><th>Notes</th><th>Manage Glass</th><th>Manage Timeline</th></tr>"

		Do While Not rs.eof
			Response.write "<tr><td> " & rs.fields("id") & "</td> "
			Response.write "<td>" & rs.fields("JOB") & "</td> " ' Job
			Response.Write "<td>" & rs.fields("FLOOR") & "</td> " ' Floor
			Response.write "<td>" & rs.fields("TAG") & "</td> " ' Tag
			Response.write "<td>" & rs.fields("ORDERBY") & "</td> " ' Ordered By
			Response.write "<td>" & rs.fields("PO") & "</td> " ' Po Number
			Response.write "<td>" & rs.fields("ExtOrderNum") & "</td> " ' Ext Order Number
			Response.write "<td>" & rs.fields("IntOrderNum") & "</td> " ' Int Order Number
			Response.write "<td>" & rs.fields("1 MAT") & "</td> " ' 1 MATERIAL
			Response.write "<td>" & rs.fields("1 SPAC") & "</td> " ' 1 SPACER
			Response.write "<td>" & rs.fields("2 MAT") & "</td> " ' 2 MATERIAL
			Response.write "<td>" & rs.fields("NOTES") & "</td> " ' Notes
			Response.write "<td><a href='GlassManageForm.asp?gid=" & rs.fields("ID") & "' target='_self' >Manage Glass</a> </td>" 
			Response.write "<td><a href='GlassManageTimeLineForm.asp?gid=" & rs.fields("ID") & "' target='_self' >Manage Time Line</a> </td>" 
			Response.write "</tr>"

			rs.movenext
		Loop
		rs.close
		set rs = nothing	
		DBConnection.close
		set DBConnection = nothing		
%>
  </ul>
</body>
</html>
