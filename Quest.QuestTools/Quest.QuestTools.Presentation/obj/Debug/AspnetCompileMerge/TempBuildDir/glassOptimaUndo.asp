<%@Language="VBScript"%>
<%Response.Buffer = False%>
<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		<!--Optima Undo Selection Page, shows all items that have an Optima Date and a checkbox-->
		<!--Created July 2014, at Request of Sasha and Jody to Undo Optima Export-->
		<!-- Sends to glassOptimaUndoConf.asp-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    <style type="text/css">
	ul{
    margin: 0;
    padding: 0;
   }
   </style>


    </head>
<body>
    <div class="toolbar">
		<h1 id="pageTitle"></h1>
		<a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
    </div>

    <form id="Optima" action="glassOptimaUndoConf.asp" name="Optima"  method="GET" target="_self" selected="true" >  
    <ul id="Profiles" title=" Optima Report" selected="true">


<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT Top 1000 * FROM Z_GLASSDB WHERE OPTIMADATE <> '' and (COMPLETEDDATE = '' OR ISNULL(COMPLETEDDATE) ) ORDER BY ID DESC"
rs.Cursortype = 1
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>

<li class='group'>Choose Optima Files to mark Uncut <input type='submit' value = 'Undo Optima' onClick='Optima.submit()' /></li>
<li><table border='1' class='sortable'><tr><th></th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Input Date</th><th>Optima Date</th><th>Required Date</th><th>Output Date</th><th>ID</th><th>Type</th><th>Order</th><th>PO</th><th>NOTES</th></tr>

<%
Do While not rs.eof
	If IsDate(RS("OPTIMADATE")) Then
		Response.write "<tr><td><input type='checkbox' name='OptimaUndo' value='" & RS("ID")& "'></td>"
		Response.write "<td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TAG") & "</td><td>" & RS("DIM X") & "''</td><td>" & RS("DIM Y") & "''</td><td>" & RS("1 MAT") & "</td><td>" & RS("1 SPAC") & "</td><td>" & RS("2 MAT") & "</td>" 
		Response.write "<td>" & RS("INPUTDATE") & "</td><td>" & RS("OPTIMADATE") & "</td><td>" & RS("REQUIREDDATE") & "</td><td>" & RS("COMPLETEDDATE") & "</td><td>" & RS("ID") & "</td><td>" & RS("DEPARTMENT") & "</td><td>" & RS("ORDERBY") & "</td><td>" & RS("PO") & "</td><td>" & RS("NOTES") & "</td>"
		Response.write "</tr>"
	End If
	rs.movenext
Loop

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>

	</table></li>

		</ul>
		</form>

</body>
</html>
