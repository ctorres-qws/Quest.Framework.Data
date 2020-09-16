<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		<!--DurapaintWIP Sent Selection Page, shows all items that are in Durapaint WIP but have not been marked Shipped and a checkbox-->
		<!--Created February 4 2018, at Request of Shaun Levy, Viviana Davids, Ali Alibeigloo to Select Specific Durapaint WIP items and mark sent-->
		<!-- Sends to glassExportSelectConf.asp-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Durapaint (WIP)</title>
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
    <style type="text/css">
   ul{
    margin: 0;
    padding: 0;
   }
   </style>
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE ([NOTE 4] IS NULL OR [NOTE 4] = '') AND WAREHOUSE = 'DURAPAINT(WIP)' ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>

       <form id="SEND" action="DurapaintWipSentConf.asp" name="SEND"  method="POST" target="_self" selected="true" >  
        <ul id="Profiles" title=" Optima Report" selected="true">

<%
Response.write "<li class='group'>Choose DURAPAINT(WIP) from Staging area to mark Sent <input type='submit' value = 'Mark Sent' onClick='Send.submit()'></li>"
Response.write "<li style='font-size:11px' ><table border='1' class='sortable'><tr><th></th><th>PART</th><th>Colour</th><th>Length</th><th>Quantity</th><th>PO</th><th>Bundle</th><th>EX Bundle</th><th>Allocation</th><th>Colour PO</th><th>Entry Date</th><th>Modify Date</th></tr>"

Do while not rs.eof
		Response.write "<tr><td><input type='checkbox' name='OptimaSelect' value='" & RS("ID")& "'></td>"
		Response.write "<td>" & RS("PART") & "</td><td>" & RS("COLOUR") & "</td><td>" & RS("LFT") &"</td><td>" & RS("QTY") & "</td>" 
		Response.write "<td>" & RS("PO") & "''</td><td>" & RS("BUNDLE") & "''</td><td>" & RS("EXBUndle") & "</td>"
		Response.write "<td>" & RS("Allocation") & "</td><td>" & RS("ColorPO") & "</td><td>" & RS("DateIn") & "</td><td>" & RS("MODIFYdATE") & "</td>"
		Response.write "</tr>"
	rs.movenext
Loop

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>

	</table>

      </ul>
		</form>

</body>
</html>
