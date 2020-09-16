<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 7th, by Michael Bernholtz - Edit and Delete Form for items in QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>QC Inventory</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    


<%
	


Dim InventoryType
InventoryType = Request.Querystring("InventoryType")


%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">QC Inventory Edit</h1>
        <a class="button leftButton" type="cancel" href="Index.HTML#_QC" target="_self">QC</a>
    </div>
	
	
	
<%

		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM QC_MASTER_Glass ORDER BY ITEMNAME ASC"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		
		Response.Write " <ul id='Glass' title=' QC Glass Inventory' selected='true'> "

		Response.Write "<li><table border = '1'>"
		Response.Write "<tr><th>Glass Name</th><th>Producer</th><th>Extra Lites</th><th>Plus</th><th>Minus</th></tr>"
		do while not rs.eof
			Response.write "<tr><td><a href='QCLitesEditForm.asp?InventoryType=" & InventoryType & "&qcid=" & rs.fields("ID") & "' target='_self'>" & rs.fields("ItemName") & "</a></td><td>" & rs.fields("Manufacturer")& " </td><td> " & rs.fields("Lites") & "</td>"
			Response.write "<td><a href='QCLitesCounterConf.asp?qcid=" & rs.fields("ID") &"&Action=plus' target ='_self'> Plus</a></td>" 
			Response.write "<td><a href='QCLitesCounterConf.asp?qcid=" & rs.fields("ID") &"&Action=minus' target ='_self'> Minus</a></td>" 
			Response.write "</tr>" 

		rs.movenext
		loop
		rs.close
		Response.write "</table></li>"

%>
  </ul>          
         
               
</body>
</html>
