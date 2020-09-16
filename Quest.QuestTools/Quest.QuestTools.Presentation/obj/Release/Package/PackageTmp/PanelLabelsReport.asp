<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<!-- Created October 2018 by Michael Bernholtz - Reports shows Information for Panels sent to NASHUA from Pending and allows Labels to be Printed-->
<!-- Z_Jobs Column LabelPrint created for David Ofir and Jody Cash, Implemented by Michael Bernholtz-->
<!-- This page finds all items marked "No" and gives Print option-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Nashua Panels</title>
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
	
	
	<script>
	// The sorting function conflicts with the Form Submit, so this method of Printing Labels replaces the traditional - Form submit
function PrintLabelCheck(PID){
	var retVal = confirm("This Label has already been printed. \n Do you want to to print again?");
	if( retVal == true ){
		PrintLabel(PID)
		return true;
	}else{
		return false;
	}
	}
function PrintLabel(PID){

		window.location.href = "PanelPrintLabels.asp?PID="+PID
	}

	</script>
    <%
	


Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQLCheck("SELECT * FROM Y_INV WHERE WAREHOUSE = 'NASHUA' AND LabelPrint = 'No' or LabelPrint = 'Yes' ORDER BY ID ASC", isSQLServer)
strSQL = FixSQLCheck("SELECT I.* ,M.Part ,M.INVENTORYTYPE FROM Y_INV AS I INNER JOIN Y_MASTER AS M ON I.Part = M.Part WHERE M.INVENTORYTYPE = 'Sheet' AND I.WAREHOUSE = 'NASHUA' ORDER BY I.ID ASC", isSQLServer)
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection


%>
 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Nashua Panels</h1>
        <a class="button leftButton" type="cancel" href="index.html#_Panel" target="_self">Panel</a>
        </div>
   
   
         
       
        <form id="PanelPrint" action="PanelPrintLabels.asp" name="GlassPrint"  method="GET" target="_self" selected="true" > 

        <ul title="Nashua Panels- Print Labels">
     <% 
listnumber =0
response.write "<li class='group'>Panel Labels </li>"
response.write "<li><table border='1' class='sortable' width='100%'><tr><th>#</th><th>Part</th><th>Size</th><th>PO</th><th>Color</th><th>Color Code</th><th>Color PO</th><th>EnterDate</th><th>Qty</th><th>Single Print</th></tr>"

do while not rs.eof

	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT * FROM Y_COLOR WHERE PROJECT = '" & RS("colour") & "' ORDER BY ID ASC"
	rs2.Cursortype = GetDBCursorType
	rs2.Locktype = GetDBLockType
	rs2.Open strSQL2, DBConnection
	ColourCode = rs2("Code")
	rs2.close
	set rs2 = nothing

	listnumber = listnumber +1
	response.write "<tr><td> " & listnumber & " </td>"
	response.write "<td>" & RS("Part") &"</td>"
	response.write "<td>" & RS("Width") & " X " & RS("HEIGHT") & " </td>"
	response.write "<td>" & RS("PO") & "</td>"
	response.write "<td>" & RS("Colour") & "</td>"
	response.write "<td>" & ColourCode & "</td>"
	response.write "<td>" & RS("ColorPO") & "</td>"
	response.write "<td>" & RS("DateIN") & "</td>"
	response.write "<td>" & RS("QTY") & "</td>"
		
	if rs("LabelPrint") = "Yes"  then
		response.write "<td><input type ='button' value = 'Print Label' onclick='PrintLabelCheck(" & trim(RS.fields("ID")) & ")'</td>"
	else
		response.write "<td><input type ='button' value = 'Print New Label' style='background-color:yellow' onclick='PrintLabel(" & trim(RS.fields("ID")) & ")'</td>"
	end if
	rs.movenext
loop

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
