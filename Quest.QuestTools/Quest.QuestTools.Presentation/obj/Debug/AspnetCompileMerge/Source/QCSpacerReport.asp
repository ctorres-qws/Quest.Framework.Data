<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 25th, by Michael Bernholtz - Reports shows Information from the QC_SPACER Table and allow Labels to be Printed-->
<!-- QC_SPACER Table created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->
<!-- One of 3 Tables - QC_GLASS, QC_SPACER, QC_SEALANT-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>QC Spacer Report</title>
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
function PrintLabelCheck(QCID){
	var retVal = confirm("This Label has already been printed. \n Do you want to to print again?");
	if( retVal == true ){
		PrintLabel(QCID)
		return true;
	}else{
		return false;
	}
	}
function PrintLabel(QCID){

		window.location.href = "QCSpacerPrintLabels.asp?QCID="+QCID
	}

	</script>	
    <%
	
threemonth = DateAdd("m",-3,Date)
	
Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = FixSQLCheck("SELECT MSP.ItemName, MSP.Manufacturer, SP.Identifier, SP.EntryDate, SP.ConsumeDate, SP.printed, SP.Id FROM QC_SPACER AS SP INNER JOIN QC_MASTER_SPACER AS MSP ON MSP.id = SP.MasterID WHERE SP.ENTRYDATE > #" & threemonth & "# ORDER BY MASTERID, SP.ID ASC", isSQLServer)
rs2.Cursortype = GetDBCursorType
rs2.Locktype = GetDBLockType
rs2.Open strSQL2, DBConnection


%>
 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">QC INVENTORY REPORT</h1>
         <a class="button leftButton" type="cancel" href="index.html#_QC" target="_self">QC</a>
        </div>
   
   
         
       <form id="SpacerPrint" action="QCSpacerPrint.asp" name="SpacerPrint"  method="GET" target="_self" selected="true" > 

        <ul id="Profiles" title="QC Report - Spacer" selected="true">
        
<% 
listnumber =0
response.write "<li class='group'>QC Spacer REPORT </li>"
response.write "<li><table border='1' class='sortable' width='100%'><tr><th>#</th><th>Item Name</th><th>Identifier</th><th>Manufacturer</th><th>Entry Date</th><th>Consumed Date</th><th><input type='submit' value = 'Print' onClick='SpacerPrint.submit()'></input><BR></th><th>Single Print</th></tr>"

do while not rs2.eof
	listnumber = listnumber +1
	response.write "<tr><td> " & listnumber & " </td><td>" & rs2("ItemName") &"</td><td>" & rs2("Identifier") & "</td><td>" & rs2("Manufacturer") & "</td><td>" & rs2("EntryDate") & "</td><td>" & rs2("ConsumeDate") & "</td>"
	
	response.write "<td><input type ='checkbox' name = 'QCID' value = '" & trim(RS2.fields("ID")) & "' </td>"
	

	if rs2("Printed") = 1  then
		response.write "<td><input type ='button' value = 'Print Label' onclick='PrintLabelCheck(" & trim(RS2.fields("ID")) & ")'</td>"
	else
		response.write "<td><input type ='button' value = 'Print New Label' style='background-color:yellow' onclick='PrintLabel(" & trim(RS2.fields("ID")) & ")'</td>"
	end if
	rs2.movenext

loop
	rs2.close
	set rs2 = nothing
	DBConnection.close
	set DBConnection = nothing
%>
 </table></li>
</ul>
</form>         
</body>
</html>
