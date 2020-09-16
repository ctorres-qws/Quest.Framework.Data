<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 25th, by Michael Bernholtz - Reports shows Information from the QC_SEALANT Table and allow Labels to be Printed-->
<!-- QC_SEALANT Table created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->
<!-- One of 3 Tables - QC_Sealant, QC_SPACER, QC_SEALANT-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>QC Sealant Report</title>
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

		window.location.href = "QCSealantPrintLabels.asp?QCID="+QCID
	}

	</script>	
    <%
	
threemonth = DateAdd("m",-3,Date)
	
Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = FixSQLCheck("SELECT MSE.ItemName, MSE.Manufacturer, SE.Identifier, SE.EntryDate, SE.ConsumeDate, SE.printed, SE.Id FROM QC_SEALANT AS SE INNER JOIN QC_MASTER_SEALANT AS MSE ON MSE.id = SE.MasterID WHERE SE.ENTRYDATE > #" & threemonth & "# ORDER BY MASTERID, SE.ID ASC", isSQLServer)
rs3.Cursortype = GetDBCursorType
rs3.Locktype = GetDBLockType
rs3.Open strSQL3, DBConnection

%>
 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">QC INVENTORY REPORT</h1>
         <a class="button leftButton" type="cancel" href="index.html#_QC" target="_self">QC</a>
        </div>
   
   
         
       
        <form id="SealantPrint" action="QCSealantPrint.asp" name="SealantPrint"  method="GET" target="_self" selected="true" > 

        <ul title="QC Report - Sealant">  
	<% 
listnumber =0
response.write "<li class='group'>QC Sealant REPORT </li>"
response.write "<li><table border='1' class='sortable' width='100%'><tr><th>#</th><th>Item Name</th><th>Identifier</th><th>Manufacturer</th><th>Entry Date</th><th>Consumed Date</th><th><input type='submit' value = 'Print' onClick='SealantPrint.submit()'></input><BR></th><th>Single Print</th></tr>"

do while not rs3.eof
	listnumber = listnumber +1
	response.write "<tr><td> " & listnumber & " </td><td>" & rs3("ItemName") &"</td><td>" & rs3("Identifier") & "</td><td>" & rs3("Manufacturer") & "</td><td>" & rs3("EntryDate") & "</td><td>" & rs3("ConsumeDate") & "</td>"
	
	response.write "<td><input type ='checkbox' name = 'QCID' value = '" & trim(RS3.fields("ID")) & "' </td>"
	
	
	if rs3("Printed") = 1  then
		response.write "<td><input type ='button' value = 'Print Label' onclick='PrintLabelCheck(" & trim(RS3.fields("ID")) & ")'</td>"
	else
		response.write "<td><input type ='button' value = 'Print New Label' style='background-color:yellow' onclick='PrintLabel(" & trim(RS3.fields("ID")) & ")'</td>"
	end if
	rs3.movenext
loop

	rs3.close
	set rs3 = nothing
	DBConnection.close
	set DBConnection = nothing

%>
</table></li>
</ul>
</form>   
      
</body>
</html>
