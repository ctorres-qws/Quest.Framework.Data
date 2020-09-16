<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 25th, by Michael Bernholtz - Reports shows Information from the QC_GLASS Table and allow Labels to be Printed-->
<!-- QC_GLASS Table created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->
<!-- One of 3 Tables - QC_GLASS, QC_SPACER, QC_SEALANT-->
<!-- February 2019 - Added USA Table -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>QC Glass Report</title>
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

		window.location.href = "QCGlassPrintLabels.asp?QCID="+QCID
	}

	</script>
    <%
	
oneyear = DateAdd("yyyy",-1,Date)

Set rs = Server.CreateObject("adodb.recordset")
if CountryLocation = "USA" then
	strSQL = FixSQLCheck("SELECT MG.ItemName, MG.Manufacturer, G.SerialNumber, G.Quantity, G.EntryDate, G.ConsumeDate, G.printed, G.Id FROM QC_GLASS_USA AS G INNER JOIN QC_MASTER_GLASS AS MG ON MG.id = G.MasterID WHERE G.ENTRYDATE > #" & oneyear & "# ORDER BY MASTERID, G.ID ASC", isSQLServer)
else
	strSQL = FixSQLCheck("SELECT MG.ItemName, MG.Manufacturer, G.SerialNumber, G.Quantity, G.EntryDate, G.ConsumeDate, G.printed, G.Id FROM QC_GLASS AS G INNER JOIN QC_MASTER_GLASS AS MG ON MG.id = G.MasterID WHERE G.ENTRYDATE > #" & oneyear & "# ORDER BY MASTERID, G.ID ASC", isSQLServer)
end if
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection


%>
 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
  <div class="toolbar">
        <h1 id="pageTitle">Glass Print</h1>
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write Homesite%>#_QC" target="_self">Glass<%response.write HomeSiteSuffix%></a>
    </div>  
   
   
         
       
        <form id="GlassPrint" action="QCGlassPrint.asp" name="GlassPrint"  method="GET" target="_self" selected="true" > 

        <ul title="QC Report - Glass">
     <% 
listnumber =0
response.write "<li class='group'>QC GLASS REPORT </li>"
response.write "<li><table border='1' class='sortable' width='100%'><tr><th>#</th><th>Item Name</th><th>Serial Number</th><th>Manufacturer</th><th>Packs</th><th>Entry Date</th><th>Consumed Date</th><th><input type='submit' value = 'Print' onClick='GlassPrint.submit()'></input><BR></th><th>Single Print</th></tr>"

do while not rs.eof
	listnumber = listnumber +1
	response.write "<tr><td> " & listnumber & " </td><td>" & RS("ItemName") &"</td><td>" & RS("SerialNumber") & "</td><td>" & RS("Manufacturer") & "</td><td>" & RS("Quantity") & "</td><td>" & RS("EntryDate") & "</td><td>" & RS("ConsumeDate") & "</td>"
	
	response.write "<td><input type ='checkbox' name = 'QCID' value = '" & trim(RS.fields("ID")) & "' </td>"
	

	if rs("Printed") = 1  then
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
