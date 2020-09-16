<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 6th, by Michael Bernholtz - Reports shows Information from all QC Tables-->
<!-- QC_Inventory Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->
<!-- 3 Tables - QC_GLASS, QC_SPACER, QC_SEALANT Matched to the three MASTER tables QC_MASTER_GLASS, QC_MASTER_SPACER, QC_MASTER_SEALANT-->
<!-- February 2019 - USA Tables Added-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>QC Glass Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <!-- DataTables CSS -->
<link rel="stylesheet" type="text/css" href="/DataTables-1.10.0/media/css/jquery.dataTables.css">
  
<!-- jQuery -->
<script type="text/javascript" charset="utf8" src="/DataTables-1.10.0/media/js/jquery.js"></script>
  
<!-- DataTables -->
<script type="text/javascript" charset="utf8" src="/DataTables-1.10.0/media/js/jquery.dataTables.js"></script>
  
  
  
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
	</script>
	


	
    <%

InitPage(IsSQLServerDefault())

Sub InitPage(isSQLServer)


if CountryLocation = "USA" then
	strSQL = FixSQLCheck("SELECT MG.ItemName, MG.Manufacturer, MG.Pieces, MG.Height, MG.Width , MG.Price, G.SerialNumber, G.Quantity, G.EntryDate, G.ConsumeDate FROM QC_GLASS_USA AS G INNER JOIN QC_MASTER_GLASS AS MG ON MG.id = G.MasterID Where ISNULL(G.ConsumeDate) ORDER BY MASTERID ASC", isSQLServer)
	strSQL2 = FixSQLCheck("SELECT MSP.ItemName, MSP.Manufacturer, SP.Identifier, SP.EntryDate, SP.Quantity, SP.ConsumeDate FROM QC_SPACER_USA AS SP INNER JOIN QC_MASTER_SPACER AS MSP ON MSP.id = SP.MasterID WHERE ISNULL(SP.ConsumeDate) ORDER BY MASTERID ASC", isSQLServer)
	strSQL3 = FixSQLCheck("SELECT MSE.ItemName, MSE.Manufacturer, SE.Identifier, SE.EntryDate, SE.Printed , SE.ConsumeDate FROM QC_SEALANT_USA AS SE INNER JOIN QC_MASTER_SEALANT AS MSE ON MSE.id = SE.MasterID WHERE ISNULL(SE.ConsumeDate) ORDER BY MASTERID ASC", isSQLServer)
	strSQL4 = FixSQLCheck("SELECT MM.ItemName, MM.Manufacturer, M.Identifier, M.Quantity, M.EntryDate, M.ConsumeDate FROM QC_MISC_USA AS M INNER JOIN QC_MASTER_MISC AS MM ON MM.id = M.MasterID WHERE ISNULL(M.ConsumeDate) ORDER BY MASTERID ASC", isSQLServer)
else
	strSQL = FixSQLCheck("SELECT MG.ItemName, MG.Manufacturer, MG.Pieces, MG.Height, MG.Width , MG.Price, G.SerialNumber, G.Quantity, G.EntryDate, G.ConsumeDate FROM QC_GLASS AS G INNER JOIN QC_MASTER_GLASS AS MG ON MG.id = G.MasterID Where ISNULL(G.ConsumeDate) ORDER BY MASTERID ASC", isSQLServer)
	strSQL2 = FixSQLCheck("SELECT MSP.ItemName, MSP.Manufacturer, SP.Identifier, SP.EntryDate,  SP.Quantity, SP.ConsumeDate FROM QC_SPACER AS SP INNER JOIN QC_MASTER_SPACER AS MSP ON MSP.id = SP.MasterID WHERE ISNULL(SP.ConsumeDate) ORDER BY MASTERID ASC", isSQLServer)
	strSQL3 = FixSQLCheck("SELECT MSE.ItemName, MSE.Manufacturer, SE.Identifier, SE.EntryDate, SE.Printed , SE.ConsumeDate FROM QC_SEALANT AS SE INNER JOIN QC_MASTER_SEALANT AS MSE ON MSE.id = SE.MasterID WHERE ISNULL(SE.ConsumeDate) ORDER BY MASTERID ASC", isSQLServer)
	strSQL4 = FixSQLCheck("SELECT MM.ItemName, MM.Manufacturer, M.Identifier, M.Quantity, M.EntryDate, M.ConsumeDate FROM QC_MISC AS M INNER JOIN QC_MASTER_MISC AS MM ON MM.id = M.MasterID WHERE ISNULL(M.ConsumeDate) ORDER BY MASTERID ASC", isSQLServer)
end if

Set rs = Server.CreateObject("adodb.recordset")
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

Set rs3 = Server.CreateObject("adodb.recordset")
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection

Set rs4 = Server.CreateObject("adodb.recordset")
rs4.Cursortype = 2
rs4.Locktype = 3
rs4.Open strSQL4, DBConnection

%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
  <div class="toolbar">
        <h1 id="pageTitle">Glass Inventory Report</h1>
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
       
        <ul id="Profiles" title="QC Report - Glass" selected="true">
        <li><a href="QCReportActiveExcel.asp" target="_self">Download to Excel</a>
        
<% 
' QC Glass Table
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li class='group'>QC GLASS REPORT </li>"
response.write "<li><table border='1' class='sortable' width='100%'><tr><th>Item Name</th><th>Size</th><th>Serial Number</th><th>Manufacturer</th><th>Qty</th><th>Pack</th><th>Sqft</th><th>$/Sqft</th><th>Value</th><th>Entry Date</th><th>Consumed Date</th></tr>"
do while not rs.eof
	response.write "<tr><td>" & RS("ItemName") &"</td><td>" & RS("Width") & " X " & RS("Height")  & "</td><td>" & RS("SerialNumber") & "</td><td>" & RS("Manufacturer") & "</td><td>" & RS("Quantity") & "</td><td>" & RS("Pieces") & "</td><td>" & Round(RS("Width")*RS("Height")*RS("Pieces")/144,2) & "</td><td>$" & RS("Price") & "</td><td>$" & Round(RS("Width")*RS("Height")*RS("Pieces")*RS("Price")/144,2) & "</td><td>" & RS("EntryDate") & "</td><td>" & RS("ConsumeDate") & "</td></tr>"
rs.movenext
loop


response.write "</table></li>"
' QC Spacer Table
response.write "<li class='group'>QC Spacer REPORT </li>"
response.write "<li><table border='1' class='sortable' width='75%'><tr><th  width='25%'>Item Name</th><th  width='25%'>Identifier</th><th  width='25%'>Manufacturer</th><th  width='2.5%'>Quantity</th><th  width='10%'>Entry Date</th><th  width='12.5%'>Consumed Date</th></tr>"
do while not rs2.eof
	response.write "<tr><td>" & rs2("ItemName") &"</td><td>" & rs2("Identifier") & "</td><td>" & rs2("Manufacturer") & "</td><td>" & rs2("Quantity") & "</td><td>" & rs2("EntryDate") & "</td><td>" & rs2("ConsumeDate") & "</td></tr>"
rs2.movenext
loop

response.write "</table></li>"

' QC Sealant Table
response.write "<li class='group'>QC Sealant REPORT </li>"
response.write "<li><table border='1' class='sortable' width='75%'><tr><th  width='25%'>Item Name</th><th  width='25%'>Identifier</th><th  width='25%'>Manufacturer</th><th  width='12.5%'>Entry Date</th><th  width='12.5%'>Consumed Date</th></tr>"
do while not rs3.eof
	response.write "<tr><td>" & rs3("ItemName") &"</td><td>" & rs3("Identifier") & "</td><td>" & rs3("Manufacturer") & "</td><td>" & rs3("EntryDate") & "</td><td>" & rs3("ConsumeDate") & "</td></tr>"
rs3.movenext
loop

response.write "</table></li>"

' QC Miscellaneous Table
response.write "<li class='group'>QC Miscellaneous REPORT </li>"
response.write "<li><table border='1' class='sortable' width='75%'><tr><th  width='30%'>Item Name</th><th  width='20%'>Identifier</th><th  width='20%'>Manufacturer</th><th  width='10%'>Quantity</th><th  width='8%'>Entry Date</th><th  width='12%'>Consumed Date</th></tr>"
do while not rs4.eof
	response.write "<tr><td>" & rs4("ItemName") &"</td><td>" & rs4("Identifier") & "</td><td>" & rs4("Manufacturer") & "</td><td>" & rs4("Quantity") & "</td><td>" & rs4("EntryDate") & "</td><td>" & rs4("ConsumeDate") & "</td></tr>"
rs4.movenext
loop

response.write "</table></li>"

rs.close
set rs = nothing
rs2.close
set rs2 = nothing
rs3.close
set rs3 = nothing
rs4.close
set rs4 = nothing
DBConnection.close
Set DBConnection = nothing

End Sub
%>
               
</ul>            
</body>
</html>
