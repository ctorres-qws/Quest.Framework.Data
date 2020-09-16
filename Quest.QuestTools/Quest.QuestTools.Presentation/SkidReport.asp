<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created by Michael Bernholtz March 2014 - Skid system to note items coming in and out of the warehouse-->
<!-- Skid Report displays all the items on all of the skids -->
<!-- Skid Report displays a Flush button that goes red after 3 days and uses a report flag to remove and then return -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Skid Views</title>
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

' Collect the Skids and all the items/barcodes on each skid

currentDate = Date()
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM SKIDItem ORDER BY name ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%> 
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
</head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Skid REPORT</h1>
        <a class="button leftButton" type="cancel" href="Index.Html#_Skid" target="_self">Skids</a>
        </div>
   
        <ul id="Profiles" title="Skids and Skid Items" selected="true">

<%

	response.write "<li>All Items on the Skids</li>"
	response.write "<li><table border='1' class='sortable' width='75%'><tr><th  width='10%'>SKID</th><th width='30%'>Barcode</th><th width='10%'>Job</th><th width='10%'>Floor</th><th width='10%'>Tag</th><th  width='10%'>Scan Date</th><th  colspan='2' width='20%'>Flushed</th></tr>"
if rs.eof then
		response.write "<tr><td colspan = '7'>No items on any skids</td></tr>"
else
	rs.movefirst
	Do while not rs.eof
			response.write "<tr><td>" & rs("name") & "</td><td>" & rs("Barcode") & "</td><td>" & rs("Job") & "</td><td>" & rs("Floor") & "</td><td>" & rs("Tag") & "</td><td>" & rs("ScanDate") & "</td><td>" & rs("FlushedDate") & "</td>"
			
			'Flush Button Turns red if Over three day - Item Unflush if already Flushed
			
			if rs("flushed") = 0 then
				if datediff("d" , rs("ScanDate"), currentDate) > 2 then
					response.write "<td><a class='redButton' target='#_self' href='skidremove.asp?report=1&barcodeid=" & trim(rs("Barcode")) & "'>Flush</a></td>"
				else
					response.write "<td><a class='whiteButton' target='#_self' href='skidremove.asp?report=1&barcodeid=" & trim(rs("Barcode")) & "'>Flush</a></td>"
				end if
			else
				response.write "<td><a class='grayButton' target='#_self' href='skidunflush.asp?report=1&barcodeid=" & trim(rs("Barcode")) & "'>Undo Flush</a></td>"
			end if
	
	rs.movenext
	loop
end if
	response.write "</table></li>"

rs.close
set rs = nothing

DBConnection.close
set DBConnection = nothing

%>


</body>
</html>	