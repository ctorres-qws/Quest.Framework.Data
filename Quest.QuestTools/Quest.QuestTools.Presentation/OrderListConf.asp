<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 
<!--Add new item to orderlist Conf page-->
<!-- New Report Designed for Lev Bedeov: Table Maintains all orders sent out to QuickTemp / Cardinal-->
<!-- Built by Michael Bernholtz, Maintained by Tomas, Michael Angel, Ruslan, October 2014-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
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
PO = UCASE(REQUEST.QueryString("PO"))
GlassCode = REQUEST.QueryString("GlassCode")
JOB = UCASE(REQUEST.QueryString("JOB"))
Floor = UCASE(REQUEST.QueryString("Floor"))
Qty = REQUEST.QueryString("Qty")
	if Qty = "" then
		Qty = 0
	end if
From = REQUEST.QueryString("From")
OrderBy = REQUEST.QueryString("OrderBy")
ShipOutDate = REQUEST.QueryString("ShipOutDate")
OrderDate = REQUEST.QueryString("OrderDate")
ExpectedDate = REQUEST.QueryString("ExpectedDate")
Notes = REQUEST.QueryString("Notes")

' Received / Broken
'COMPLETED = REQUEST.QueryString("COMPLETED")
'If COMPLETED = "on" then
'	COMPLETED = TRUE
'Else
'	COMPLETED = FALSE
'End If


Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM GlassOrder"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.addnew
rs("PO") = PO
rs("GlassCode") = GlassCode
rs("Job") = Job
rs("Floor") = Floor
rs("Qty") = Qty
rs("From") = From
rs("orderBy") = orderby
rs("Notes") = Notes
if isDate(ShipOutDate) then
	rs("ShipOutDate") = ShipOutDate
End if
if isDate(OrderDate) then
	rs("OrderDate") = OrderDate
End if
if isDate(ExpectedDate) then
	rs("ExpectedDate") = ExpectedDate
End if
rs("Active") = True
rs.update

%>
	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="OrderListENTER.asp" target="_self">Enter Job</a>

    </div>


    
<ul id="Report" title="Added" selected="true">
	
    <li><% response.write "PO: " & PO %></li>
	<li><% response.write "Glass: " & GlassCode %></li>
    <li><% response.write "Job: " & JOB %></li>
	<li><% response.write "Floor: " & FLOOR %></li>
    <li><% response.write "Quantity: " & QTY %></li>
    <li><% response.write "Ordered From: " & FROM %></li>
    <li><% response.write "ordered By: " & ORDERBY %></li>
	<li><% response.write "Notes: " & Notes %></li>

	<a class="whiteButton" href="OrderList.asp" target="_self">Return to Order List</a>
</ul>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection = nothing
%>

</body>
</html>



