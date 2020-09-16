<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 <!--#include file="dbpath.asp"-->
		 
<!--Edit Form Conf Page to update Orders-->
<!-- New Report Designed for Lev Bedeov: Table Maintains all orders sent out to QuickTemp / Cardinal-->
<!-- Built by Michael Bernholtz, Maintained by Tomas, Michael Angel, Ruslan, October 2014-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Order Update Confirmation</title>
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

ORID = request.querystring("ORID")
%>

	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
                <a class="button leftButton" type="cancel" href="OrderListEditForm.asp?ORID=<% response.write ORid %>" target="_self">Edit Job</a>
    </div>
    
      
    
<form id="conf" title="Order Updated" class="panel" name="conf" action="OrderList.asp" method="GET" target="_self" selected="true" >              

  
   
        <h2>Order Updated</h2>
  
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

Received = REQUEST.QueryString("Received")
	If Received = "on" then
		Received = TRUE
	Else
		Received = FALSE
	End If
Ack = REQUEST.QueryString("Ack")
	If Ack = "on" then
		Ack = TRUE
	Else
		Ack = FALSE
	End If
Broken = REQUEST.QueryString("Broken")
	If Broken = "on" then
		Broken = TRUE
	Else
		Broken = FALSE
	End If
ReturnB = REQUEST.QueryString("Return")
	if ReturnB = "" then
		ReturnB = 0
	end if		   

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM GlassOrder where OrderID = " & ORID
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
if rs.eof then
else
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
	rs("Broken") = Broken
	rs("Return") = ReturnB
	rs("Ack") = Ack
	rs("Received") = Received
	If rs("Received") = True then
		rs("Active") = False
		rs("ReceivedDate") = Date()
	End if
	rs.update
end if
			
	

Rs.close
Set RS = Nothing
DBConnection.close
set DBConnection=nothing


%>

    
<ul id="Report" title="Added" selected="true">

	<li>Job Summary Edited:</li>
	
    <li><% response.write "PO: " & PO %></li>
	<li><% response.write "Glass: " & GlassCode %></li>
    <li><% response.write "Job: " & JOB %></li>
	<li><% response.write "Floor: " & FLOOR %></li>
    <li><% response.write "Quantity: " & QTY %></li>
    <li><% response.write "Ordered From: " & FROM %></li>
    <li><% response.write "ordered By: " & ORDERBY %></li>
	<li><% response.write "Notes: " & Notes %></li>
	<hr>
	<li><% response.write "Broken: " & Broken %></li>
	<li><% response.write "Acknowledged: " & Acknowledged %></li>
	<li><% response.write "Received: " & Received %></li>
	
</ul>
        <BR>
       
         <a class="whiteButton" href="javascript:conf.submit()">Home</a>
            
            </form>

            
    
</body>
</html>


