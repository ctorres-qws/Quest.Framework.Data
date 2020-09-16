<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 <!--#include file="dbpath.asp"-->
		 
<!--Edit Form to update Orders-->
<!-- New Report Designed for Lev Bedeov: Table Maintains all orders sent out to QuickTemp / Cardinal-->
<!-- Built by Michael Bernholtz, Maintained by Tomas, Michael Angel, Ruslan, October 2014-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Manage/Update Orders</title>
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
ORID = Request.QueryString("ORid")


		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM GlassOrder WHERE OrderID = " & ORid
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection	

%>
	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle">Manage Order Details</h1>
                <a class="button leftButton" type="cancel" href="OrderList.asp" target="_self">Orders</a>

    </div>			
    
    
    <form id="OrderEdit" title="Manage Orders" class="panel" action="orderListEditConf.asp" name="OrderEdit"  method="GET" target="_self" selected="true" > 
  
	<fieldset>
	
	<div class="row">
		<label>PO </label>
		<input type="text" name='PO' id='PO' value ='<% response.write Trim(rs.fields("PO")) %>' >
	</div>

	<div class="row">
		<label>Glass </label>
		<select name="GlassCode">
			<% mat = mat1 %>
			<!--#include file="QSU.inc"-->
			<% 
			' Coded this 3 times - to show Description again, despite collected value is TYPE ( USER CHOOSES DESCRIPTION, but SYSTEM NEEDS TYPE)
			rs5.filter = "Type = '" & RS("GlassCode") & "'"
			if rs5.eof then
			else
			%>
			<option value = "<% response.write rs5("TYPE") %>" selected><% response.write rs5("DESCRIPTION") %></option> 
			<%
			end if
			%>
			</select>
    </div>

    <div class="row">
		<label>Job </label>
		<select name="Job">
					<% ActiveOnly = True %>
			<!--#include file="JobsList.inc"-->
			<% 
			' Coded this 3 times - to show Description again, despite collected value is TYPE ( USER CHOOSES DESCRIPTION, but SYSTEM NEEDS TYPE)
			rsJob.filter = "Job = '" & rs("Job") & "'"
			if rsJob.eof then
			%><option value = "" selected>-</option><%
			else
			%>
			<option value = "<% response.write rsJob("Job") %>" selected><% response.write rsJob("Job") %></option> 
			<%
			end if
			%>
			</select>
    </div>
	<div class="row">
		<label>Floor</label>
		<input type="text" name='Floor' id='Floor' value ='<% response.write Trim(rs.fields("Floor")) %>' >
    </div>

	<div class="row">
		<label>Quantity</label>
		<input type="number" name='Qty' id='Qty' value ='<% response.write Trim(rs.fields("QTY")) %>' >
    </div>
	
	<div class="row">
		<label>From</label>
        <select name= 'From' id = 'From'>
			<option selected='selected' value="<% response.write Trim(rs.fields("From")) %>"> <% response.write Trim(rs.fields("From")) %> </option>
			<option value="QuickTemp">QuickTemp</option>
			<option value="Cardinal">Cardinal</option>
			<option value="Woodbridge">Woodbridge</option>
			<option value="Saand">Saand</option>
			<option value="TruLite">TruLite</option>
		</select>
     </div>	
	
	<div class="row">
		<label>Ordered By</label>
		<input type="text" name='orderBy' id='orderBy' value ='<% response.write Trim(rs.fields("orderBy")) %>' >
    </div>
	
	<div class="row">
		<label>Notes </label>
		<input type="text" name='Notes' id='Notes' value ='<% response.write Trim(rs.fields("Notes")) %>' >
    </div>

	<div class="row">
		<label>Ship QT Date </label>
		<% 
		
		ShipOutDate = rs.fields("ShipOutDate")
		if isDate(ShipOutDate) then
			InputDate = rs.fields("ShipOutDate")
			mm = Month(InputDate)
			 dd = Day(InputDate)
			 yy = Year(InputDate)
			 IF len(mm) = 1 THEN
			   mm = "0" & mm
			 END IF
			 IF len(dd) = 1 THEN
			   dd = "0" & dd
			 END IF
			 ShipDate = yy & "-" & mm & "-" & dd 
		else
		ShipDate = NULL
		end if
		%>
            <input type="date" name='ShipOutDate' id='ShipOutDate' value='<% response.write ShipDate %>' >
	</div>
	<div class="row">
		<label>Order Date </label>
		<% 
		OrderDate = rs.fields("orderDate")
		if isDate(orderDate) then
			InputDate = rs.fields("OrderDate")
			mm = Month(InputDate)
			dd = Day(InputDate)
			yy = Year(InputDate)
			IF len(mm) = 1 THEN
			  mm = "0" & mm
			END IF
			IF len(dd) = 1 THEN
			  dd = "0" & dd
			END IF
			orDate = yy & "-" & mm & "-" & dd 
		else
		orDate = NULL
		end if
		%>
            <input type="date" name='OrderDate' id='OrderDate' value='<% response.write orDate %>' >
	</div>
	<div class="row">
		<label>Expected Date </label>
		<% 
		ExpectedDate = rs.fields("ExpectedDate")
		if isDate(ExpectedDate) then
					InputDate = rs.fields("ExpectedDate")
			mm = Month(InputDate)
			dd = Day(InputDate)
			yy = Year(InputDate)
			IF len(mm) = 1 THEN
			  mm = "0" & mm
			END IF
			IF len(dd) = 1 THEN
			  dd = "0" & dd
			END IF
			ExDate = yy & "-" & mm & "-" & dd 
		else
		ExDate = NULL
		end if
		%>
            <input type="date" name='ExpectedDate' id='ExpectedDate' value='<% response.write ExDate %>' >
	</div>
    <div class="row">
		<label>Acknowledged</label>
        <input type="checkbox" name='Ack' id='Ack' <% if rs.fields("Ack") = TRUE THEN response.write "checked" END IF%>>
    </div>  
    <div class="row">
		<label>Received</label>
        <input type="checkbox" name='Received' id='Received' <% if rs.fields("Received") = TRUE THEN response.write "checked" END IF%>>
    </div>     
	<div class="row">
		<label>Broken</label>
        <input type="checkbox" name='Broken' id='Broken' <% if rs.fields("Broken") = TRUE THEN response.write "checked" END IF%>>
    </div>  
	<div class="row">
		<label># Broken</label>
        <input type="Number" name='ReturnB' id='ReturnB' value='<% response.write rs.fields("Return") %>'>
    </div> 
	
						<input type="hidden" name='ORid' id='ORid' value="<%response.write ORid %>" />
</fieldset>


        <BR>
        
		
		<a class="whiteButton" onClick="OrderEdit.action='OrderListEditConf.asp'; OrderEdit.submit()">Submit Changes</a><BR>
		
            
            </form>
                        
    
</body>
</html>

<% 

rs.close
set rs=nothing
rs5.close
set rs5=nothing
rsJob.close
set rsJob=nothing
DBConnection.close
set DBConnection=nothing
%>

