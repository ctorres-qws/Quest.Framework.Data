                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<!-- New Report Designed for Lev Bedeov: Table Maintains all orders sent out to QuickTemp / Cardinal-->
<!-- Built by Michael Bernholtz, Maintained by Tomas, Michael Angel, Ruslan, October 2014-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Order Glass</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
	 
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

<!-- DataTables CSS -->
<link rel="stylesheet" type="text/css" href="../DataTables-1.10.2/media/css/jquery.dataTables.css">
<!-- jQuery -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.js"></script>
<!-- DataTables -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.dataTables.js"></script>
<script type="text/javascript">
  $(document).ready( function () {
    $('#GlassOrder').DataTable();
} );
  
  </script>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>
   
   <%
   
Added = False


PO = REQUEST.QueryString("PO")
GlassCode = REQUEST.QueryString("GlassCode")
if GlassCode = "" then
GlassCode = " - "
end if
JOB = REQUEST.QueryString("JOB")
FLOOR = REQUEST.QueryString("FLOOR")
QTY = REQUEST.QueryString("QTY")
if QTY = "" then
QTY = 0
end if
From= REQUEST.QueryString("From")
OrderBy = REQUEST.QueryString("OrderBy")
Notes = REQUEST.QueryString("Notes")

ShipOutDate = REQUEST.QueryString("ShipOutDate")
OrderDate = REQUEST.QueryString("OrderDate")
ExpectedDate = REQUEST.QueryString("ExpectedDate")


if NOT JOB = "" then

Set rs = Server.CreateObject("adodb.recordset")
'strSQL = "INSERT INTO GlassOrder (PO, GlassCode, Job, Floor, Notes, Qty, From, OrderBy, Active) VALUES ('" & PO & "', '" & GlassCode & "', '" & Job & "', '" & Floor & "', '" & Notes & "', " & QTY & ", '" & FROM & "', '" & ORDERBY & "', TRUE)"

strSQL = "INSERT INTO GlassOrder (PO, GlassCode, Job, Floor, Notes, Qty, From, OrderBy, Active) VALUES ('" & PO & "', '" & GlassCode & "', '" & Job & "', '" & Floor & "', '" & Notes & "', " & Qty & ", '" & From & "', '" & OrderBy & "', True)"

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

' -------------------------------------------------------ShipOutDate ---------------------------------------		
		if isDate(ShipOutDate) then		
			
			StrSQL2 = "UPDATE GlassOrder SET ShipOutDate='" & ShipOutDate & "' WHERE ID = " & OPID
		'Get a Record Set
				Set RS2 = DBConnection.Execute(strSQL2)
		end if		
		
		if ShipOutDate = "" then		
			StrSQL2 = "UPDATE GlassOrder SET ShipOutDate= NULL WHERE ID = " & OPID
		'Get a Record Set
				Set RS2 = DBConnection.Execute(strSQL2)
			
		end if
' -------------------------------------------------------OrderDate ---------------------------------------		
		if isDate(OrderDate) then		
			
			StrSQL2 = "UPDATE GlassOrder SET OrderDate='" & OrderDate & "' WHERE ID = " & OPID
		'Get a Record Set
				Set RS2 = DBConnection.Execute(strSQL2)
		end if		
		
		if OrderDate = "" then		
			StrSQL2 = "UPDATE GlassOrder SET OrderDate= NULL WHERE ID = " & OPID
		'Get a Record Set
				Set RS2 = DBConnection.Execute(strSQL2)
			
		end if
' -------------------------------------------------------ExpectedDate ---------------------------------------		
		if isDate(ExpectedDate) then		
			
			StrSQL2 = "UPDATE GlassOrder SET ExpectedDate='" & ExpectedDate & "' WHERE ID = " & OPID
		'Get a Record Set
				Set RS2 = DBConnection.Execute(strSQL2)
		end if		
		
		if ExpectedDate = "" then		
			StrSQL2 = "UPDATE GlassOrder SET ExpectedDate= NULL WHERE ID = " & OPID
		'Get a Record Set
				Set RS2 = DBConnection.Execute(strSQL2)
			
		end if		


Added= True 
else

Added = False

End if

   
   %>
   
   
            
            <form id="enter" title="Enter New Glass Order" class="panel" name="enter" action="glassordertable.asp" method="GET" target="_self" selected="true">
              
                              
        <h2>Enter Glass Order Information:</h2>
		
		 <ul id="Profiles" title="Enter Glass Order" selected="true">
		<li><table border='1' class="GlassOrder"> 
		 <tr><th>PO</th><th>Glass Code</th><th>Job</th><th>Floor</th>
		 <th>Quantity</th><th>QuickTemp/Cardinal</th><th>OrderBy</th>
		 </tr>

		<tr>
		<td><input class="NoMargin" type="text" name='PO' id='PO' size='10' value = "<% response.write PO %>" ></td>
		<td><select name="GlassCode">
		
			<% mat = mat1 %>
			<!--#include file="QSU.inc"-->
			<% 
			' Coded this 3 times - to show Description again, despite collected value is TYPE ( USER CHOOSES DESCRIPTION, but SYSTEM NEEDS TYPE)
			rs5.filter = "Type = '" & GlassCode & "'"
			if rs5.eof then
			else
			%>
			<option value = "<% response.write rs5("TYPE") %>" selected><% response.write rs5("DESCRIPTION") %></option> 
			<%
			end if
			%>
			</select></td>
			<td><select name="Job">
					<% ActiveOnly = True %>
			<!--#include file="JobsList.inc"-->
			<% 
			' Coded this 3 times - to show Description again, despite collected value is TYPE ( USER CHOOSES DESCRIPTION, but SYSTEM NEEDS TYPE)
			rsJob.filter = "Job = '" & Job & "'"
			if rsJob.eof then
			else
			%>
			<option value = "<% response.write rsJob("Job") %>" selected><% response.write rsJob("Job") %></option> 
			<%
			end if
			%>
			</select></td>
		<td><input class="NoMargin" type="text" name='FLOOR' id='FLOOR' size='10' value = "<% response.write FLOOR %>" ></td>
		<td><input class="NoMargin"  type="number" name='Qty' id='Qty' size='10' value = "<% response.write Qty %>" ></td>
		<td><select name= 'FROM' id = 'FROM'>
			<option value = "<% response.write FROM %>" ><% response.write FROM %></option> 
			<option value="QuickTemp">QuickTemp</option>
			<option value="Cardinal">Cardinal</option>
			</select></td>
		<td><input class="NoMargin"  type="text" name='OrderBy' id='OrderBy' size='10' value = "<% response.write OrderBy %>" ></td>
		</tr>
		</table><li>
		<li><table border='1' class> 
		<th>QuickTemp Ship Date</th><th>Order Date</th><th>Expected Receive Date</th><th>Notes</th>
		<tr>
		<td>
		<% 
		if isDate(ShipOutDate) then
		ShipDate = ShipOutDate
		else
		ShipDate = NULL
		end if
		%>
            <input type="date" class="NoMargin" name='ShipOutDate' id='ShipOutDate' size='10' value='<% response.write ShipDate %>' ></td>
		<td>
		<% 
		if isDate(OrderDate) then
		OrDate = OrderDate
		else
		OrDate = NULL
		end if
		%>
            <input type="date" class="NoMargin" name='OrderDate' id='OrderDate' size='10' value='<% response.write OrDate %>' ></td>
		<td>
		<% 
		if isDate(ExpectedDate) then
		ExDate = ExpectedDate
		else
		ExDate = NULL
		end if
		%>
            <input type="date" class="NoMargin" name='ExpectedDate' id='ExpectedDate' size='10' value='<% response.write ExDate %>' ></td>
				
			
	
		<td><input  class="NoMargin" type="text" name='Notes' id='Notes' size='30'  value = "<% response.write Notes %>" ></td>
		
		</table><li>
        <br>    
         <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
        <br>
<%



'Already above
'PoNum = Request.querystring("PoNum")
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * from GlassOrder where [Active] = True"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

		response.write "<li class='group'>All Active Glass Orders</li>"
		response.write "<li><table border='1' class='Glassorder' id='Glassorder'><thead><tr><th>PO</th><th>Glass Code</th><th>Job</th><th>Floor</th><th>Quantity</th><th>QuickTemp/Cardinal</th><th>Order By</th><th>Ship to QuickTemp</th><th>Order Date</th><th>Expected Date</th><th>Notes</th><th>Department</th><th>Order</th><th>PO</th><th>Gas</th><th>Notes</th></tr></thead><tbody>"

if rs.eof then
Response.write "<tr><td>No current orders</td></tr>"
end if		
do while not rs.eof
	response.write "<tr><td>" & RS("PO") & "</td><td>" & RS("GlassCode") & "</td><td>" & RS("Job") &"</td><td>" & RS("Floor") & "</td><td>" & RS("QTY") & "''</td><td>" & RS("From") & "''</td>" 
	response.write "<td>" & RS("OrderBy") & "</td><td>" & RS("ShipoutDate") & "</td><td>" & RS("OrderDate") & "</td><td>" & RS("ExpectedDate") & "</td><td>" & RS("Notes") & "</td></tr>"
	
	rs.movenext
loop
response.write "</tbody></table></li>"


rs.close
set rs = nothing
rs5.close
set rs5 = nothing
rsJob.close
set rsJob = nothing
DBConnection.close 
set DBConnection = nothing
%>
<br>
</ul>
            </form>
                
             
               
</body>
</html>
