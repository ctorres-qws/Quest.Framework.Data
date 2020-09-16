<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Glass Timeline is a new page designed January 2015 -->
<!-- Opens in a new window with Date information for a record -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Production Glass Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
<style>
table{
zoom: 70%;
};
 </style>
    <%
	
GID = request.querystring("GID")
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Z_GLASSDB WHERE [ID] = " & GID 
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM X_BARCODE where JOB = '"& RS("Job") & "' and Floor = '"& RS("Floor") & "'"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection
rs2.filter = "Tag = '-"& RS("Tag") & "'"

'afilter = request.QueryString("aisle")


%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        </div>
   
   
         
       
        <ul id="Profiles" title="Glass Report - Production" selected="true">
        
        
<% 
response.write "<li class='group'>Time Line of " & RS("ID") & ": " & RS("JOB") & RS("FLoor") & "-" & RS("Tag") & " </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li> Close this tab to continue working</li>  "
response.write "<li><table border = '1'> <tr><th>Original Assembled</th><th>Original Glazed</th><th>Required Date</th><th>Ordered</th><th>Optima</th><th>Expected Exterior</th><th>Cut/Received Exterior</th><th>Expected Interior</th><th>Cut/Received Interior</th><th>Service Sealed</th><th>Shipped</th><tr>"

do while not rs.eof
	response.write "<tr>"
	
	if not rs2.eof then
		rs2.filter = "DEPT = 'ASSEMBLY'"
		if not rs2.eof then
			response.write "<td>" & RS2("DateTime") & "</td>"
		else 
			response.write "<td></td>"
		end if
		rs2.filter = "DEPT = 'GLAZING'"
		if not rs2.eof then
			response.write "<td>" & RS2("DateTime") & "</td>"
		else 
			response.write "<td></td>"
		end if
	else
		response.write "<td></td>"
		response.write "<td></td>"
	end if
	
	
	response.write "<td>" & RS("REQUIREDDATE") & "</td>"
	response.write" <td>" & RS("INPUTDATE") & "</td>"
	response.write "<td>" & RS("OPTIMADATE") & "</td>"
	response.write "<td>" & RS("ExtExpected") & "</td>"
	response.write "<td>" & RS("ExtReceived") & "</td>"
	response.write "<td>" & RS("IntExpected") & "</td>"
	response.write "<td>" & RS("IntReceived") & "</td>"
	response.write "<td>" & RS("COMPLETEDDATE") & "</td>"


	
	response.write "<td>" & RS("SHIPDATE") & "</td>"
	response.write " </tr>"
	response.write "<tr>"
	if not rs2.eof then
		rs2.filter = "DEPT = 'ASSEMBLY'"
		if not rs2.eof then
			response.write "<td>" & RS2("EMPLOYEE") & "</td>"
		else 
			response.write "<td></td>"
		end if
		rs2.filter = "DEPT = 'GLAZING'"
		if not rs2.eof then
			response.write "<td>" & RS2("EMPLOYEE") & "</td>"
		else 
			response.write "<td></td>"
		end if
	else
		response.write "<td></td>"
		response.write "<td></td>"
	end if
	
	response.write "<td>" & RS("PO") & "</td>"
	response.write" <td>" & RS("ORDERBY") & "</td>"
	response.write "<td>" & RS("QTFILE") & "</td>"
	response.write "<td>" & RS("ExtOrderNum") & "</td>"
	response.write "<td>" & RS("ExtFrom") & "</td>"
	response.write "<td>" & RS("IntOrderNum") & "</td>"
	response.write "<td>" & RS("IntFrom") & "</td>"
	response.write "<td>" & RS("Barcode") & "</td>"

	response.write "<td></td>"
	response.write"</tr>"
	
	rs.movenext
loop
response.write "</table></li>"
response.write "<br><br> <li><h2> Old Timeline </h2><li>"
response.write "<li><table border='1' class='sortable'><tr><th>Added to System</th><th>Required Completion Date</th><th>Sent to Optima</th><th>Completed Date</th><th>Shipped Date</th><th>Ordered from Cardinal</th><th>Expected From Cardinal</th><th>Received from Cardinal</th><th>ordered from QuickTemp</th><th>Received From QuickTemp</th><th>QT File Name</th><th>PO</th></tr>"
rs.movefirst
do while not rs.eof
	response.write "<tr>"
	response.write" <td>" & RS("INPUTDATE") & "</td>"
	response.write "<td>" & RS("REQUIREDDATE") & "</td>"
	response.write "<td>" & RS("OPTIMADATE") & "</td>"
	response.write "<td>" & RS("COMPLETEDDATE") & "</td>"
	response.write "<td>" & RS("SHIPDATE") & "</td>"
	response.write "<td>" & RS("CARDINALSENT") & "</td>"
	response.write "<td>" & RS("CARDINALEXPECTED") & "</td>"
	response.write "<td>" & RS("CARDINALRECEIVED") & "</td>"
	response.write "<td>" & RS("QUICKTEMPSENT") & "</td>"
	response.write "<td>" & RS("QUICKTEMPRECEIVED") & "</td>"
	response.write "<td>" & RS("QTFILE") & "</td>"
	response.write "<td>" & RS("PO") & "</td>"
	response.write " </tr>"
	rs.movenext
loop
response.write "</table></li>"


rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close 
set DBConnection = nothing

%>
               
    </ul>        
            
            
       
            
              
               
                
             
               
</body>
</html>
