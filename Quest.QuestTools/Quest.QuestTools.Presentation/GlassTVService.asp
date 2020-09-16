<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Service Glass Report - GT Codes found in Forel and Willian Scans - With Completion of PO included -->
<!-- Michael Bernholtz June 2018 -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Service Glass</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
<% Server.ScriptTimeout = 500 %> 


</head>
<body>
<!--#include file="todayandyesterday.asp"-->

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>

<ul id="screen1" title="Service Glass Status" selected="true">

		<li class="group">Service Glass Number</li>
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select  Top 500 * FROM X_BARCODEGA WHERE LEFT(BARCODE,2) = 'GT' AND ( Month = " & cMonth & " AND YEAR = " & cYear & ")   ORDER BY ID DESC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "Select Top 10000 [Barcode], ID, PO, JOB, FLOOR, TAG FROM Z_GLASSDB ORDER BY ID DESC"
rs2.Cursortype = GetDBCursorType
rs2.Locktype = GetDBLockType
rs2.Open strSQL2, DBConnection
rs.filter = "DAY = " & cDay & " AND Month = " & cMonth & " AND YEAR = " & cYear		
		
Response.write "<li class='group'>Today</li>"	
if rs.eof then
	response.write "<li> There are currently no items of Activity Today, Please check later</li>"
else
	
	Do While not rs.eof
		rs2.filter = "Barcode = '" & rs("Barcode") & "'"
		response.write "<li>" & rs("Barcode") & " : " & rs2("PO") & " :: " & rs2("JOB") & rs2("FLOOR") & "-" & rs2("TAG") & "</li>"
	
	
	rs2.filter = ""
	rs.movenext
	loop
	
	

end if

rs.filter = "Year = " & cYear & " AND Week = " & weekNumber 
	
Response.write "<li class='group'>This Week</li>"	
if rs.eof then
	response.write "<li> There are currently no items of Activity Today, Please check later</li>"
else
	
	Do While not rs.eof
		rs2.filter = "Barcode = '" & rs("Barcode") & "'"
		if not rs2.eof then
			response.write "<li>" & rs("Barcode") & " : " & rs2("PO") & " :: " & rs2("JOB") & rs2("FLOOR") & "-" & rs2("TAG") & "</li>"
		else
			response.write "<li>" & rs("Barcode") & " : Not found in Z_GLASSDB"
		end if
	rs2.filter = ""
	rs.movenext
	loop
	
	

end if

	
rs.filter = " Month = " & cMonth & " AND YEAR = " & cYear	
Response.write "<li class='group'>This Month</li>"		
if rs.eof then
	response.write "<li> There are currently no items of Activity Today, Please check later</li>"
else
	
	Do While not rs.eof
		rs2.filter = "Barcode = '" & rs("Barcode") & "'"
		response.write "<li>" & rs("Barcode") & " : " & rs2("PO") & " :: " & rs2("JOB") & rs2("FLOOR") & "-" & rs2("TAG") & "</li>"
	
	
	rs2.filter = ""
	rs.movenext
	loop
	

end if
%>

	</ul>
        
  
<% 
On Error Resume Next

rs.close
set rs=nothing
rs2.CLOSE
Set rs2= nothing
DBConnection.close
set DBConnection=nothing

%>


</body>
</html>
