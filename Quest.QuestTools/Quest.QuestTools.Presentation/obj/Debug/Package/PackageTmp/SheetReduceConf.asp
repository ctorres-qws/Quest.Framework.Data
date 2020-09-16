<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">  
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="dbpath.asp"-->
<!--Sheet Inventory Reducing Page based on Scan removes from Y_SHEET_INV, Michael Bernholtz, March 2017 -->

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

currentDate = Date()
Job  = TRIM(REQUEST.QueryString("Job"))
Thickness = TRIM(REQUEST.QueryString("Thickness"))
Qty = REQUEST.QueryString("Qty")
if Qty = "" then
QTY = 0
end if
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_SHEET_INV"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

rs.filter = "JOB = '" & Job & "' AND Thickness = '" & Thickness & "'"
Error = "FALSE"
Updated = "FALSE"
if rs.eof then
	Error = "No Existing Inventory of " & JOB & " : " & Thickness
else
	if rs("QTY") + 0 < QTY + 0 then
		Error = "Cannot remove " & QTY & ", Inventory only has " & rs("qty")
	end if
end if

if Error = "FALSE" then
	OldQty = rs("qty")
	NewQty = OldQty - qty
	rs.Fields("qty") = NewQty
	rs.Fields("LastModify") = currentDate
	rs.update
	Updated = "TRUE"
end if

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="SheetReduce.asp" target="_self">Reduce</a>
    </div>

<ul id="Report" title="Sheets Reduced" selected="true">	
<% if Updated = "TRUE" then
%>
    

    <li><% response.write "Old Qty " & OldQty %></li>
    <li><% response.write "New Qty " & NewQty %></li>
    <li><% response.write "Modified Date " & currentDate %></li>
<% ELSE %>

    <li><% response.write "Error: Could not Update "%></li>
    <li><% response.write "Error " & ERROR %></li>


<% END IF %>


<li><a class = 'whiteButton' href='Index.html#_Panel' target='_self'>Home</a></li>
</ul>

<% 

rs.close
set rs=nothing
DBConnection.close
set DBConnection=nothing
%>

</body>
</html>



