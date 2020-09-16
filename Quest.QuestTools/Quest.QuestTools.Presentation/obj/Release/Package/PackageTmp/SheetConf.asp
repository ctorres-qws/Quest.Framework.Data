

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">  
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="dbpath.asp"-->
<!--Sheet Inventory Entry Page enters into Y_SHEET_INV, Michael Bernholtz, March 2017 -->

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
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_SHEET_INV"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection


Location = TRIM(REQUEST.QueryString("Location"))

Material = TRIM(REQUEST.QueryString("Material"))
Qty = REQUEST.QueryString("Qty")
if Qty = "" then
QTY = 0
end if
PO = TRIM(REQUEST.QueryString("PO"))


rs.filter = "JOB = '" & Job & "' AND Thickness = '" & Thickness & "'"
Result = "None"

if rs.eof then

rs.AddNew
	rs.Fields("Job") = Job
	rs.Fields("Location") = Location
	rs.Fields("qty") = qty
	rs.Fields("EntryQty") = qty
	rs.Fields("Thickness") = Thickness
	rs.Fields("Material") = Material
	rs.Fields("PO") = PO
	rs.Fields("EntryDate") = currentDate
	rs.Fields("LastModify") = currentDate	
rs.update
Result = "New"

else
OldQty = rs("qty")
NewQty = OldQty + qty
rs("QTY") = NewQty
rs("LastModify") = currentDate
rs.update
Result = "Add"

end if

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Panel" target="_self">Sheets</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>


    
<ul id="Report" title="Sheet Entered" selected="true">

<% if Result = "New" then %>

	<li><% response.write "Job " & Job %></li>
	<li><% response.write "Material " & Material %></li>
    <li><% response.write "Qty " & Qty %></li>
    <li><% response.write "Location " & Location%></li>
    <li><% response.write "Entry Date " & currentDate %></li>
    <li><% response.write "PO " & PO %></li>
	<li><% response.write "Thickness" & Thickness %></li>
	<li><% response.write "Bundle " & Bundle %></li>
<% End if %>

<% if Result = "Add" then %>
    <li><% response.write "Job " & Job%></li>
	<li><% response.write "Thickness " & Thickness %></li>
	<li><% response.write "New Qty " & NewQty %></li>
	<li><% response.write "Modified " & currentDate %></li>
<% End if %>


<li><a class = 'whiteButton' href='SheetAdd.asp' target='_self'>Add Another Item </a></li>
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



