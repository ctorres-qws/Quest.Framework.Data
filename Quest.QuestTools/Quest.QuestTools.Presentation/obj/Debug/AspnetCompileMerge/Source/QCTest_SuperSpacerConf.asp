		<!--#include file="QCdbpath.asp"--> 
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 
<!-- Testing Results stored in the system - Designed for Victor Babuskins - November 2014, Michael Bernholtz-->
<!-- Confirm Page - Entered from QCTest_SuperSpacerEnter.asp -->

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
InDate = REQUEST.QueryString("Date")
InTime = REQUEST.QueryString("Time")
Condition = REQUEST.QueryString("Condition")
Width = REQUEST.QueryString("Width")
if Width = "" then
	Width = 0
end if
Adhesion = REQUEST.QueryString("Adhesion")
If Adhesion = "on" then
	Adhesion = TRUE
Else
	Adhesion = FALSE
End If
Initials = REQUEST.QueryString("Initials")

Notes = REQUEST.QueryString("Notes")

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM TEST_SuperSpacer"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.addnew
if isDate(InDate) then
	rs("Date") = InDate
End if
rs("Time") = InTime
rs("Condition") = Condition
rs("Width") = Width
rs("Adhesion") = Adhesion
rs("Initials") = Initials
rs("Notes") = Notes

rs.update
rs.close
set rs=nothing

DBConnection.close
set DBConnection = nothing
%>
	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="QCTest_SuperSpacerENTER.asp" target="_self">Back</a>

    </div>


    
<ul id="Report" title="Added" selected="true">
	
    <li><% response.write "Date: " & InDate %></li>
	<li><% response.write "Time: " & InTime %></li>
	<li><% response.write "Condition: " & Condition %></li>
	<li><% response.write "Width: " & Width %></li>
    <li><% response.write "Adhesion Pass?: " & Adhesion %></li>
	<li><% response.write "Initials: " & Initials %></li>
	<li><% response.write "Notes: " & Notes %></li>

	<a class="whiteButton" href="QCTest_SuperSpacer.asp" target="_self">Confirm and Return</a>
</ul>

</body>
</html>



