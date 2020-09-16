		<!--#include file="QCdbpath.asp"--> 
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 
<!-- Water Filter Testing for Willian Line - Designed for Ruslan Bedoev - May 2015, Michael Bernholtz-->
<!-- Confirm Page - Entered from QCTest_WaterFilterEnter.asp -->

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

Level = REQUEST.QueryString("Level")
if Level = "" then
	Level = 0
end if
PassFail = REQUEST.QueryString("PassFail")
If PassFail = "on" then
	PassFail = TRUE
Else
	PassFail = FALSE
End If

Notes = REQUEST.QueryString("Notes")

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM TEST_WaterFilter"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.addnew
if isDate(InDate) then
	rs("Date") = InDate
End if
rs("Time") = InTime
rs("Level") = Level
rs("PassFail") = PassFail
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
                <a class="button leftButton" type="cancel" href="QCTest_WaterFilterEnter.asp" target="_self">Back</a>

    </div>


    
<ul id="Report" title="Added" selected="true">
	
    <li><% response.write "Date: " & InDate %></li>
	<li><% response.write "Time: " & Time %></li>
	<li><% response.write "Filter Level: " & Level %></li>
    <li><% response.write "Filter Pass?: " & PassFail %></li>
	<li><% response.write "Tested BY: " & Initials %></li>
	<li><% response.write "Notes: " & Notes %></li>

	<a class="whiteButton" href="QCTest_WaterFilter.asp" target="_self">Confirm and Return</a>
</ul>

</body>
</html>



