<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- Created May 5th, by Michael Bernholtz at Request of Ariel Aziza -->
<!-- Stock by Colour list -->
<!-- Drills to Stockbycolor / stockbycolortable-->
<!-- February 2019 - Add USA VIEW -->
<!--#include file="dbpath.asp"-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock By Colour</title>
  <meta name="viewport" content="width=devicewidth, initial-scale=1.0, maximum-scale=1.0, user-scalable=0"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    </head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle">Stock by Color</h1>
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write Homesite%>#_Inv" target="_self">Inventory<%response.write HomeSiteSuffix%></a>
    </div>
        <ul id="Profiles" title="Profiles" selected="true">
        <li class='group'>Stock Colours</li>
		
<%
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT yC.*, zJ.Job_Name FROM Y_Color yC LEFT Join z_Jobs zJ on zJ.Job = yC.Job WHERE yC.ACTIVE = True AND (yC.EXtrusion = TRUE OR yC.SHEET = TRUE) ORDER BY yC.PROJECT ASC")
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Do While Not rs.eof
	colour = rs("PROJECT")
	activejob = rs("Job")
	code = rs("Code")
	Jobname = rs("Job_Name")
	If colour2 = colour Then
'
	Else
	
		%>
		<li>
		<%
		if rs("Extrusion") = FALSE and rs("Sheet") then
		response.write "Sheet Colour only: "
		end if 
		%>
		<a href="stockbycolorTable.asp?colour=<%response.write colour%>" target="_self"><%response.write colour & " : " & Jobname & " : " & code %></a>
		</li>
		<%
	End If
	colour2 = colour
	rs.movenext
Loop

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing

%>
<li>//END//</li>
/
      </ul>

</body>
</html>
