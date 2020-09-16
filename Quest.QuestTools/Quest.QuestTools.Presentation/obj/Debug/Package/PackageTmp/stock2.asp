<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- February 2019 Added USA option -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>iUI Theaters</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
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
        <h1 id="pageTitle">Active Stock</h1>
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
        <li class='group'>Aluminum Profiles</li>
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT yI.*, yM.Description FROM Y_INV yI LEFT JOIN Y_Master yM ON yM.Part = yI.Part ORDER BY yI.PART ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Do While Not rs.eof

	If UCASE(rs("Part")) = UCASE(part) Then
	Else
		If rs("Description") & "" = "" Then
%>
		<li><a href="stockbydie.asp?part=<%response.write rs("Part")%>" target="_self"><%response.write rs("Part") %></a></li>
<%
		Else
%>
		<li><a href="stockbydie.asp?part=<%response.write rs("Part")%>" target="_self"><%response.write rs("Part") %> (<%response.write rs("Description") %>)</a></li>
<%
		End If
	End If
	part = rs("Part")
	rs.movenext
Loop

rs.close
set rs = nothing
DBConnection.close
Set DBConnection = nothing

%>
      <li>//END//</li>
	  </ul>
/
</body>
</html>
