<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

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
<%

If CountryLocation = "USA" then
	strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'JUPITER' ORDER BY AISLE, RACK, SHELF DESC"
else
	strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'GOREWAY' ORDER BY AISLE, RACK, SHELF DESC"
end if

Set rs = Server.CreateObject("adodb.recordset")

rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection
%>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Stock by Aisle</h1>
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

<%
aisle = "0"
do while not rs.eof
part = rs("part")
qty = rs("qty")
id = rs("ID")
shelf = rs("shelf")
	if aisle = rs("aisle") then
	else
	response.write "<a href='stockbyrack2.asp?aisle=" & rs("aisle") & "' target='_self'><li>Aisle " & rs("aisle") & "</a></li>"
	end if

aisle = rs("aisle")
rack = rs("Rack")
rs.movenext
loop

rs.close
set rs = nothing
DBConnection.close
Set DBConnection = nothing

%>
      </ul>
</body>
</html>
