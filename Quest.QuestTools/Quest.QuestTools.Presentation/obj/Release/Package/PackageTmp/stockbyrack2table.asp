<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- August 18th, Shaun Levy asked for Table version of the StockbyRack2 - but original form looks better on the phone, so kept both-->
<!-- Both StockbyRack2 and StockbyRack2Table are options to run-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>iUI Theaters</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
 <script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

<%
afilter = request.QueryString("aisle")	


If CountryLocation = "USA" then
	strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'JUPITER' AND AISLE = '" & afilter & "' ORDER BY RACK, SHELF ASC"
else
	strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'GOREWAY' AND AISLE = '" & afilter & "' ORDER BY RACK, SHELF ASC"
end if

Set rs = Server.CreateObject("adodb.recordset")

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Stock by Rack Table</h1>
        <a class="button leftButton" type="cancel" href="stockbyaisle.asp" target="_self">Stock</a>
        </div>

        <ul id="Profiles" title="Profiles" selected="true">
       <li class='group'><a href="stockbyrack2.asp?aisle=<% response.write afilter %>" target="_self"> Switch to Row Form - Aisle <% response.write afilter %></a></li>

<% 
rack = "ABC"
aisle = "000"

Do While Not rs.eof
	part = rs("part")
	qty = rs("qty")
	id = rs("ID")
	po = rs("PO")
	bundle = rs("Bundle")
	ExBundle = rs("ExBundle")
	shelf = rs("shelf")
	colour = rs("colour")
	datein = rs("datein")
	If aisle = rs("aisle") Then
	Else
		response.write "<li class='group'>Aisle " & rs("aisle") & "</li>"
	End If

	If rack = rs("rack") Then

	Else
		If ISNULL(rack) = -1 Then
		Else
			Response.write "</table>"
			Response.write "<li class='group'>Rack " & rs("rack") & "</li>"
			Response.write "<li><table border='1' class='sortable'><tr><th>Shelf</th><th>Part</th><th>Quantity</th><th>PO</th><th>Bundle</th><th>External Bundle</th><th>Colour</th><th>Input Date</th></tr>"

		End If
	End If

%>
<!-- At request of Ruslan, Added Colour to this screen, January 17 2014, Michael Bernholtz-->
<tr>
<td><%response.write shelf %></td>
<td><%response.write part %></td>
<td><%response.write qty %></td>
<td><%response.write PO %></td>
<td><%response.write bundle %></td>
<td><%response.write ExBundle %></td>
<td><%response.write colour %></td>
<td><%response.write datein %></td>
<td><a href="stockbyrackedit.asp?id=<% response.write id %>&aisle=<% response.write afilter %>&ticket=GOREWAYTABLE" target="_self">Manage</a></td>
</tr>

<%

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
