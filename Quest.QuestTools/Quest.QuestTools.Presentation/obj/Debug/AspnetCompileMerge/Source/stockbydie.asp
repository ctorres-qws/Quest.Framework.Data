<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 <!--Updated January 2014 to be both Table and Row form-->	
<!-- February 2019 - Added USA option-->	
		 
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
	
ticket = request.querystring("ticket")
	
Set rs = Server.CreateObject("adodb.recordset")

if CountryLocation = "USA" then
	strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'JUPITER' ORDER BY AISLE, RACK, SHELF ASC"
else
	strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'NPREP' ORDER BY AISLE, RACK, SHELF ASC"
end if

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_MASTER ORDER BY PART ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

sortby = REQUEST.QueryString("sortby")


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)


job = request.QueryString("job")
%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
	<%	
	Select Case ticket
	Case "pic"
		%>
		<a class="button leftButton" type="cancel" href="stockbypic.asp" target="_self">Stock by Pic</a>
		<%
	Case "level"
		
		if job = "" or isnull(job) then
		%>
		<a class="button leftButton" type="cancel" href="stocklevels.asp" target="_self">Stock Levels</a>
		<%
		else
		%>
		<a class="button leftButton" type="cancel" href="stocklevels.asp?job=<%response.write job%>" target="_self">Stock Levels</a>
		<%
		End if
	Case Else
		%>
			<a class="button leftButton" type="cancel" href="stock2.asp" target="_self">Stock by Die</a>
		<%
	End Select
	%>
    </div>

<ul id="screen1" title="Stock by Die" selected="true">

    

    <%
	part = request.QueryString("part")
	
	
	
	rs2.filter = "Part = '" & part & "'"
	
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
		InventoryType = rs2("InventoryType")
	end if
	Response.Write "<li class='group'><a href='stockbydieTable.asp?part=" & part & "&ticket=" & ticket &"' target='_self' >Stock (Row Form) - Switch to Table Form</a></li>"
	
	if not rs.eof then
	rs.movefirst
	else
	end if
	if job = "" or isnull(job) then
		rs.filter = "PART = '" & part & "' "
	else
		rs.filter = "Part = '" & part & "' AND (Colour LIKE '*" & job & "*') OR (Colour = 'Mill' AND Allocation Like '%" & Job &"%' )"
	end if


response.write "<li><img src='/partpic/" & part & ".png'/> - " & Description & " </li>"
do while not rs.eof
po = rs("PO")


response.write "<li>" & rs.fields("PART") & " "
	if isnull(rs.fields("Colour")) then 
	response.write rs.fields("Project")
	else
	response.write rs.fields("Colour")
	end if
	'Added Lmm (length in mm at Request of Ruslan - January 16, Michael Bernholtz

 Select Case InventoryType
	Case "Plastic"

response.write " " & rs.fields("Lft") & "' " & " " & rs.fields("Qty") & " SL" & " PO" & po & " / " & rs.fields("Bundle") & " " & rs.fields("Lmm") & " mm" &"</li>"
	
	Case "Sheet"
	
response.write " Thickness " & rs.fields("Thickness") & " / " & rs.fields("Qty") & " SL" & " PO" & po &"</li>"

	Case Else 'Extrusion
	
response.write " " & rs.fields("Lft") & "' " & " " & rs.fields("Qty") & " SL" & " PO" & po & " / " & rs.fields("Bundle") & " " & rs.fields("Lmm") & " mm" &"</li>"
End Select

if not rs.fields("Allocation") = "" then
response.write "<li>- Aisle " & rs.fields("Aisle") & " Rack " & rs.fields("rack") & " Shelf " & rs.fields("shelf") & "- Allocated to: " & rs.fields("Allocation") & "</li>"
else
response.write "<li>- Aisle " & rs.fields("Aisle") & " Rack " & rs.fields("rack") & " Shelf " & rs.fields("shelf") & "</li>"
end if
'response.write "<li><a href='stockdet.asp?id=" & rs.fields("ID") & "&part=" & part & "' target='_self'>" & rs.fields("PART") & " " & rs.fields("Colour") & " " & rs.fields("Lft") & "' " & " " & rs.fields("Qty") & " SL" & "</a></li>"
'if color is missing put project
rs.movenext
loop

RESPONSE.WRITE "</UL>"


wcount=0
JFCHECKID=0


rs.close
set rs=nothing
rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>


</body>
</html>


