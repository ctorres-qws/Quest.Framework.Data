<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Productiontoday.asp Written as a Basic Page for both Online use and E-mail Report-->
<!--ProductionTodayTable.asp - Table Version -->
<!--Shows all items transfered to Production Today into  Window or COM Production-->
<!--July 2014, as requested by Shaun Levy and Jody Cash, by Michael Bernholtz-->
<!--May 2017 Jody requested change from Sort by Part to Sort by COLOUR and Add PRINT TO EXCEL button NOT TEXAS-->
<!--Feb 2019 - Added USA View -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Production Today</title>
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
	CurrentDate = Request.Querystring("CDay")
	CDay = currentDate  
	if CDay = "" then
		currentDate = Date()
		Yesterday = DateAdd("d",-1,Date())
	else

		Yesterday = DateAdd("d",-1,CDay)
	End if

Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT * FROM Y_INV WHERE DATEOUT = #" & currentDate & "# OR DATEOUT = #" & Yesterday & "# ORDER BY WAREHOUSE, Colour, Part")
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection


Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_MASTER ORDER BY PART ASC"
rs2.Cursortype = GetDBCursorType
rs2.Locktype = GetDBLockType
rs2.Open strSQL2, DBConnection


%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Production Today</h1>
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
   
   
         
       
	<ul id="Profiles" title="Production <% response.write CurrentDate %> " selected="true">
	<li class="group"><a href="productiontodaytable.asp" target="_self" >Production today (Row Form) - Switch to Table Form</a></li>
<% 
if CountryLocation = "USA" then
else	
%>
	<li class="group"><a href="productiontodaytableexcel.asp" target="_self" >PRINT TO EXCEL</a></li>
<%
end if
%>
 
<% 

if CountryLocation = "USA" then

rs.filter = "WAREHOUSE='JUPITER PRODUCTION' AND DATEOUT = #" & currentDate & "#"
response.write "<li class='group'>--------------" & currentDate & " --------------</li>"
response.write "<li class='group'>---JUPITER PRODUCTION---</li>"

do while not rs.eof
part = rs("part")

	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po")
FloorNote = rs("Note")
JobComplete = rs("JobComplete")
Allocation = rs("Allocation")
Location = rs("Aisle") & ":" & rs("Rack") & ":" & rs("Shelf")

%>

<li><a href="stockbyrackedit.asp?ticket=productiontoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & description & ", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' | " & JobComplete & " " & FloorNote & " - " & Allocation & " : " & Location %></a></li>


 
<% 
rs.movenext
loop


rs.filter = "WAREHOUSE='JUPITER PRODUCTION' AND DATEOUT = #" & Yesterday & "#"

response.write "<li class='group'>--------------" & Yesterday & " --------------</li>"
response.write "<li class='group'>--- YESTERDAY JUPITER PRODUCTION---</li>"

do while not rs.eof
part = rs("part")

	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po")
FloorNote = rs("Note")
JobComplete = rs("JobComplete")
Allocation = rs("Allocation")
Location = rs("Aisle") & ":" & rs("Rack") & ":" & rs("Shelf")
%>

<li><a href="stockbyrackedit.asp?ticket=prodtoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & description  &", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' | " & JobComplete & " " & FloorNote & " - " & Allocation & " : " & Location %></a></li>


 
<% 
rs.movenext
loop


else


rs.filter = "WAREHOUSE='WINDOW PRODUCTION' AND DATEOUT = #" & currentDate & "#"

response.write "<li class='group'>---WINDOW PRODUCTION---</li>"

do while not rs.eof
part = rs("part")

	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po")
FloorNote = rs("Note")
JobComplete = rs("JobComplete")
Allocation = rs("Allocation")
Location = rs("Aisle") & ":" & rs("Rack") & ":" & rs("Shelf")

%>

<li><a href="stockbyrackedit.asp?ticket=productiontoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & description & ", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' | " & JobComplete & " " & FloorNote & " - " & Allocation & " : " & Location %></a></li>


 
<% 
rs.movenext
loop

rs.filter = "WAREHOUSE='COM PRODUCTION' AND DATEOUT = #" & currentDate & "#"

response.write "<li class='group'>---COM PRODUCTION---</li>"

do while not rs.eof
part = rs("part")

	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po")
FloorNote = rs("Note")
JobComplete = rs("JobComplete")
Allocation = rs("Allocation")
Location = rs("Aisle") & ":" & rs("Rack") & ":" & rs("Shelf")
%>

<li><a href="stockbyrackedit.asp?ticket=prodtoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & description & ", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' | " & JobComplete & " " & FloorNote & " - " & Allocation & " : " & Location %></a></li>
<%

rs.movenext
loop

rs.filter = "WAREHOUSE='WINDOW PRODUCTION' AND DATEOUT = #" & Yesterday & "#"

response.write "<li class='group'>--------------" & Yesterday & " --------------</li>"
response.write "<li class='group'>--- YESTERDAY WINDOW PRODUCTION---</li>"

do while not rs.eof
part = rs("part")

	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po")
FloorNote = rs("Note")
JobComplete = rs("JobComplete")
Allocation = rs("Allocation")
Location = rs("Aisle") & ":" & rs("Rack") & ":" & rs("Shelf")
%>

<li><a href="stockbyrackedit.asp?ticket=prodtoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & description  &", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' | " & JobComplete & " " & FloorNote & " - " & Allocation & " : " & Location %></a></li>


 
<% 
rs.movenext
loop

rs.filter = "WAREHOUSE='COM PRODUCTION' AND DATEOUT = #" & Yesterday & "#"

response.write "<li class='group'>---YESTERDAY COM PRODUCTION---</li>"

do while not rs.eof
part = rs("part")

	rs2.filter = "Part = '" & part & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
qty = rs("qty")
id = rs("ID")
Lft = rs("Lft")
colour = rs("colour")
PO = rs("po")
FloorNote = rs("Note")
JobComplete = rs("JobComplete")
Allocation = rs("Allocation")
Location = rs("Aisle") & ":" & rs("Rack") & ":" & rs("Shelf")
%>

<li><a href="stockbyrackedit.asp?ticket=prodtoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & description & ", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' | " & JobComplete & " " & FloorNote & " - " & Allocation & " : " & Location %></a></li>
<%

rs.movenext
loop

end if 'CANADA/USA

rs.close
set rs = nothing
rs2.close
set rs2 = nothing
DBConnection.close
Set DBConnection = nothing

%>
      </ul>                 
            
            
            
       
            
              
               
                
             
               
</body>
</html>
