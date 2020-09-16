                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Productiontoday.asp Written as a Basic Page for both Online use and E-mail Report-->
<!--ProductionTodayTable.asp - Table Version -->
<!--Shows all items transfered to Production Today into  Window or COM Production-->
<!--July 2014, as requested by Shaun Levy and Jody Cash, by Michael Bernholtz-->

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
	Job = Request.Querystring("Job")
	Floor = Request.Querystring("Floor")
	


Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE (COLOUR LIKE '%" & JOB & "%' OR JOBCOMPLETE LIKE '%" & JOB & "%' )ORDER BY WAREHOUSE, PART"
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
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Production <% response.write CurrentDate %> " selected="true">
         
<% 
response.write "<li class='group'>---GOREWAY---</li>"

rs.filter = "WAREHOUSE = 'GOREWAY'"
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
%>

<li><a href="stockbyrackedit.asp?ticket=prodtoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & description & ", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' | " & JobComplete & " " & FloorNote %></a></li>


 
<% 
rs.movenext
loop

response.write "<li class='group'>---NASHUA---</li>"
rs.filter = "WAREHOUSE = 'NASHUA'"
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
%>

<li><a href="stockbyrackedit.asp?ticket=prodtoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & description & ", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' | " & JobComplete & " " & FloorNote %></a></li>


 
<% 
rs.movenext
loop

response.write "<li class='group'>---DURAPAINT---</li>"

rs.filter = "WAREHOUSE = 'DURAPAINT'"
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
%>

<li><a href="stockbyrackedit.asp?ticket=prodtoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & description & ", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' | " & JobComplete & " " & FloorNote %></a></li>


 
<% 
rs.movenext
loop

response.write "<li class='group'>---DURAPAINT(WIP)---</li>"

rs.filter = "WAREHOUSE = 'DURAPAINT(WIP)'"
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
%>

<li><a href="stockbyrackedit.asp?ticket=prodtoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & description & ", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' | " & JobComplete & " " & FloorNote %></a></li>


 
<% 
rs.movenext
loop

response.write "<li class='group'>---HORNER---</li>"

rs.filter = "WAREHOUSE = 'HORNER'"
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
%>

<li><a href="stockbyrackedit.asp?ticket=prodtoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & description & ", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' | " & JobComplete & " " & FloorNote %></a></li>


 
<% 
rs.movenext
loop

response.write "<li class='group'>---NASHUA PREP---</li>"

rs.filter = "WAREHOUSE = 'NPREP'"
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
%>

<li><a href="stockbyrackedit.asp?ticket=prodtoday&id=<% response.write id %>" target="_self"> <%response.write part & " - " & description & ", " & qty & " SL" & ", " & Colour & " " & PO & " " & Lft & "' | " & JobComplete & " " & FloorNote %></a></li>


 
<% 
rs.movenext
loop



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
