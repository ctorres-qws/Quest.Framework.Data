                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!--Productiontoday.asp Written as a Basic Page for both Online use and E-mail Report-->
<!--ProductionTodayTable.asp - Table Version -->
<!--Shows all items transfered to Production Today into  Window or COM Production-->
<!--July 2014, as requested by Shaun Levy and Jody Cash, by Michael Bernholtz-->
<!--May 2017 Jody requested change from Sort by Part to Sort by COLOUR-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Production Today</title>
 
    
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
strSQL = "SELECT * FROM Y_INV WHERE DATEOUT = #" & currentDate & "# OR DATEOUT = #" & Yesterday & "# ORDER BY WAREHOUSE, COLOUR, PART"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection


Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_MASTER ORDER BY PART ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>
   
   
         
       
        <ul id="Profiles" title="Production <% response.write CurrentDate %>" selected="true">
         <li class="group"><a href="productiontoday.asp?part=<%response.write part%>&CDay=<%response.write CDay%>" target="_self" >Production Today (Table Form) - Switch to Row Form</a></li>
        
<% 
rs.filter = "WAREHOUSE='WINDOW PRODUCTION' AND DATEOUT = #" & currentDate & "#"

response.write "<li class='group'>WINDOW PRODUCTION</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Floor / Notes</th></tr>"


do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if

Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=prodtodaytabletable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"


rs.filter = "WAREHOUSE='COM PRODUCTION' AND DATEOUT = #" & currentDate & "#"

response.write "<li class='group'>COM PRODUCTION</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th></tr>"


do while not rs.eof
	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
	
Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=prodtodaytabletable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"

rs.filter = "WAREHOUSE='WINDOW PRODUCTION' AND DATEOUT = #" & Yesterday & "#"
response.write "<li class='group'>--------------" & Yesterday & " --------------</li>"
response.write "<li class='group'> YESTERDAY WINDOW PRODUCTION</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Floor / Notes</th></tr>"


do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
	
Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=prodtodaytabletable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"


rs.filter = "WAREHOUSE='COM PRODUCTION' AND DATEOUT = #" & Yesterday & "#"

response.write "<li class='group'>YESTERDAY COM PRODUCTION</li>"
RESPONSE.WRITE "<li><table border='1' class='sortable'>"
RESPONSE.WRITE "<tr><th>Part</th><th>Description</th><th>Quantity (SL)</th><th>Colour</th><th>PO</th><th>Length(Ft)</th><th>Floor / Notes </th></tr>"


do while not rs.eof

	rs2.filter = "Part = '" & rs("part") & "'"
	if rs2.eof then 
		Description = "N/A"
	else
		Description = rs2("Description")
	end if
	
Response.write "<tr>"
Response.write "<td><a href='stockbyrackedit.asp?ticket=prodtodaytable&id=" & rs("ID") & "' target='_self'>" & rs("part") & "</a></td>"
Response.write "<td>" & Description & " </td>"
Response.write "<td>" & rs("qty") & " </td>"
Response.write "<td>" & rs("colour") & " </td>"
Response.write "<td>" & rs("po") & " </td>"
Response.write "<td>" & rs("Lft") & " </td>"
Response.write "<td>" & rs("Note") & " </td>"
Response.write "</tr>"


rs.movenext
loop
Response.write "</table></li>"

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
