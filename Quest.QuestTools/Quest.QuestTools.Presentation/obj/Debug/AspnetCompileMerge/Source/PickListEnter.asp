                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Pick List Entry Form -->
<!-- Displays all the other JOB and FLOOR information below and Remembers last entry-->
<!-- Requested By Ariel Aziza and Michael Angel, Built by Michael Bernholtz, January 2015-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Enter Demand</title>
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
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Pick" target="_self">Pick List</a>
        </div>
   
   <%
   
Added = False

JOB = REQUEST.QueryString("JOB")
FLOOR = REQUEST.QueryString("FLOOR")
DIE = REQUEST.QueryString("DIE")
LENGTH = REQUEST.QueryString("LENGTH")
if LENGTH = "" then
LENGTH = 0
end if
COLOUR = REQUEST.QueryString("COLOUR")
QTY = REQUEST.Querystring("QTY")
if QTY ="" then
QTY = 0
end if
INPUTDATE = Date
PickDate = REQUEST.QueryString("PickDate")

if NOT JOB = "" AND NOT FLOOR = "" AND NOT DIE = "" AND NOT LENGTH = 0 AND NOT COLOUR = "" AND NOT QTY = 0 AND NOT PickDate = "" then

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "INSERT INTO PickList ([JOB], [FLOOR], [DIE], [LENGTH], [COLOUR], [QTY], [ENTRYDATE], [PickDate]) VALUES( '" & JOB & "', '" & FLOOR &  "', '" & DIE & "', '" & LENGTH & "', '" & COLOUR & "', '" & QTY & "', #" & INPUTDATE & "# , #" & PickDate & "#)"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Added= True 
else

Added = False



end if

   %>
   
   
            
              <form id="enter" title="Enter New PickList Data" class="panel" name="enter" action="PickListEnter.asp" method="GET" target="_self" selected="true">
              
                              
        <h2>Enter New Glass Information:</h2>
		
		 <ul id="Profiles" title="Enter New PickList Data" selected="true">
<%		 
		 
IF Added = TRUE then
Response.write "<li>Demand added for " & JOB & " - " & FLOOR & "</li>"
Else
response.write "<li>ALERT! - - - Please fill in <b>all</b> Fields to add Demand - - -  ALERT!</li>"
end if		 

%>		 
		 <li><table border='1'> 
		 <tr>
		 <th>Job</th><th>Floor</th><th>Die</th><th>Colour</th><th>Length(Ft)</th><th>QTY</th><th>Pick Date</th>
		 </tr>

		<tr>
		<td><select name="Job">
			<% ActiveOnly = True %>
			<!--#include file="JobsList.inc"-->
			<% 
			' Coded this 3 times - to show Description again, despite collected value is TYPE ( USER CHOOSES DESCRIPTION, but SYSTEM NEEDS TYPE)
			rsJob.filter = "Job = '" & Job & "'"
			if rsJob.eof then
			%><option value = "" selected>-</option><%
			else
			%>
			<option value = "<% response.write rsJob("Job") %>" selected><% response.write rsJob("Job") %></option> 
			<%
			end if
			
			rsJob.Close
			set rsJob = nothing
			%>
		</select></td>
		<td><input class="NoMargin" type="text" name='FLOOR' id='FLOOR' size='10' value = "<% response.write FLOOR %>" ></td>
		<td><select name="Die">
			<option value = "" selected>-</option>
			<!--#include file="DiesList.inc"-->
			<% 
			rsDie.filter = "PART = '" & DIE & "'"
			if rsDie.eof then
			%><option value = "" selected>-</option><%
			else
			%>
			<option value = "<% response.write rsDie("Part") %>" selected><% response.write rsDie("PART") %> - <% response.write rsDie("DESCRIPTION") %></option> 
			<%
			end if
			rsDie.Close
			set rsDie = nothing
			%>
		</select></td>
		<td><select name="Colour">
		<option value = "" selected>-</option>		
		<%		
		Set rsColour = Server.CreateObject("adodb.recordset")
			ColourSQL = "Select Distinct CODE FROM Y_Color order by CODE ASC"
			rsColour.Cursortype = 2
			rsColour.Locktype = 3
			rsColour.Open ColourSQL, DBConnection
			
			do while not rsColour.eof
				Response.Write "<option name='Colour', value = '" & rsColour("Code") & "'>"
				Response.Write rsColour("Code") 
				Response.Write "</option>"
			rsColour.movenext
			loop
	
			
			rsColour.filter = "CODE = '" & COLOUR & "'"
			if rsColour.eof then
			%><option value = "" selected>-</option><%
			else
			%>
			<option value = "<% response.write rsColour("Code") %>" selected><% response.write rsColour("Code") %></option> 
			<%
			end if
			rsColour.Close
			set rsColour = nothing
			%>
		</select></td>		
		<td><input class="NoMargin"  type="number" name='LENGTH' id='LENGTH' size='10' value = "<% response.write LENGTH %>" ></td>
		<td><input class="NoMargin"  type="number" name='QTY' id='QTY' size='10' value = "<% response.write QTY %>" ></td>
		<td><input class="NoMargin"  type="date" name='PickDate' id='PickDate' size='10' value = "<% response.write PickDate %>" ></td>
		</table><li>        
        <br>    
         <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
        <br>
<%

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "Select * from PickList where [JOB] = '" & JOB & "' AND [FLOOR] = '" & FLOOR & "' order by ID DESC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection

		response.write "<li class='group'>PickList of Current Job(" & JOB & ") / Floor(" & FLOOR & ") </li>"
		response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
		response.write "<li><table border='1' class='sortable'><tr><th>Job</th><th>Floor</th><th>Die</th><th>Colour</th><th>Length</th><th>QTY</th><th>Pick Date</th><th>EntryDate</th></tr>"

if rs2.eof then
Response.write "<tr><td>No current Items</td></tr>"
end if		
do while not rs2.eof
	response.write "<tr>"
	response.write "<td>" & RS2("JOB") & "</td>"
	response.write "<td>" & RS2("FLOOR") & "</td>"
	response.write "<td>" & RS2("DIE") & "</td>"
	response.write "<td>" & RS2("COLOUR") & "</td>"
	response.write "<td>" & RS2("LENGTH") & "</td>"
	response.write "<td>" & RS2("QTY") & "</td>"
	response.write "<td>" & RS2("PickDate") & "</td>"
	response.write "<td>" & RS2("ENTRYDATE") & "</td>"
	response.write "</tr>"
	
	rs2.movenext
loop
response.write "</table></li>"


rs2.close
set rs2 = nothing

DBConnection.close 
set DBConnection = nothing
%>
<br>
</ul>
            </form>
                
             
               
</body>
</html>
