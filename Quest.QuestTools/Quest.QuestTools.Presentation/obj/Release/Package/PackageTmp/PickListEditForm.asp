<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

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
  
  <script>
function startTime()
{
var today=new Date();
var h=today.getHours();
var m=today.getMinutes();
var s=today.getSeconds();
// add a zero in front of numbers<10
m=checkTime(m);
s=checkTime(s);
document.getElementById('clock').innerHTML=h+":"+m+":"+s;
t=setTimeout(function(){startTime()},500);
}

function checkTime(i)
{
if (i<10)
  {
  i="0" + i;
  }
return i;
}
</script>


<% 

PKid = REQUEST.QueryString("PKid")

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="PickListedit.asp" target="_self">Manage PL</a>
    </div>
    
      <%                  

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM PickList"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "ID = " & PKid

if rs.eof then 
else
JOB = rs("JOB")
FLOOR = rs("FLOOR")
DIE = rs("DIE")
QTY = rs("QTY")
LENGTH = rs("LENGTH")
COLOUR = rs("COLOUR")
PICKDATE = rs("PICKDATE")
end if

%>
    
    
              <form id="edit" title="Edit Pick List Item" class="panel" name="edit" action="PickListeditconf.asp" method="GET" target="_self" selected="true" > 
        <h2>Edit Pick List Item</h2>
		<fieldset>


        <div class="row">
           <label>Job Code</label>
           <select name="Job">
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
		</select>
        </div>

		<div class="row">
            <label>Floor</label>
            <input name='Floor' type="text" id='Floor' value="<% response.write FLOOR %>" >
        </div>
		
		<div class = "row">
		<label> Die / Part </Label>
		<select name="Die">
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
		</select>
		</div>
		
		<div class="row">
		<label> Colour </Label>
		<select name="Colour">
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
		</select>
		</div>
	
		<div class="row">
			<label> Length (Ft) </label>
			<input type="number" name='LENGTH' id='LENGTH' size='10' value = "<% response.write LENGTH %>" >
		</div>
		<div class="row">
			<label> Quantity </label>
			<input type="number" name='QTY' id='QTY' size='10' value = "<% response.write QTY %>" ></td>
		</div>   
		<div class="row">
			<label> Pick Date </label>
			<% 
		if isDate(PickDate) then
			InputDate = PickDate
			mm = Month(InputDate)
			dd = Day(InputDate)
			yy = Year(InputDate)
			IF len(mm) = 1 THEN
			  mm = "0" & mm
			END IF
			IF len(dd) = 1 THEN
			  dd = "0" & dd
			END IF
			pDate = yy & "-" & mm & "-" & dd 
		else
		pDate = NULL
		end if
		%>
			<input type="date" name='PickDate' id='PickDate' size='20' value = "<% response.write pDate %>" ></td>
		</div>		
		
            <input type="hidden" name='PKid' id='PKid' value="<%response.write rs.fields("id") %>" >
                      
</fieldset>


        <BR>
        <a class="whiteButton" href="javascript:edit.submit()">Submit Changes</a><BR>
            
            
      
            
            </form> 

  
<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

          
    
</body>
</html>