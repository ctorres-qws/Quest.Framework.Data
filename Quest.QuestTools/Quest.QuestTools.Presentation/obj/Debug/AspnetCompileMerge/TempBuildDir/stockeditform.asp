<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

		 
		 <!--April 2015 - this page is a duplicate of stockbyrackedit - any instances of it should be replaced-->
		 
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
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
'
''Create a Query
'    SQL = "Select * FROM Y_INV ORDER BY PART ASC"
''Get a Record Set
'    Set RS = DBConnection.Execute(SQL)
	
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV ORDER BY PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL3 = "SELECT * FROM Y_MASTER ORDER BY PART ASC"
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection


part = request.QueryString("part")

id = REQUEST.QueryString("ID")


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="stockedit.asp?part=<% response.write part %>" target="_self">Edit Stock</a>
    </div>
    
      <%                  
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
rs.filter = "ID = " & id
	rs3.filter = "Part = '" & part & "'"
	if rs3.eof then 
		Description = "N/A"
	else
		Description = rs3("Description")
	end if


%>
    
    
              <form id="edit" title="Edit Stock" class="panel" name="edit" action="stockeditconf.asp" method="GET" target="_self" selected="true" > <form id="delete" title="Edit Stock" class="panel" name="delete" action="stockeditconf.asp" method="GET" target="_self" selected="true" >
        <h2>Edit Stock - <% response.write rs("Part") %> - <% response.write Description %></h2>
		<h2>PLEASE INFORM MICHAEL BERNHOLTZ IF YOU LAND ON THIS PAGE!</h2>
   

<fieldset>


              <div class="row">
<!-- Colour Edited to be a Drop-Down from the Y_Color table - At Request of Ruslan - Michael Bernholtz, January 20, 2014-->
            <div class="row">
             <label>Color</label>
            <select name="color" id='color' >
			
<%
Response.Write "<option name=jobname value='" & rs.fields("colour") & "'> " &rs.fields("colour") & "</option>"

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_Color  WHERE ACTIVE = TRUE Order by Project ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection


rs2.movefirst
Do While Not rs2.eof

Response.Write "<option name=color value='"
Response.Write rs2("Project")' & " - " & rs2("DESC") & " - " & rs2("SIDE")
Response.Write "'>"
Response.Write rs2("Project")' & " - " & rs2("DESC") & " - " & rs2("SIDE")
response.write "</option>"

rs2.movenext
loop
rs2.close
set rs2 = nothing

%></select></DIV>
            

             <div class="row">
                <label>Length</label>
                <input type="text" name='length' id='length' value="<%response.write rs.fields("linch") %>">
            </div>
            
            <div class="row">
                <label>Qty</label>
                <input type="text" name='Qty' id='Qty' value="<%response.write rs.fields("qty") %>">
            </div>
            
            <div class="row">
                <label>Aisle</label>
                <input type="text" name='aisle' id='Aisle' value="<%response.write rs.fields("aisle") %>" >
            </div>
            
            <div class="row">
                <label>Rack</label>
                <input type="text" name='rack' id='Rack' value="<%response.write rs.fields("rack") %>" >
            </div>
            
            <div class="row">
                <label>Shelf</label>
                <input type="text" name='shelf' id='Shelf' value="<%response.write rs.fields("shelf") %>"> 
            </div>
                           
			<div class="row">
                <label>PO</label>
                <input type="text" name='po' id='po' value="<%response.write rs.fields("po") %>">
            </div>
			
			<div class="row">
                <label>Colour PO</label>
                <input type="text" name='colorpo' id='colorpo' value="<%response.write rs.fields("colorpo") %>">
            </div>
			
			<div class="row">
                <label>Bundle</label>
                <input type="text" name='bundle' id='bundle' value="<%response.write rs.fields("bundle") %>">
            </div>
						<div class="row">
                <label>Ext. Bundle</label>
                <input type="text" name='exbundle' id='exbundle' value="<%response.write rs.fields("exbundle") %>">
            </div>
			<%
			' Correct Format must be applied to Date field			
dayin = Day(rs.fields("ExpectedDate"))
if dayin <10 then
	dayin = "0" & dayin
end if
monthin = Month(rs.fields("ExpectedDate"))
if monthin <10 then
	monthin = "0" & monthin
end if
yearin = Year(rs.fields("ExpectedDate"))

DateEdit = yearin & "-" & monthin & "-"& dayin		
			
			%>
			<div class="row"> <!-- Date Field Added April 2014 - also updated in Stockbyrackeditconf treated as text for simplicity-->
                <label>Expected Date</label>
                <input type="date" name='expdate' id='expdate' value="<% response.write DateEdit %>"  >	
            </div>
			
			
			
			<div class="row">
				<label>Allocation</label>
				<select name="Allocation">
					<% ActiveOnly = True %>
					<option value="White" >White</option>
					<option value="" >None</option>
					<!--#include file="JobsList.inc"-->
					<% 
						' Coded this 3 times - to show Description again, despite collected value is TYPE ( USER CHOOSES DESCRIPTION, but SYSTEM NEEDS TYPE)
						rsJob.filter = "Job = '" & rs("Allocation") & "'"
						if rsJob.eof then
							%><option value = "" selected>-</option><%
						else
							%>
							<option value = "<% response.write rsJob("Job") %>" selected><% response.write rsJob("Job") %></option> 
							<%
						end if
						%>
				</select>
				
				<%
				rsJob.close
				set rsJob=nothing
				%>
			</div>
			
			
			
			<div class="row">
                <label>Thickness</label>
                <input type="text" name='thickness' id='thickness' value="<%response.write rs.fields("thickness") %>">
            </div>
			
            <div class="row">
				<label>Warehouse</label>
				<select name="warehouse">
					<option selected name=jobname value="-">-<%

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_WAREHOUSE"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

varw = 0

Response.Write "<option SELECTED=jobname value='"
Response.Write rs("WAREHOUSE")
Response.Write "'>"
Response.Write rs("WAREHOUSE")
response.write ""

rs2.movefirst
Do While Not rs2.eof
if rs2("NAME") = rs("WAREHOUSE") then
response.write ""
else
Response.Write "<option name=jobname value='"
Response.Write rs2("NAME")
Response.Write "'>"
Response.Write rs2("NAME")
response.write ""
end if

rs2.movenext

loop
rs2.close
set rs2 = nothing

%></select></DIV>
	<div class="row">
		<label>Floors</label>
		<input type="text" name='FloorNote' id='FloorNote' value="<%response.write rs.fields("Note") %>">
	</div>
            
     <div class="row">
		<label>Status Note</label>
		<input type="text" name='StatusNote' id='StatusNote' value="<%response.write rs.fields("Note 2") %>">
	</div>             
            
            
</fieldset>


        <BR>
        <a class="whiteButton" href="javascript:edit.submit()">Submit Changes</a><BR>
        <!--<a class="redButton" onClick="edit.action='stockdelform.asp'; edit.submit()">Delete Stock</a><BR>-->
		<!-- Removed Jan 2017 at request of Mary Darnell and Shaun Levy-->
            
            <input type="hidden" name='id' id='id' value="<%response.write rs.fields("id") %>">
            <input type="hidden" name='part' id='part' value="<%response.write rs.fields("part") %>">
            <BR>
		 <h2>Transfer Partial Stock to new Location</h2>
<fieldset>
	<div class="row">
		<label>Qty to Go</label>
		<input type="text" name='QtyMOVE' id='QtyMOVE' value="<%response.write rs.fields("qty") %>">
	</div>
	<div class="row">
		<label>Floors</label>
		<input type="text" name='FloorNote2' id='FloorNote2' value="<%response.write rs.fields("Note") %>">
	</div>

  <div class="row">
             <label>Go To</label>
            <select name="warehouseMOVE">
<%

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_WAREHOUSE"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

varw = 0

Response.Write "<option name=warehouseMOVE SELECTED=True value='"
Response.Write "WINDOW PRODUCTION"
Response.Write "'>"
Response.Write "WINDOW PRODUCTION"
response.write ""

rs2.movefirst
Do While Not rs2.eof
if rs2("NAME") = rs("WAREHOUSE") then
response.write ""
else
Response.Write "<option name=warehouseMOVE value='"
Response.Write rs2("NAME")
Response.Write "'>"
Response.Write rs2("NAME")
response.write ""
end if

rs2.movenext

loop
rs2.close
set rs2 = nothing
%></select></DIV>
</fieldset>
		<a class="greenButton" onClick="edit.action='stockMoveconf.asp'; edit.submit()">Transfer Portion</a><BR>
		
            </form> </form>
            
 <form id="conf" title="Edit Stock" class="panel" name="conf" action="stock.asp#_remove" method="GET" target="_self">
        <h2>Stock Edited</h2>
  

            
            </form>
            
    
</body>
</html>

<% 

rs.close
set rs=nothing
rs3.close
set rs3=nothing

DBConnection.close
set DBConnection=nothing
%>

