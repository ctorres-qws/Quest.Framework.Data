<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

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
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM PICKLIST ORDER BY JOB ASC, FLOOR ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
                <a class="button leftButton" type="cancel" href="index.html#_Pick" target="_self" >Pick List</a>
    </div>
    
        
    
              <form id="edit" title="Select Stock by PO" class="panel" name="edit" action="PickListFilterView.asp" method="GET" target="_self" selected="true" > 
        <h2>Search Pick Lists by Job or Colour or Die</h2>
  
<fieldset>
     <div class="row">
		<label>Job</label>
		<select name="Job">
			<option value = 'ANY' selected>Any</option>
			<% ActiveOnly = True %>
			<!--#include file="JobsList.inc"-->
			<%
			rsJob.Close
			set rsJob = nothing
			%>
			</select>
      </div>
	   <div class="row">
                <label>Floor</label>
				<input type='number' name='floor' id='floor' /> 
	  </div>
	  
	  <div class="row">
                <label>Colour</label>
				
			<select name="Colour">
		<option value = 'ANY' selected>Any</option>	
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
			rsColour.Close
			set rsColour = nothing
			%>
		</select>
      </div>
	  
	   <div class="row">
                <label>Die</label>		
		<select name="Die">
		<option value = 'ANY' selected>Any</option>
			<% ActiveOnly = True %>
			<!--#include file="DiesList.inc"-->
			<%
			rsDie.Close
			set rsDie = nothing
			%>
			</select>
      </div>
		
	  	   <div class="row">

                <label>Pick Date</label>
				<input type='date' name='PickDate' id='PickDate' placeholder= "dd/mm/yyyy" /> 
	  </div>
      </fieldset> <BR>
        <a class="lightblueButton" href="javascript:edit.submit()">Search Pick List by Die</a><BR>
	
		
		
            
            </form> 
            

            
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

