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



	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="PickListEdit.asp?" target="_self">Edit Pick List</a>
    </div>
    
      
    
<form id="conf" title="Finalize Pick List" class="panel" name="conf" action="PickListEdit.asp" method="GET" target="_self" selected="true" >              

  
   
        
  
<%    
'Added extra code to match the Job table to the Colour Table, January 2015, Michael Bernholtz
JobTable = 0
              
PKid = request.querystring("PKid")

JOB = REQUEST.QueryString("JOB")
FLOOR = REQUEST.QueryString("FLOOR")
DIE = REQUEST.QueryString("DIE")
COLOUR = REQUEST.QueryString("COLOUR")
LENGTH = REQUEST.QueryString("LENGTH")
PICKDATE1 = REQUEST.QueryString("PICKDATE")
if isDate(PICKDATE1) then
PICKDATE = "#" & PICKDATE1 & "#"
else
PICKDATE = NULL
End if


If LENGTH = "" then
LENGTH = 0
End if
QTY = REQUEST.QueryString("QTY")
If QTY = "" then
QTY = 0
End if

	
			'Set Pick List Update Statement
				StrSQL = "UPDATE PickList  SET Job= '"& JOB & "', FLOOR = '"& FLOOR & "', DIE = '"& DIE & "', COLOUR = '"& COLOUR & "', LENGTH = "& LENGTH & ", QTY = " & QTY & ", PickDate = " & PickDate & " WHERE ID = " & PKID 
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
		
	

%>

<h2>Color: <%response.write Project %> Edited</h2>
        <BR>
	<ul>
    <li> Pick List Edited:</li>
	<li><% response.write "Job: " & JOB %></li>
	<li><% response.write "Floor: " & FLOOR %></li>
	<li><% response.write "Die/Part: " & DIE %></li>
	<li><% response.write "Colour: " & COLOUR %></li>
	<li><% response.write "Length (Ft): " & LENGTH %></li>
	<li><% response.write "Qty: " & QTY %></li>
	<li><% response.write "PickDate: " & PickDate1 %></li>

	</ul>
         <a class="whiteButton" href="javascript:conf.submit()">Back to Pick List Edit</a>
            
            </form>

  
<% 



DBConnection.close
set DBConnection=nothing
%>

          
    
</body>
</html>
