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
                <a class="button leftButton" type="cancel" href="coloredit.asp?id=<% response.write pid %>" target="_self">Edit Color</a>
    </div>
    
      
    
<form id="conf" title="Edit Color" class="panel" name="conf" action="coloredit.asp#_screen1" method="GET" target="_self" selected="true" >              

  
   
        
  
<%    
'Added extra code to match the Job table to the Colour Table, January 2015, Michael Bernholtz
              
cid = request.querystring("cid")

JOB = REQUEST.QueryString("JOB")
SIDE = REQUEST.QueryString("SIDE")
PROJECT = JOB & " " & SIDE
CODE = REQUEST.QueryString("CODE")
COMPANY = REQUEST.QueryString("COMPANY")
DESC = REQUEST.QueryString("DESCRIPTION")

PAINTCAT = REQUEST.QueryString("PAINTCAT")
ACTIVE = REQUEST.QueryString("ACTIVE")
if ACTIVE = "on" then
	ACTIVE = TRUE
else
	ACTIVE = FALSE
end if
EXTRUSION = REQUEST.QueryString("EXTRUSION")
if EXTRUSION = "on" then
	EXTRUSION = TRUE
else
	EXTRUSION = FALSE
end if
SHEET = REQUEST.QueryString("SHEET")
if SHEET = "on" then
	SHEET = TRUE
else
	SHEET = FALSE
end if

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

			'Set Color Update Statement
				StrSQL = FixSQLCheck("UPDATE Y_COLOR  SET PROJECT= '"& PROJECT & "', CODE = '"& CODE & "', JOB = '"& JOB & "', COMPANY = '"& COMPANY & "', SIDE = '"& SIDE & "', [DESC] = '" & DESC & "', PRICECAT ='" & PAINTCAT & "', ACTIVE= " & ACTIVE & " , EXTRUSION= " & EXTRUSION & " , SHEET= " & SHEET & " WHERE ID = " & CID, isSQLServer)
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)

DbCloseAll

End Function

%>

<h2>Color: <%response.write Project %> Edited</h2>
        <BR>
	<ul>
    <li> Colour Edited:</li>
	<li><% response.write "Job: " & PROJECT %></li>
	<li><% response.write "Paint Code: " & CODE & " at " & COMPANY %></li>
	<li><% response.write "Paint Location: " & DESC %></li>
	<li><% response.write "Side: " & SIDE %></li>
    <li><% response.write "Price Catagory: " & PAINTCAT %></li>
	<li><% response.write "ACTIVE: " & ACTIVE %></li>
		<%
	if EXTRUSION = TRUE then
	response.write "<li>Colour For: Extrusion</li>"
	end if
	%>
	<%
	if SHEET = TRUE then
	response.write "<li>Colour For: Sheet</li>"
	end if
	%>

	</ul>
         <a class="whiteButton" href="javascript:conf.submit()">Back to Colors</a>
            
            </form>

  
<% 



'DBConnection.close
'set DBConnection=nothing
%>

          
    
</body>
</html>
