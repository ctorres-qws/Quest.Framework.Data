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

cid = REQUEST.QueryString("CID")


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
                <a class="button leftButton" type="cancel" href="colordel.asp" target="_self">Delete Color</a>
    </div>
    
      <%                  
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->


Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_COLOR ORDER BY PROJECT ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "ID = " & cid



%>
    
    
              <form id="cDelete" title="Delete Color" class="panel" name="cDelete" action="colordelconf.asp" method="GET" target="_self" selected="true" >
        <h2>Delete Color- <% response.write rs("Project") %></h2>
  
<fieldset>


    <h2> Are you sure you want to Delete:</h2>
	
	     <div class="row">
            <label>Job:  <% response.write rs("PROJECT")%></label>
        </div>
	     <div class="row">
            <label> Paint Code: <% response.write rs("CODE") & " at " & rs("COMPANY")%></label>
        </div>
	     <div class="row">
            <label>Side: <% response.write rs("SIDE")%></label>
        </div>
	     <div class="row">
            <label>Price Category: <% response.write rs("PRICECAT")%></label>
        </div>
		<div class="row">
            <label>ACTIVE: <% response.write rs("ACTIVE")%></label>
        </div>
		
		<div class="row">
            <label>EXTRUSION: <% response.write rs.fields("EXTRUSION") %> </label>
        </div> 	
		
		<div class="row">
            <label>SHEET: <% response.write rs.fields("SHEET") %> 
        </div> 	
		<div class="row">
            <label>Delete from Database (Delete Error / Do not Delete Old)</label>
            <input type="checkbox" name='del' id='del' >
        </div> 	


                  
			<input type="hidden" name='cid' id='cid' value="<%response.write rs.fields("id") %>">
            
</fieldset>


        <BR>
        <a class="redButton" type ="submit" href="javascript:cDelete.submit()">Delete Color</a><BR>

            
            </form>
            
  
<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

          
    
</body>
</html>
