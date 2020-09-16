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
        <h1 id="pageTitle">Print Lantek Labels</h1>
		<% 
		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>
                <a class="button leftButton" type="cancel" href="<%response.write Homesite%>#_Panel" target="_self">Panel<%response.write HomeSiteSuffix%></a>
    </div>   
        
    
              <form id="edit" title="Panel Job Name" class="panel" name="edit" action="label_lantek_v1.asp" method="GET" target="_self" selected="true" >
        <h2>Enter LANTEK Job Name for Labels</h2>
  
   

<fieldset>
    <div class="row">
		<label>Job</label>
		<input type="text" name='Job_Name' id='Job_Name'>
    </div>
	<div class="row">
		<label>Country</label>
		<select name="Country">
			<option value="CANADA" >CANADA</option>
			<option value="USA" >USA</option>
		</select>
    </div>
            
</fieldset>


        <BR>
		<a class="lightblueButton" onClick="edit.action='label_lantek_v1.asp'; edit.submit()">Create Label</a><BR>
        <a class="redButton" onClick="edit.action='label_lantek_colour.asp'; edit.submit()">Create Colour Label</a><BR>
        
            </form> </form>
<% 
DBConnection.close
set DBConnection=nothing
%>           
    
</body>
</html>



