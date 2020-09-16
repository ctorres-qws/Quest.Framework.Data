<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
  <script src="sorttable.js"></script>
  
  
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
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>


    <form id="GlassProduction" title="Glass Line Report" class="panel" name="GlassProduction" action="GlassTimeReport.asp" method="GET" selected="true">

        <h2>Choose Day, Month, Year</h2>
        <fieldset>
       

         <div class="row">
                       
            <label>Day </label>
            <input type="number" name='sDay' id='sDay' value = "1">
		</div>
		<div class="row">
                       
            <label>Month </label>
			<select name='sMonth' id='sMonth'>
				<option value="1">January</option>
				<option value="2">February</option>
				<option value="3">March</option>
				<option value="4">April</option>
				<option value="5">May</option>
				<option value="6">June</option>
				<option value="7">July</option>
				<option value="8">August</option>
				<option value="9">September</option>
				<option value="10">October</option>
				<option value="11">November</option>
				<option value="12">December</option>
			</select>
		</div>		
         <div class="row">
                       
            <label>Year </label>
			<select name="sYear" id="sYear">
				<option value="2015">2015</option>
				<option value="2014">2014</option>
				<option value="2013">2013</option>
			</select>
		</div>
		
		<a class="whiteButton" onClick=" GlassProduction.submit()">View Glass Production</a><BR>
        </fieldset>	
		<p> Note: This report only works for a few weeks before the current day, It is only accurate based on X_BarcodeGA.</p> 
	</form>

<%
DBConnection.close
set DBConnection=nothing
%>


</body>
</html>

