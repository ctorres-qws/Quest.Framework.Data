<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
			<!--#include file="dbpath.asp"-->
<!--Search Stock Levels by Warehouse page -->

		 <!--Created May 23rd, by Michael Bernholtz - Overarching tool-->
		 <!--All  Warehouse version of Stock levels -->
		 <!-- Unsure if this will be a production tool, currently not in Use-->



<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Search Warehouses</title>
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
        <h1 id="pageTitle"></h1>
                <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self" >Inventory</a>
				<a class="button" href="#" id="clock"></a>
    </div>
    
        
    
              <form id="edit" title="Select Stock Level" class="panel" name="edit" action="StockLevelsWarehouse.asp" method="GET" target="_self" selected="true" > 
        <h2>Search Stock Level by Warehouse</h2>
  
   

<fieldset>
    
 <div class="row">
             <label>Warehouse</label>
            <select name="warehouse">
<%

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_WAREHOUSE ORDER BY ID ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

rs2.movefirst
Do While Not rs2.eof

Response.Write "<option value='"
Response.Write rs2("NAME")
Response.Write "'>"
Response.Write rs2("NAME")
response.write ""

rs2.movenext

loop
%></select>
</div>
</fieldset>


        <BR>
        <a class="whiteButton" href="javascript:edit.submit()">Search Warehouse by PO</a><BR>
            

            
            </form>
            

            
    
</body>
</html>

<% 

rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>

