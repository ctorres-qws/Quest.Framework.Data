<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Created April 2014, Michael Bernholtz at Request of Ariel Aziza and Jody Cash -->
<!-- View Stock by Date Pending Today and Future-->
<!--#include file="dbpath.asp"-->
<!-- StockbyPendingDate1.asp and stockbyPendingDate.asp-->

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

'part = request.QueryString("part")

'id = REQUEST.QueryString("ID")
'aisle = REQUEST.QueryString("aisle")


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
    </div>
    
        
    
              <form id="edit" title="Select Stock by PO" class="panel" name="edit" action="stockbypendingdate.asp" method="GET" target="_self" selected="true" > <form id="delete" title="Edit Stock" class="panel" name="delete" action="stockeditconf.asp" method="GET" target="_self" selected="true" >
        <h2>Search Pending Stock by Expected Date (DD/MM/YYYY)</h2>
  
   

<fieldset>
     <div class="row">
                <label>Date</label>
				
<%
dayin = Day(Date)
if dayin <10 then
	dayin = "0" & dayin
end if
monthin = Month(Date)
if monthin <10 then
	monthin = "0" & monthin
end if
yearin = Year(Date)

DateToday = yearin & "-" & monthin & "-"& dayin
%>				
				
				
				
				
                <input type="date" name='expdate' id='expdate' value='<% response.write DateToday %>' >
            </div>
            
</fieldset>


        <BR>
        <a type= "button" class="whiteButton" href="javascript:edit.submit()">Search pending stock by Pending Date</a><BR>
            
          <!--  <input type="hidden" name='id' id='id' value="<%response.write rs.fields("id") %>"> -->

            
            </form> </form>
            
 <form id="conf" title="Edit Stock" class="panel" name="conf" action="stock.asp#_remove" method="GET" target="_self">
        <h2>Stock Edited</h2>
  

            
            </form>
            
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

