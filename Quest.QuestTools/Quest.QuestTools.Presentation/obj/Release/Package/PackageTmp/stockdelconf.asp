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

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From Y_INVLOG"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL, DBConnection



ticket = request.QueryString("ticket")
part = request.QueryString("part")
po = request.Querystring("po")
pid = REQUEST.QueryString("ID")


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

color = REQUEST.QueryString("color")
length = REQUEST.QueryString("length")


if length < 100 then
linch = Round(length * 12,0)
lmm = Round(linch * 25.4,0)
lft = length
else
linch = length
lmm = Round(linch * 25.4,0)
lft = Round(length /12,0)
end if

if length > 300 then
linch = Round(length / 25.4,0)
lmm = length
lft = Round(length / 304.8,0)
end if

aisle = REQUEST.QueryString("aisle")
rack = REQUEST.QueryString("rack")
shelf = REQUEST.QueryString("shelf")
qty = REQUEST.QueryString("qty")
po = REQUEST.QueryString("po")
deleteby = REQUEST.QueryString("deleteby")
warehouse = REQUEST.QueryString("warehouse")

rs2.AddNew
	rs2.Fields("Part") = part
	rs2.Fields("colour") = color
	rs2.Fields("qty") = qty
	rs2.Fields("linch") = linch
	rs2.Fields("lmm") = lmm
	rs2.Fields("lft") = lft
	rs2.Fields("warehouse") = warehouse
	rs2.Fields("aisle") = aisle
	rs2.Fields("rack") = rack
	rs2.Fields("shelf") = shelf
	rs2.Fields("transaction") = "exit"
	rs2.Fields("day") = cday
	rs2.Fields("month") = cmonth
	rs2.Fields("year") = cyear
	rs2.Fields("week") = weeknumber
	rs2.Fields("time") = cctime
	rs2.Fields("ItemId") = pid
	rs2.Fields("PO") = po
	rs2.Fields("Project") = deleteby
	
	rs2.update

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
			<%
			Select Case ticket
	Case "pending"
		%>
		<a class="button leftButton" type="cancel" href="stockbypo.asp?PO=<% response.write po %>" target="_self">Pending PO</a>

		<%
	Case "goreway"
		%>
		<a class="button leftButton" type="cancel" href="stockgbypo.asp?PO=<% response.write po %>" target="_self">Goreway PO</a>
		<%
	Case "order"
		%>
		<a class="button leftButton" type="cancel" href="stockpending.asp" target="_self">On Order</a>	
		<%
	Case "other"
		%>
		<a class="button leftButton" type="cancel" href="stockother.asp" target="_self">Prod Stock</a>	
		<%
	Case else
		%>
                <a class="button leftButton" type="cancel" href="stockdel.asp?part=<% response.write part %>" target="_self">Delete Menu</a>
		<%
	End Select
		%>

			   
    </div>
    
      <%                  
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
pid = request.querystring("id")

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "ID = " & pid

part = request.querystring("part")
color = request.querystring("color")
qty = request.querystring("Qty")
linch = request.querystring("length")
aisle = request.querystring("aisle")
rack = request.querystring("rack")
shelf = request.querystring("shelf")

%>
    
<form id="conf" title="Delete Stock" class="panel" name="conf" action="stock.asp#_del" method="GET" target="_self" selected="true" >              

        <h2>Stock Deleted</h2>

<% 
rs.Delete

   %>


        <BR>
       <!-- Commented out, Back Button goes back to home 
		<input type="hidden" name='ticket' id='ticket' value="<%response.write ticket %>">
        <input type="text" name='part' id='part' value="<%response.write part %>"> -->
		
         <a class="whiteButton" href="javascript:conf.submit()">Back to Stock</a>
            
            </form>

            
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

