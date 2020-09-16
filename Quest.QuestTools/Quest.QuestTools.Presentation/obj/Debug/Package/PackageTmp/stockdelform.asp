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
                <a class="button leftButton" type="cancel" href="stockdel.asp?part=<% response.write part %>" target="_self">Delete Stock</a>
    </div>
    
      <%                  
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->


Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "ID = " & id



%>
              <form id="sDelete" title="Delete Stock" class="panel" name="sDelete" action="stockdelconf.asp" method="GET" target="_self" selected="true" > 
        <h2>Delete Stock- <% response.write rs("Part") %></h2>
		<h2>Location:- <% response.write rs("Aisle") & ":" & rs("Rack") & ":"  & rs("Shelf") %></h2>
		<h2>Colour- <% response.write rs("colour") %></h2>
		<h2>Qty- <% response.write rs("qty") %></h2>
		<h2>Length- <% response.write rs("linch") %></h2>

<fieldset>
			<HR>
			<HR>
            <div class="row">

                <label>Deleted by: (REQUIRED)</label>
				<select name="Deleteby" id='Deleteby' >
				<option value= "Shaun">Shaun</option>
				<option value= "Ben">Ben</option>
				<option value= "David">David</option>
				<option value= "Mary">Mary</option>

				</select>

            </div>
			<HR>
			<HR>


            
			<input type="hidden" name='id' id='id' value="<%response.write rs.fields("id") %>">
            <input type="hidden" name='part' id='part' value="<%response.write rs.fields("part") %>">
                  
            
            
</fieldset>


        <BR>
        <a class="redButton" href="javascript:sDelete.submit()">Delete Stock</a><BR>
            

            
            </form>
            
 <form id="conf" title="Not in use" class="panel" name="conf" action="stock.asp#_remove" method="GET" target="_self">
        <h2>Not in use</h2>
  

            
            </form>
            
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

