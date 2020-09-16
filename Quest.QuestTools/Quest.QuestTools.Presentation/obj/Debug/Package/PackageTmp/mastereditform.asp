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
                <a class="button leftButton" type="cancel" href="masteredittable.asp?part=<% response.write part %>" target="_self">Edit Master</a>
    </div>

<%
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_MASTER"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "ID = " & id

%>
    
    
              <form id="edit" title="Edit Master" class="panel" name="edit" action="mastereditconf.asp" method="GET" target="_self" selected="true" > <form id="delete" title="Edit Stock" class="panel" name="delete" action="stockeditconf.asp" method="GET" target="_self" selected="true" >
        <h2>Edit Master - 
          <% response.write rs("Part") %></h2>
  
   

<fieldset>


            <div class="row">
                <label>Q Part #</label>
                <input type="text" name='part' id='part' value="<%response.write rs.fields("part") %>">
            </div>
			
			
			<div class="row">
                <label>Description</label>
                <input type="text" name='description' id='description' value="<%response.write rs.fields("description")%>" >
            </div>
			<div class="row">
                <label>InventoryType</label>
                <select name="inventorytype">
					<option selected value="<%response.write rs.fields("inventorytype") %>"><%response.write rs.fields("inventorytype") %></option>
					<option value="Extrusion">Extrusion</option>
					<option value="Gasket">Gasket</option>
					<option value="Hardware">Hardware</option>
					<option value="Plastic">Plastic</option>
					<option value="Sheet">Sheet</option>
					<option value="NPrep FG">NPrep FG</option>
				</select>
            </div>

            <div class="row">
                <label>Supplier #</label>
                <input type="text" name='supplierpart' id='supplierpart' value="<%response.write rs.fields("supplierpart") %>">
            </div>
            
            <div class="row">
                <label>kgm</label>
                <input type="text" name='kgm' id='kgm' value="<%response.write rs.fields("kgm") %>">
            </div>
			<div class="row">
                <label>Lbf</label>
                <input type="text" name='lbf' id='lbf' value="<%response.write rs.fields("lbf") %>">
            </div>
             
			 <div class="row">
                <label>Min Stock #</label>
                <input type="number" name='MinLevel' id='MinLevel' value="<%response.write rs.fields("MinLevel") %>">
            </div>
			<%
			if rs.fields("InventoryType") = "Plastic" then
			%>
			<div class="row">
                <label>Min Stock 16</label>
                <input type="number" name='Min-16' id='Min-16' value="<%response.write rs.fields("Min-16") %>">
            </div>
			<div class="row">
                <label>Min Stock 18</label>
                <input type="number" name='Min-18' id='Min-18' value="<%response.write rs.fields("Min-18") %>">
            </div>
			<div class="row">
                <label>Min Stock 20</label>
                <input type="number" name='Min-20' id='Min-20' value="<%response.write rs.fields("Min-20") %>">
            </div>
			<div class="row">
                <label>Min Stock 21</label>
                <input type="number" name='Min-21' id='Min-21' value="<%response.write rs.fields("Min-21") %>">
            </div>
			<div class="row">
                <label>Min Stock 22</label>
                <input type="number" name='Min-22' id='Min-22' value="<%response.write rs.fields("Min-22") %>">
            </div>
			
			<%
			End if
			%>
			
			<div class="row">
                <label>Paint Cat.</label>
                <input type="text" name='paintcat' id='paintcat' value="<%response.write rs.fields("paintcat") %>" >
            </div>
			<div class="row">
                <label>HYDRO</label>
                <input type="text" name='HYDRO' id='HYDRO' value="<%response.write rs.fields("HYDRO") %>" >
            </div>
			<div class="row">
                <label>CanArt</label>
                <input type="text" name='canart' id='canart' value="<%response.write rs.fields("canart") %>" >
            </div>			
			<div class="row">
                <label>KeyMark</label>
                <input type="text" name='keymark' id='keymark' value="<%response.write rs.fields("keymark") %>" >
            </div>
			<div class="row">
                <label>Extal</label>
                <input type="text" name='extal' id='extal' value="<%response.write rs.fields("extal") %>" >
            </div>


				
</fieldset>


        <BR>
        <a class="whiteButton" href="javascript:edit.submit()">Submit Changes</a><BR>
            
            <input type="hidden" name='id' id='id' value="<%response.write rs.fields("id") %>">
      
            
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

