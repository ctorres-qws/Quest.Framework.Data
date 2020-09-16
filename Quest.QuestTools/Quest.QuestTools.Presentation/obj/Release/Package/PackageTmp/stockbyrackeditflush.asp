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

'Added Flag to recognize which page sent to the edit page
'Helps design consistant back buttons - Michael Bernholtz at Request of Ruslan, March 2014
'edit here and at Back Button
ticket= request.QueryString("ticket")
po = request.QueryString("po")
bundle = request.QueryString("bundle")
thickness = request.QueryString("thickness")

part = request.QueryString("part")

id = REQUEST.QueryString("ID")
aisle = REQUEST.QueryString("aisle")


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
		<!--Back Button Flags were created so that many different pages could access this page, but all have working "Back Buttons"--> 
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
	Case "gorewayb"
		%>
		<a class="button leftButton" type="cancel" href="stockgbybundle.asp?bundle=<% response.write bundle %>" target="_self">Goreway Bundle</a>
		<%
	Case "allb"
		%>
		<a class="button leftButton" type="cancel" href="allbybundle.asp?bundle=<% response.write bundle %>" target="_self">All Bundle</a>
		<%
	Case "allbp"
		%>
		<a class="button leftButton" type="cancel" href="allbypobundle.asp?pobundle=<% response.write bundle %>" target="_self">All PO /Bundle</a>
		<%
	Case "allbex"
		%>
		<a class="button leftButton" type="cancel" href="allbyexbundle.asp?exbundle=<% response.write bundle %>" target="_self">All Ex. Bundle</a>
		<%
	Case "allpo"
		%>
		<a class="button leftButton" type="cancel" href="allbypo.asp?PO=<% response.write po %>" target="_self">ALL PO</a>
		<%
	Case "sapa"
		%>
		<a class="button leftButton" type="cancel" href="sapabypo.asp?PO=<% response.write po %>" target="_self">SAPA PO</a>
		<%
	Case "stocksapa"
		%>
		<a class="button leftButton" type="cancel" href="stocksapa.asp?PO=<% response.write po %>" target="_self">SAPA ALL</a>
		<%
	Case "warehouse"
		%>
		<a class="button leftButton" type="cancel" href="warehousebypo.asp?PO=<% response.write po %>" target="_self">Warehouse PO</a>
		<%
	Case "order"
		%>
		<a class="button leftButton" type="cancel" href="stockpendingflush.asp" target="_self">On Order</a>	
		<%
	Case "pendDate"
		%>
		<a class="button leftButton" type="cancel" href="stockbypendingdate.asp" target="_self">Pending Date</a>	
		<%
	Case "prod"
		%>
		<a class="button leftButton" type="cancel" href="productionbypo.asp?PO=<% response.write po %>" target="_self">Production PO</a>	
		<%
	Case "other"
		%>
		<a class="button leftButton" type="cancel" href="stockother.asp" target="_self">Prod All</a>	
		
		<%
	Case else
		%>
                <a class="button leftButton" type="cancel" href="stockbyrack2.asp?aisle=<% response.write aisle %>" target="_self">Edit Stock</a>
		<%
	End Select
		%>
		
		
    </div>			
    
      <%                  
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
ID = REQUEST.QueryString("ID")

rs.filter = "ID = " & id



%>
    
    
              <form id="edit" title="Edit Stock" class="panel" name="edit" action="stockbyrackeditconf.asp" method="GET" target="_self" selected="true" > <form id="delete" title="Edit Stock" class="panel" name="delete" action="stockeditconf.asp" method="GET" target="_self" selected="true" >
        <h2>Edit Stock - <% response.write rs("Part") %></h2>
  
   

<fieldset>
     <div class="row">
                <label>Part</label>
                <input type="text" name='part' id='part' value="<%response.write rs.fields("part") %>">
            </div>


              <div class="row">
<!-- Colour Edited to be a Drop-Down from the Y_Color table - At Request of Ruslan - Michael Bernholtz, January 20, 2014-->
            <div class="row">
             <label>Color</label>
            <select name="color" id='color' >
			
<%
Response.Write "<option name=color value='" & rs.fields("colour") & "'> " &rs.fields("colour") & "</option>"

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_Color  WHERE ACTIVE = TRUE Order by PROJECT ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection


rs2.movefirst
Do While Not rs2.eof

Response.Write "<option name=color value='"
Response.Write rs2("Project")' & " - " & rs2("DESC") & " - " & rs2("SIDE")
Response.Write "'>"
Response.Write rs2("Project")' & " - " & rs2("DESC") & " - " & rs2("SIDE")
response.write "</option>"

rs2.movenext

loop
rs2.close
set rs2 = nothing
%></select></DIV>
				
			
			
         <div class="row">

                        <div class="row">
                <label>Length</label>
                <input type="text" name='length' id='length' value="<%response.write rs.fields("linch") %>">
            </div>
            
              <div class="row">
                <label>Qty</label>
                <input type="text" name='Qty' id='Qty' value="<%response.write rs.fields("qty") %>">
            </div>
            
            
            
                     <div class="row">

                        <div class="row">
                <label>Aisle</label>
                <input type="text" name='aisle' id='Aisle' value="<%response.write rs.fields("aisle") %>" >
            </div>
            
                           <div class="row">
                <label>Rack</label>
                <input type="text" name='rack' id='Rack' value="<%response.write rs.fields("rack") %>" >
            </div>
            
             <div class="row">
                <label>Shelf</label>
                <input type="text" name='shelf' id='Shelf' value="<%response.write rs.fields("shelf") %>">
          
               
            </div>
            
			<% if ticket = "pending" then
' Correct Format must be applied to Date field			
dayin = Day(rs.fields("ExpectedDate"))
if dayin <10 then
	dayin = "0" & dayin
end if
monthin = Month(rs.fields("ExpectedDate"))
if monthin <10 then
	monthin = "0" & monthin
end if
yearin = Year(rs.fields("ExpectedDate"))

DateEdit = yearin & "-" & monthin & "-"& dayin		
			
			%>
			<div class="row"> <!-- Date Field Added April 2014 - also updated in Stockbyrackeditconf treated as text for simplicity-->
                <label>Expected Date</label>
                <input type="date" name='expdate' id='expdate' value="<% response.write DateEdit %>"  >	
            </div>
			<% end if%>
			
            <div class="row">
                <label>PO</label>
                <input type="text" name='PO' id='PO' value="<%response.write rs.fields("PO") %>">
            </div>
			<div class="row">
                <label>Colour PO</label>
                <input type="text" name='colorpo' id='colorpo' value="<%response.write rs.fields("colorpo") %>">
            </div>
			<div class="row">
				<label>Bundle</label>
                <input type="text" name='Bundle' id='Bundle' value="<%response.write rs.fields("Bundle") %>">
            </div>
			
			<div class="row">
				<label>Ext. Bundle</label>
                <input type="text" name='ExBundle' id='ExBundle' value="<%response.write rs.fields("ExBundle") %>">
            </div>
			
			<div class="row">
                <label>Thickness</label>
                <input type="text" name='thickness' id='thickness' value="<%response.write rs.fields("thickness") %>">
				<input type="hidden" name='ticket' id='ticket' value="<%response.write ticket %>">
            </div>
            
            <div class="row">
             <label>Warehouse</label>
            <select name="warehouse">
			<option name='flush' SELECTED> Flush </option>
<option name=jobname value="-">-<%

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "SELECT * FROM Y_WAREHOUSE"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

varw = 0

Response.Write "<option value='"
Response.Write rs("WAREHOUSE")
Response.Write "'>"
Response.Write rs("WAREHOUSE")
response.write ""

rs2.movefirst
Do While Not rs2.eof
if rs2("NAME") = rs("WAREHOUSE") then
response.write ""
else
Response.Write "<option name=jobname value='"
Response.Write rs2("NAME")
Response.Write "'>"
Response.Write rs2("NAME")
response.write ""
end if

rs2.movenext

loop
rs2.close
set rs2 = nothing
%></select></DIV>

            
                  
                        <input type="hidden" name='id' id='id' value="<%response.write rs.fields("id") %>">
            
</fieldset>


        <BR>
        
		
		<a class="whiteButton" onClick="edit.action='stockbyrackeditconf.asp'; edit.submit()">Submit Changes</a><BR>
        <a class="redButton" onClick="edit.action='stockdelconf.asp'; edit.submit()">Delete Stock</a><BR>


		<!--<a class="whitButton" href="javascript:edit.submit()">Submit Changes</a><BR>  -->


            
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

