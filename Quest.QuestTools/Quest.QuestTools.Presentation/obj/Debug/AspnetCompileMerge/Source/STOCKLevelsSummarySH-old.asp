<!--#include file="dbpath.asp"-->
              <!-- Changed September 2014, to include Durapaint and Goreway as Stock and rest as Pending: (DuraPaint removed from Pending)-->
			  <!-- Changed August 2015 to add Torbram / Tilton-->
				<!-- Change requested by Shaun Levy, Approved by Jody Cash -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Levels</title>
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


<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Inventory</a>
    </div>

<%
Job = Request.QueryString("Job")
%>  
<ul id="screen1" title="Stock Level <% response.write ": " & Job %>" selected="true">   
<li class='group'><a href='StockLevelsSummaryExcel.asp?Job=<%response.write JOB %>' target='_self'>Send to Excel</a></li>         
<form id="Job" title="Stock Level By Job" class="panel" name="job" action="stocklevelsSummarySH.asp" method="GET" target="_self" >
        <h2>Select Job</h2>
  <fieldset>
            <div class="row">
                <label>Job</label>
				<select name="Job" onchange="this.form.submit()" >
				<% ActiveOnly = True %>
				<option value ="">-</option>
                <!--#include file="Jobslist.inc"-->
				<option value ="ALL">ALL</option>
				rsJob.close
				</select>
            </div>
	</fieldset>		
</form>	

<%

'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER where INVENTORYTYPE = 'Sheet' order by PART ASC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)

response.write "<li class='group'>Stock by Die - Mill/Painted " & Job & " </li>"
if job = "" then
job = "ALL"
end if
response.write "<li><table border='1' class='sortable' width ='100%'><tr><th>Stock</th><th>Description</th><th>Total</th><th>Min level</th><th>Goreway Mill</th><th>" & Job & ": Goreway Painted </th><th>Durapaint Mill</th><th>" & Job & ": Durapaint Painted </th><th>Durapaint (WIP) Mill</th><th>" & Job & ": Durapaint(WIP) Painted </th><th>Horner Mill</th><th>" & Job & ": Horner Painted </th><th>Nashua Mill</th><th>" & Job & ": Nashua Painted </th><th>Tilton Mill</th><th>" & Job & ": Tilton Painted </th></tr>"

rs2.movefirst
	do while not rs2.eof
	'Goreway and Goreway Painted
	Gqty = 0
	GAqty = 0
	'Durapaint and Durapaint Painted
	Dqty = 0
	DAqty = 0
	'Durapaint(WIP) and Durapaint(WIP) Painted
	DWqty = 0
	DWAqty = 0
	'Horner and Horner Painted
	Hqty = 0
	HAqty = 0
		'Nashua and Nashua Painted
	Nqty = 0
	NAqty = 0
		'Tilton and Tilton Painted
	Tiqty = 0
	TiAqty = 0
	
	partqty2 = 0
	partqty3 = 0
	
Set rs = Server.CreateObject("adodb.recordset")
if Job = "" or Job = "ALL" then
	strSQL = "SELECT * FROM Y_INV WHERE (Warehouse <> 'WINDOW PRODUCTION' AND Warehouse <> 'COM PRODUCTION' AND Warehouse <> 'SCRAP') And Part = '" & rs2("Part") & "' order by Colour ASC"
else
	strSQL = "SELECT * FROM Y_INV WHERE ((Colour = 'Mill' AND Allocation Like '%" & Job &"%' ) OR Colour LIKE '%" & Job & "%')  And (Warehouse <> 'WINDOW PRODUCTION' AND Warehouse <> 'COM PRODUCTION' AND Warehouse <> 'SCRAP') And Part = '" & rs2("Part") & "' order by Colour ASC"
end if
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection	
	if rs.eof then
	else
	rs.movefirst
	do while not rs.eof
	Select Case RS("WAREHOUSE")
	CASE "GOREWAY"
		if rs("colour") = "Mill" then
				Gqty = rs("Qty") + Gqty
			else
				GAqty = rs("Qty") + GAqty
			end if
		
	CASE "DURAPAINT"
		if rs("colour") = "Mill" then

				Dqty = rs("Qty") + Dqty
			else
				DAqty = rs("Qty") + DAqty
			end if

	CASE "DURAPAINT(WIP)"
		if rs("colour") = "Mill" then

				DWqty = rs("Qty") + DWqty
			else
				DWAqty = rs("Qty") + DWAqty
			end if

	CASE "HORNER"
		if rs("colour") = "Mill" then

				Hqty = rs("Qty") + Hqty
			else
				HAqty = rs("Qty") + HAqty
			end if
		
	CASE "NASHUA","NPREP"
		if rs("colour") = "Mill" then
				Nqty = rs("Qty") + Nqty
			else
				NAqty = rs("Qty") + NAqty
			end if
		
	CASE "TILTON"
		if rs("colour") = "Mill" then

				Tiqty = rs("Qty") + Tiqty
			else
				TiAqty = rs("Qty") + TiAqty
			end if
		
	End Select

	rs.movenext
	loop
	
	
	'Added At Request of Ruslan - to Note the Min Levels
	'Min Levels was added to the Y_Master and then all the add/edit forms for Y_MASTER - March 13, 2014 - Michael Bernholtz
	MinLevelAlert = ""
	if Gqty + Dqty + DWqty + Hqty + Toqty + Tiqty < rs2("MinLevel") then
		MinLevelAlert = "Below"
	end if

		if Gqty = 0 and GAqty = 0 and Dqty = 0 and DAqty = 0 and DWqty = 0 and DWAqty = 0 and Hqty = 0 and HAqty = 0 and NAqty = 0 and Nqty = 0 and TiAqty = 0 and Tiqty = 0  then
		else
			response.write "<tr><td><a href='stockLengthDrillDown.asp?ticket=sheet&part=" & rs2("Part") & "' target='_self'>" & rs2("part") & "</a></td>"
			response.write "<td>" & rs2("description") & "</td>"
			response.write "<td> " & Gqty+ GAqty + Dqty + DAqty + DWAqty + Hqty + HAqty + Nqty + NAqty + Tiqty + TiAqty & "</td>"
			if MinLevelAlert ="Below" then
				response.write "<td><font color='red'> " & rs2("MinLevel") & "</font></td>"
			else
				response.write "<td>" & rs2("MinLevel") & "</td>"
		
			end if
			response.write "<td>" & Gqty & "</td><td>" & GAqty & "</td>"
			response.write "<td>" & Dqty & "</td><td>" & DAqty & "</td>"
			response.write "<td>" & DWqty & "</td><td>" & DWAqty & "</td>"
			response.write "<td>" & Hqty & "</td><td>" & HAqty & "</td>"
			response.write "<td>" & Nqty & "</td><td>" & NAqty & "</td>"
			response.write "<td>" & Tiqty & "</td><td>" & TiAqty & "</td>"
			
			response.write "</tr>"
			'<td>" & DWqty & "</td>
		end if 

	
rs.close
set rs = nothing
end if
rs2.movenext
loop
response.write "</table></li>"

%>

   
            
   </ul>
</body>
</html>

<% 


rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>

