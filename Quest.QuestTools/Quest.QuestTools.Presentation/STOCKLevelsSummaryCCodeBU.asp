<!--#include file="dbpath.asp"-->
              <!-- Changed September 2014, to include Durapaint and Goreway as Stock and rest as Pending: (DuraPaint removed from Pending)-->
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
<style TYPE="text/css">
table {
  zoom: 70%;
}
}
</style>
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
CCode = Request.QueryString("CCode")
Colour1 = 0
Colour2 = 0
Name = ""
%>  
<ul id="screen1" title="Stock Level <% response.write ": " & Job %>" selected="true">      
<form id="Job" title="Stock Level By Job" class="panel" name="job" action="stocklevelsSummaryCCode.asp" method="GET" target="_self" >
        <h2>Select CCode</h2>
  <fieldset>
            <div class="row">
                <label>Colour Code</label>
				<select name="CCode" onchange="this.form.submit()" >
				<option value ="">-</option>
				<%
                Set rsCCODE = Server.CreateObject("adodb.recordset")
				jobSQL = "Select * FROM Y_Color order by Code ASC"
				rsCCODE.Cursortype = 2
				rsCCODE.Locktype = 3
				rsCCODE.Open jobSQL, DBConnection
do while not rsCCODE.eof
	' Remove Duplicate Colour Codes
	Colour2 = Colour1
	Colour1 = rsCCODE("Code")
	ColName2 = ColName1
	ColName1 = rsccode("project")
	

	if Colour1 = Colour2 or Colour1 = "" then
	ColName = ColName & ", " & ColName2
	else
		if ColName = "" then
			ColName = ColName2
		else
			ColName = ColName & ", " & ColName2
		end if
		
		Response.Write "<option name='CCode', value = '" & Colour2 & "'>"
		Response.Write Colour2 & " - " & ColName
		Response.Write "</option>"
		if Colour2 = CCode then
			ColNameTitle = ColName
		end if
		ColName = ""
	end if
rsCCODE.movenext
loop

%>
				<option value ="ALL">ALL</option>
				</select>
            </div>
	</fieldset>		
</form>	

<%



'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER order by PART ASC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)

response.write "<li class='group'>Stock by Die - Mill/Painted " & CCode & " :::::: "& ColNameTitle &" </li>"
if CCode = "" then
CCode = "ALL"
end if
response.write "<li><table border='1' class='sortable' width ='100%'><tr><th>Stock</th><th>Description</th><th>Goreway Mill</th><th>" & CCode & ": Goreway </th><th>Durapaint Mill</th><th>Durapaint(WIP) Mill</th><th>Horner Mill</th><th>" & CCode & ": Horner </th><th>Nashua Mill</th><th>" & CCode & ": Nashua </th><th>SAPA Mill</th><th>" & CCode & ": SAPA </th><th>Pending</th><th>Total: " & CCode & "</th><th>Min level</th></tr>"
'<th>" & CCode & ": Durapaint </th>
' <th>" & CCode & ": Durapaint(WIP)  </th>
rs2.movefirst
	do while not rs2.eof
	'Goreway Mill and Colour
	GMqty = 0
	GCqty = 0
	'Durapaint Mill and Colour
	DMqty = 0
	DCqty = 0
	'Durapaint(WIP) Mill and Colour
	DWMqty = 0
	DWCqty = 0
	'Sapa Mill and Colour
	SMqty = 0
	SCqty = 0
	'Horner Mill and Colour
	HMqty = 0
	HCqty = 0
	'Nashua Mill and Colour
	NMqty = 0
	NCqty = 0
	
	partqty2 = 0
	partqty3 = 0
	
	
Set rs = Server.CreateObject("adodb.recordset")

if CCode = "" or CCode = "ALL" then
	strSQL = "SELECT * FROM Y_INV WHERE (Warehouse <> 'WINDOW PRODUCTION' AND Warehouse <> 'COM PRODUCTION' AND Warehouse <> 'SCRAP') And Part = '" & rs2("Part") & "' order by Colour ASC"
else
	
	strSQL = "SELECT * FROM Y_INV Inner Join Y_COLOR ON Y_INV.colour = Y_color.Project WHERE ( Y_Color.Code LIKE '%" & CCODE & "%' or Y_Color.Code = 'Mill' )  And (Warehouse <> 'WINDOW PRODUCTION' AND Warehouse <> 'COM PRODUCTION' AND Warehouse <> 'SCRAP') And Part = '" & rs2("Part") & "' order by Colour ASC"

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
				GMqty = rs("Qty") + GMqty
			else
				GCqty = rs("Qty") + GCqty
			end if
	CASE "DURAPAINT"
		if rs("colour") = "Mill" then
				DMqty = rs("Qty") + DMqty
			else
				DCqty = rs("Qty") + DCqty
			end if
	CASE "DURAPAINT(WIP)"
		if rs("colour") = "Mill" then
				DWMqty = rs("Qty") + DWMqty
			else
				DWCqty = rs("Qty") + DWCqty
			end if
	CASE "HORNER"
		if rs("colour") = "Mill" then
				HMqty = rs("Qty") + HMqty
			else
				HCqty = rs("Qty") + HCqty
			end if
	CASE "NASHUA"
		if rs("colour") = "Mill" then
				NMqty = rs("Qty") + NMqty
			else
				NCqty = rs("Qty") + NCqty
			end if
	CASE "SAPA"
		if rs("colour") = "Mill" then
				SMqty = rs("Qty") + SMqty
			else
				SCqty = rs("Qty") + SCqty
			end if
		
	CASE Else
		partqty3 = rs("Qty") + partqty3

	End Select

	rs.movenext
	loop
	
	
	'Added At Request of Ruslan - to Note the Min Levels
	'Min Levels was added to the Y_Master and then all the add/edit forms for Y_MASTER - March 13, 2014 - Michael Bernholtz
	MinLevelAlert = ""
	if GMqty + DMqty + DWMqty + HMqty + NMqty + SMqty + partqty3 + GCqty + DCqty + DWCqty + HCqty + NCqty + SCqty< rs2("MinLevel")   then
		MinLevelAlert = "Below"
	end if

	if CCode = "ALL" then
		if  GMqty = 0 and GCqty = 0 and DMqty = 0 and DCqty = 0 and DWMqty = 0 and DWCqty = 0 and HMqty = 0 and HCqty = 0 and NMqty = 0 and NCqty = 0 and SMqty = 0 and SCqty = 0 then
		else
			response.write "<tr><td><a href='stockLengthDrillDown.asp?part=" & rs2("Part") & "' target='_self'>" & rs2("part") & "</a></td>"
			response.write "<td>" & rs2("description") & "</td>"
			response.write "<td>" & GMqty & "</td><td>" & GCqty & "</td>"
			response.write "<td>" & DMqty & "</td>" '<td>" & DCqty & "</td>" Durapaint painted should always be 0
			response.write "<td>" & DWMqty & "</td>" '<td>" & DWCqty & "</td>" Durapaint painted should always be 0
			response.write "<td>" & HMqty & "</td><td>" & HCqty & "</td>"
			response.write "<td>" & NMqty & "</td><td>" & NCqty & "</td>"
			response.write "<td>" & SMqty & "</td><td>" & SCqty & "</td>"
			response.write "<td> " & partqty3 & "</td>"
			response.write "<td> " & GCqty + DCqty + DWCqty + HCqty + SCqty + NCqty  & "</td>"
			if MinLevelAlert ="Below" then
				response.write "<td><font color='red'> " & rs2("MinLevel") & "</font></td>"
			else
				response.write "<td>" & rs2("MinLevel") & "</td>"
		
			end if
			response.write "</tr>"
			'<td>" & DWqty & "</td>
		end if 
	else
		if  GCqty = 0 and DCqty = 0 and DWCqty = 0 and HCqty = 0 and NCqty = 0 and SCqty = 0  then
		else
			response.write "<tr><td><a href='stockLengthDrillDown.asp?part=" & rs2("Part") & "' target='_self'>" & rs2("part") & "</a></td>"
			response.write "<td>" & rs2("description") & "</td>"
			response.write "<td>" & GMqty & "</td><td>" & GCqty & "</td>"
			response.write "<td>" & DMqty & "</td>" '<td>" & DCqty & "</td>" Durapaint painted should always be 0
			response.write "<td>" & DWMqty & "</td>" '<td>" & DWCqty & "</td>" Durapaint painted should always be 0
			response.write "<td>" & HMqty & "</td><td>" & HCqty & "</td>"
			response.write "<td>" & NMqty & "</td><td>" & NCqty & "</td>"
			response.write "<td>" & SMqty & "</td><td>" & SCqty & "</td>"
			response.write "<td> " & partqty3 & "</td>"
			response.write "<td> " & GCqty + DCqty + DWCqty + HCqty + SCqty + NCqty & "</td>"
			if MinLevelAlert ="Below" then
				response.write "<td><font color='red'> " & rs2("MinLevel") & "</font></td>"
			else
				response.write "<td>" & rs2("MinLevel") & "</td>"

			end if
			response.write "</tr>"
		'<td>" & DWqty & "</td>
		end if 
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
rsCCODE.close
set rsCCode = nothing

rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>

