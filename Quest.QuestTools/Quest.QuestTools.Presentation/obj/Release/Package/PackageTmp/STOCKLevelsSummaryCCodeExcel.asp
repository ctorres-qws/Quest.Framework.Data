<!--#include file="dbpath.asp"-->
              <!-- Changed September 2014, to include Durapaint and Goreway as Stock and rest as Pending: (DuraPaint removed from Pending)-->
			  <!-- Changed August 2015 to remove Mill and add Torbram / Tilton-->
				<!-- Change requested by Shaun Levy, Approved by Jody Cash -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Levels</title>
  
  
 <%


Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=ColourCode" & Date() & ".xls"
%>

 
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
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
	

	if TRIM(Colour1) = TRIM(Colour2) or Colour1 = "" then
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
response.write "<li><table border='1' class='sortable' width ='100%'><tr><th>Stock</th><th>Description</th><th>Total: " & CCode & "</th><th>Min level</th><th>" & CCode & ": Goreway </th><th>" & CCode & ": Horner </th><th>" & CCode & ": Durapaint </th><th>" & CCode & ": Durapaint(WIP) </th><th>" & CCode & ": HYDRO </th><th>" & CCode & ": NASHUA </th><th>" & CCode & ": Tilton </th><th>All MILL</th><th>Pending</th></tr>"

rs2.movefirst
	do while not rs2.eof
	'Goreway
	GCqty = 0
	'Durapaint 
	DCqty = 0
	'Durapaint(WIP) 
	DWCqty = 0
	'Sapa =
	SCqty = 0
	'Horner 
	HCqty = 0
		'Torbram
	NCqty = 0
		'Tilton
	TiCqty = 0
	
	'All Mill
	Mqty = 0
	
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
				Mqty = rs("Qty") + Mqty
			else
				GCqty = rs("Qty") + GCqty
			end if
	CASE "DURAPAINT"
		if rs("colour") = "Mill" then
				Mqty = rs("Qty") + Mqty
			else
				DCqty = rs("Qty") + DCqty
			end if
	CASE "DURAPAINT(WIP)"
		if rs("colour") = "Mill" then
				Mqty = rs("Qty") + Mqty
			else
				DWCqty = rs("Qty") + DWCqty
			end if
	CASE "HORNER"
		if rs("colour") = "Mill" then
				Mqty = rs("Qty") + Mqty
			else
				HCqty = rs("Qty") + HCqty
			end if
	CASE "SAPA","HYDRO"
		if rs("colour") = "Mill" then
				Mqty = rs("Qty") + Mqty
			else
				SCqty = rs("Qty") + SCqty
			end if
	CASE "NASHUA","NPREP"
		if rs("colour") = "Mill" then
				Mqty = rs("Qty") + Mqty
			else
				NCqty = rs("Qty") + NCqty
			end if
	CASE "TILTON"
		if rs("colour") = "Mill" then
				Mqty = rs("Qty") + Mqty
			else
				TiCqty = rs("Qty") + TiCqty
			end if
		
	CASE Else
		partqty3 = rs("Qty") + partqty3

	End Select

	rs.movenext
	loop
	
	
	'Added At Request of Ruslan - to Note the Min Levels
	'Min Levels was added to the Y_Master and then all the add/edit forms for Y_MASTER - March 13, 2014 - Michael Bernholtz
	MinLevelAlert = ""
	if Mqty + partqty3 + GCqty + DCqty + DWCqty + HCqty + SCqty+ TiCqty + NCqty < rs2("MinLevel")   then
		MinLevelAlert = "Below"
	end if
	
		if   GCqty = 0 and DCqty = 0 and DWCqty = 0 and HCqty = 0 and SCqty = 0 and NCqty = 0 and TiCqty = 0 then
		else
			response.write "<tr><td><a href='stockLengthDrillDown.asp?part=" & rs2("Part") & "' target='_self'>" & rs2("part") & "</a></td>"
			response.write "<td>" & rs2("description") & "</td>"
			
			response.write "<td> " & GCqty + DCqty + DWCqty + HCqty + SCqty  + TiCqty + NCqty & "</td>"
			if MinLevelAlert ="Below" and  CCode = "ALL" then
				response.write "<td><font color='red'> " & rs2("MinLevel") & "</font></td>"
			else
				response.write "<td>" & rs2("MinLevel") & "</td>"
			end if
			response.write "<td>" & GCqty & "</td>"
			response.write "<td>" & DCqty & "</td>"
			response.write "<td>" & DWCqty & "</td>"
			response.write "<td>" & HCqty & "</td>"
			response.write "<td>" & SCqty & "</td>"
			response.write "<td>" & NCqty & "</td>"
			response.write "<td>" & TiCqty & "</td>"
			response.write "<td>" & Mqty & "</td>"
			response.write "<td> " & partqty3 & "</td>"
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
rsCCODE.close
set rsCCode = nothing

rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>

