<!--#include file="dbpath_secondary.asp"-->
<!--// dbpath_Quest_InventoryReports.asp //-->
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen - January 21st, 2014 Michael Bernholtz-->
 <script src="sorttable.js"></script>
  <script >
 table {
border-collapse:collapse;
}
</script>
<%

ReportName = Request.Querystring("ReportName")
ReportMonth = Left(Right(ReportName,6),2)
IF Left(ReportMonth,1) = 0 Then
	ReportMonth = Right(ReportMonth,1)
end if
ReportYear = Right(ReportName,4)
AlumPrice = Request.Querystring("AlumPrice")
RecordNumbers = 0

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM [" & ReportName & "] WHERE (WAREHOUSE = 'GOREWAY' OR WAREHOUSE = 'NASHUA' OR WAREHOUSE = 'TORBRAM' OR WAREHOUSE = 'DURAPAINT' OR WAREHOUSE = 'DURAPAINT(WIP)' OR WAREHOUSE = 'HORNER' OR WAREHOUSE = 'NPREP') AND MONTH(DATEIN) = '" & REPORTMONTH & "' AND YEAR(DATEIN) = '" & REPORTYEAR & "'order by PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection2

'Aluminum price
alumprice = AlumPrice

'response.buffer = false 
%> 
<!-- January 21 - turned <Table> into <table border='1' class='sortable'> and added the sortable table script-->
<h2> Inventory from Goreway / Nashua / Nashua Prep / Durapaint / Durapaint (WIP) / Horner / TORBRAM </h2>
<h2> Entered during <%Response.write ReportMonth%> / <%Response.write ReportYear%></h2>
<h3> Using Inventory from: <%Response.write ReportName%></h3>
<h3> Aluminium Price: <%Response.write AlumPrice%></h3>

<table border='1' class='sortable'>
  <tr><th>Part</th>
    <th>Qty</th>
    <th>Length (mm)</th>
    <th>Colour</th>
	<th>KGM</th>
   <th>Bundle</th>
   <th>Date IN</th>
   <th>Warehouse</th>
    <th>Value</th>
	 <th>White Paint</th>
	  <th>Color Paint</th>
    </tr>
<%

rs.movefirst
do while not rs.eof
	invpart = rs("part")
	partqty = rs("Qty")
	lmm = CDBL(rs("lmm"))
	linch = CDBL(rs("linch"))
	colour = rs("colour")
	project = rs("project")
	bundle = rs("bundle")
	DateIN = rs("DateIN")
	Warehouse = rs("Warehouse")
	kgm = rs("kgm")
	errormsg = 0

'Create a Query
    SQL2 = "SELECT * FROM Y_MASTER where PART = '" & invpart & "' order BY ID DESC"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL2)
	
	if rs2.eof then
		errormsg = 1
	else
		if kgm = "0" then
			kgm =0
			kgm = rs2("kgm")
		end if
			errormsg = 0
	end if
	rs2.close
	set rs2 = nothing

	if errormsg = 1 then
		response.write invpart & " not in inventory master <BR>"
	end if

'response.write partqty & "<BR>"
If Not IsNull(PARTQTY) Then

	If kgm > 5 then 
		pricebar = partqty * kgm
		value2 = value2 + pricebar
	'this code is 5 times overvalued
	Else
		If invpart = "Que-157" Then
			tempvalue =  0 * (partqty * (kgm * ( lmm / 1000 )))
		Else
			tempvalue =  ALUMPRICE * (partqty * (kgm * ( lmm / 1000 )))
		End If

		value = value + tempvalue
	End If

	If colour = "White" Then
		paintlft = (linch/12) * partqty
		totalpaintlft = totalpaintlft + paintlft
	Else

	If colour = "Mill" Then
	Else
		paintlft2 = (linch/12) * partqty
		totalpaintlft2 = totalpaintlft2 + paintlft2
	End If
End If

%>

<tr><td><% response.write invpart %></td>
	<td><% response.write partqty %></td>
	<td><% response.write lmm %></td>
	<td><% response.write colour %></td>
	 <td><% response.write kgm %></td>
	<td><% response.write bundle %></td>
	<td><% response.write DateIn %></td>
	<td><% response.write Warehouse %></td>
  <!--  <td></td> -->
	<td>$<% if kgm > 5 then
	response.write round(pricebar,2)
	else
	response.write round(tempvalue,2)
	end if %></td>

<%
	if colour = "White" then
		response.write "<td>$" & round(paintlft*0.14,2) & "</td><td></td>"
	else
		If colour = "Mill" then	
			response.write "<td>&nbsp;</td><td>&nbsp;</td>"
		else
			response.write "<td></td><td>$" & round(paintlft2*0.38,2) & "</td>"
		end if 
	end if
%>
</tr>
<%

Else
	Response.Write "QTY IN INVENTORY IS ZERO <BR>"
End If

	RecordNumbers = RecordNumbers + 1
	rs.movenext
loop
rs.close
set rs=nothing

DBConnection.close
set DBConnection = nothing
DBConnection2.close
set DBConnection2 = nothing

%> </Table> <%

paintvalue1 = (totalpaintlft *0.14)
paintvalue2 = (totalpaintlft2 *0.38)
response.write "<BR><HR>"
response.write "Total kgm Material $" & round(value,2) & "<BR>"
response.write "Total perbar Material $" & round(value2,2)  & "<BR>"
response.write "<HR>"
response.write "SubTotal = $" & round(value,2) + round(value2,2) & "<BR><BR>"
response.write "Total Paint White $" & paintvalue1 & "<BR>"
response.write "Total Paint Project $" & paintvalue2 & "<BR>"
response.write "<HR>"
response.write RecordNumbers & " Records"& "<BR>"
response.write "Grand Total = $" & round(value,2) + round(value2,2) + round(paintvalue1,2) + round(paintvalue2,2)

%>

</body>
</html>
