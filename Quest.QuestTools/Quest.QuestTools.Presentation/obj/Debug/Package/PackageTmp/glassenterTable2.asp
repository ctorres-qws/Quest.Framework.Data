<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Glass Enter Form turned into a Table to mimic entry form-->
<!-- Displays all the other PO information below and Remembers last entry-->
<!-- Requested By Eric Bedeov, Built by Michael Bernholtz, Permission by Jody Cash-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Enter Glass</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
	 <script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>
   
   <%
   
Added = False

FLOOR = REQUEST.QueryString("FLOOR")
TAG = REQUEST.QueryString("TAG")
JOB = REQUEST.QueryString("PROJECT")
DEPARTMENT = REQUEST.QueryString("DEPARTMENT")

WIDTH = REQUEST.QueryString("WIDTH")
HEIGHT = REQUEST.QueryString("HEIGHT")
NOTES = REQUEST.QueryString("NOTES")
	NOTES = Replace(NOTES,"'","")

ONEMAT = REQUEST.QueryString("ONEMAT")
TWOMAT = REQUEST.QueryString("TWOMAT")
ONESPAC = REQUEST.QueryString("ONESPAC")
SPACERCOLOUR = REQUEST.QueryString("SPACERCOLOUR")
REQUIREDDATE = REQUEST.QueryString("REQUIREDDATE")
ORDERBY = REQUEST.QueryString("orderBy")
ORDERFor = REQUEST.QueryString("orderFor")
PoNum = REQUEST.QueryString("PoNum")
	PoNum = Replace(PoNum,"'","")

ExtorderNum = REQUEST.QueryString("ExtOrderNum")	
IntorderNum = REQUEST.QueryString("IntorderNum")	
AIR = REQUEST.QueryString("AIR")
CardOr = REQUEST.QueryString("CardinalOrder")
CardEX = REQUEST.QueryString("CardinalExpected")
ExtExpected = REQUEST.QueryString("ExtExpected")	
IntExpected = REQUEST.QueryString("IntExpected")	
ExtFrom = REQUEST.QueryString("ExtFrom")	
IntFrom = REQUEST.QueryString("IntFrom")	
' Added April 2015 for IVAN and SASHA
ONEMAT2 = ONEMAT
ExtMethod = REQUEST.QueryString("EXTMethod")
	if EXTMethod = "ALREADY-HAVE" then
		ONEMAT2 = "X" & ONEMAT & "X"
		
	end if
TWOMAT2 = TWOMAT
IntMethod = REQUEST.QueryString("INTMethod")
	if INTMethod = "ALREADY-HAVE" then
		TWOMAT2 = "X" & TWOMAT & "X"
	end if


if SPACERCOLOUR = "Grey" then 
	ONESPACCOLOUR = ONESPAC & "G"
else
	ONESPACCOLOUR = ONESPAC
End if

'Special Code added to mark all TMP/HS items in a seperate field (Note 4)
	Condition = ""
If Instr(1, ONEMAT,"TMP") <> 0 AND Instr(1, TWOMAT,"TMP") = 0 Then
Condition = ONEMAT
End IF
IF Instr(1, ONEMAT,"TMP") = 0 AND Instr(1, TWOMAT,"TMP") <> 0 Then
Condition = TWOMAT
End IF

' If only one is Heat Strengthened then Note 4 receives that Material
If Instr(1, ONEMAT,"HS") <> 0 AND Instr(1, TWOMAT,"HS") = 0 Then
Condition = ONEMAT
End IF
IF Instr(1, ONEMAT,"HS") = 0 AND Instr(1, TWOMAT,"HS") <> 0 Then
Condition = TWOMAT
End IF

' If both are Tempered Note 4 receives TMP/TMP
If Instr(1, ONEMAT,"TMP") <> 0 AND Instr(1, TWOMAT,"TMP") <> 0 Then
Condition = "TMP/TMP"
End IF
	
INPUTDATE = Date
CUSTOMER=JOB
QTY = 1
'SPACERTEXT = "QUEST WINDOW SYSTEMS INC | " & Year(now) & " " 
'SPACERTEXT moved to go with Barcode below

if NOT JOB = "" OR NOT TAG = "" then

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "INSERT INTO Z_GLASSDB ([JOB], [FLOOR], [TAG], [DEPARTMENT], [DIM X], [DIM Y], [1 MAT], [2 MAT], [1 SPAC], [BARCODE], [INPUTDATE], [REQUIREDDATE], QTY, CUSTOMER, [ORDERBY], [ORDERfor], [PO], [Extordernum], [IntOrdernum], [ExtFrom],[IntFrom], [NOTES], [AIR], [EXTMethod], [INTMethod], [Condition]) VALUES( '" & JOB & "', '" & FLOOR &  "', '" & TAG & "', '" & DEPARTMENT & "', '" & WIDTH & "', '" & HEIGHT & "', '" & ONEMAT2 & "', '" & TWOMAT2 & "', '" & ONESPACCOLOUR & "', '" & BARCODE & "', '" & INPUTDATE & "', '" & REQUIREDDATE & "', '" & QTY & "', '" & CUSTOMER & "', '" & ORDERBY & "', '" & ORDERfor & "', '" & PoNum & "', '" & ExtorderNum & "', '" & IntOrderNum & "', '" & ExtFrom & "', '" & IntFrom & "', '" & NOTES & "', '" & AIR & "', '" & EXTMethod & "', '" & INTMethod & "', '" & Condition & "')"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection


' Reads the Last Row to get the ID and then updates to the Barcode
Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "Select * from Z_GLASSDB"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

do while not rs2.eof
rs2.movelast
	BARCODE = "GT" & rs2.fields("ID")
	rs2.fields("BARCODE") = BARCODE
	rs2.fields("SPACER TEXT") = rs2.fields("ID")
	if isdate(EXTExpected) then
		rs2.fields("ExtExpected") = rs2.fields("ExtExpected")
	end if
	if isdate(IntExpected) then
		rs2.fields("IntExpected") = rs2.fields("IntExpected")
	end if
	rs2.update
rs2.movenext
loop


Added= True 
else

Added = False


end if

if ONESPAC = "" or ONESPAC = "-" then
ONESPAC = 0
End if

   
   %>
   
   
            
              <form id="enter" title="Enter New Glass Form" class="panel" name="enter" action="glassentertable.asp" method="GET" target="_self" selected="true">
              
                              
        <h2>Enter New Glass Information:</h2>
		
		 <ul id="Profiles" title="Enter Glass in Table Form" selected="true">
<%		 
IF Added = TRUE then
Response.write "<li>Glass added for " & JOB & FLOOR & "-" & TAG & "</li>"
Else
response.write "<li>Please fill in all Fields to add Record </li>"
end if		 

%>		 
		 <li><table border='1'> 
		 <tr><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th>
		 <th>Height</th><th>Required Date</th><th>Department</th><th>Order By</th><th>Order For</th><th>PO</th><th>Ext Order #</th><th>Int Order #</th><th>Gas</th>
		 </tr>

		<tr>
		<td><input class="NoMargin" type="text" name='PROJECT' id='PROJECT' size='5' value = "<% response.write JOB %>" ></td>
		<td><input class="NoMargin" type="text" name='FLOOR' id='FLOOR' size='6' value = "<% response.write FLOOR %>" ></td>
		<td><input class="NoMargin"  type="text" name='TAG' id='TAG' size='6' value = "<% response.write TAG %>" ></td>
		<td><input class="NoMargin"  type="text" name='WIDTH' id='WIDTH' size='5' value = "<% response.write WIDTH %>" ></td>
		<td><input class="NoMargin"  type="text" name='HEIGHT' id='HEIGHT' size='5' value = "<% response.write HEIGHT %>" ></td>
		<td>
		<% 
		if isDate(REQUIREDDATE) then
		ReqDateTime = RequiredDate
		else
		ReqDateTime = DateAdd("d",10,Date()) 
		end if
		%>
            <input type="text" class="NoMargin" name='REQUIREDDATE' id='REQUIREDDATE' size='10' value='<% response.write ReqDateTime %>' ></td>
		<td><select name= 'DEPARTMENT' id = 'DEPARTMENT'>
			<option value = "<% response.write DEPARTMENT %>" ><% response.write DEPARTMENT %></option> 
			<option value="Production">Production</option>
			<option value="Service">Service</option>
			<option value="Commercial">Commercial</option>
			<option value="Recut">Recut</option>
			</select></td>
		<td><select name ='orderBy'>
				<option value = "<% response.write ORDERBY %>" ><% response.write ORDERBY %></option> 
				<option value = 'Yegor'>Yegor</option>
				<option value = 'Eric'>Eric</option>
				<option value = 'Joe'>Joe</option>
				<option value = 'Ariel'>Ariel</option>
				<option value = 'Tomas'>Tomas</option>
				<option value = 'Michael'>Michael</option>
				<option value = 'WIS'>WIS</option>
			</select></td>
		<td><select name ='orderfor'>
				<option value = "<% response.write ORDERfor %>" ><% response.write ORDERfor %></option> 
				<option value = 'Arten'>Artem</option>
				<option value = 'Ellerton'>Ellerton</option>
				<option value = 'Eric'>Eric</option>
				<option value = 'George'>George</option>
				<option value = 'Hamlet'>Hamlet</option>
				<option value = 'Igor'>Igor</option>
				<option value = 'Ivan'>Ivan</option>
				<option value = 'John'>John</option>
				<option value = 'Kenny'>Kenny</option>
				<option value = 'Rob'>Rob</option>
				<option value = 'Vince'>Vince</option>
				<option value = 'Yegor'>Yegor</option>
				<option value = 'WIS'>WIS</option>
		</select></td>
		<td><input  class="NoMargin" type="text" name='PoNum' id='PoNum' size='10'  value = "<% response.write PoNum %>" ></td>
		<td><input  class="NoMargin" type="text" name='ExtOrdernum' id='ExtorderNum' size='15'  value = "<% response.write ExtorderNum %>" ></td>
		<td><input  class="NoMargin" type="text" name='IntorderNum' id='intorderNum' size='15'  value = "<% response.write IntorderNum %>" ></td>
		<td><select name ='AIR'>
				<option value = "<% response.write AIR %>" ><% response.write AIR %></option> 
				<option value = 'Argon'>Argon</option>
				<option value = 'Air'>Air</option>
				<option value = 'N/A'>N/A</option>
			</select></td>
		

		</tr>
		</table></li>
		<li><table  border='1'>
		
		
		<tr><th>Ext - Method</th><th>Exterior Glass</th><th>Spacer</th><th>Black/Grey</th><th>INT - Method</th><th>Interior Glass</th><th>Notes</th></tr>
		<tr>
		<td>
				<select name ='EXTMethod'>
					<option value = '<%response.write EXTMethod%>' selected ><%response.write EXTMethod%></option>
					<option value = 'CUT'>CUT</option>
					<option value = 'ORDER'>ORDER</option>
					<option value = 'ALREADY-HAVE'>ALREADY-HAVE</option>
                </select>
         </td>
		<td><select name="ONEMAT">
		
			<% mat = mat1 %>
			<!--#include file="QSU.inc"-->
			<% 
			' Coded this 3 times - to show Description again, despite collected value is TYPE ( USER CHOOSES DESCRIPTION, but SYSTEM NEEDS TYPE)
			rs5.filter = "Type = '" & ONEMAT & "'"
			if rs5.eof then
			else
			%>
			<option value = "<% response.write rs5("TYPE") %>" selected><% response.write rs5("DESCRIPTION") %></option> 
			<%
			end if
			%>
			</select></td>
		<td><select name="ONESPAC">
			<% mat = spac1 %>
			<!--#include file="QSU2.inc"-->
			<%
			rs6.movefirst
			response.write ONESPAC
			response.write ONESPAC
			rs6.filter = "SPACER = '" & ONESPAC & "'"
			if rs6.eof then
			else
			%>
			<option value = "<% response.write rs6("SPACER") %>" selected><% response.write rs6("OT") %></option> 
			<%
			end if
			%>
			</select></td>
		<td><select name ='SPACERCOLOUR'>
		<% if SPACERCOLOUR = "" Then
		else
		%>
				<option value = "<% response.write SPACERCOLOUR %>" ><% response.write SPACERCOLOUR %></option> 
		<%
		end if
		%>
				<option value = 'Black'>Black</option>
				<option value = 'Grey'>Grey</option>
		</select></td>	
		<td>
				<select name ='INTMethod'>
					<option value = '<%response.write INTMethod%>' selected ><%response.write INTMethod%></option>
					<option value = 'CUT' >CUT</option>
					<option value = 'ORDER'>ORDER</option>
					<option value = 'ALREADY-HAVE'>ALREADY-HAVE</option>
                </select>
         </td>	
		<td><select name="TWOMAT">
			<% mat = mat1 %>
			<!--#include file="QSU.inc"-->
			<% 
			rs5.filter = "Type = '" & TWOMAT & "'"
			if rs5.eof then
			else
			%>
			<option value = "<% response.write rs5("TYPE") %>" selected><% response.write rs5("DESCRIPTION") %></option> 
			<%
			End if
			%>
			</select></td>
			<td><input  class="NoMargin" type="text" name='NOTES' id='NOTES' size='30'  value = "<% response.write NOTES %>" ></td>
		</tr>
         </table></li>   
		<li><table  border='1'>
		<tr><th>Ext Expected Date</th><th>Ext Glass From</th><th>Int Expected Date</th><th>Int Glass From</th></tr>
		<tr>		
<% 
		if isDate(ExtExpected) or ExtExpected = "" then
		ExtDateTime = ExtExpected
		else
		ExtDateTime = DateAdd("d",10,Date()) 
		end if
%>		
		 <td><input type="text" name='ExtExpected' id='ExtExpected' size='10' value='<% response.write ExtDateTime %>'></td>
		 <td><select name ='ExtFrom'>
				<option value = "<% response.write ExtFrom %>" ><% response.write ExtFrom %></option> 
				<option value = 'Quest' >Quest</option>
				<option value = 'Cardinal'>Cardinal</option>
				<option value = 'TruLite'>TruLite</option>
				<option value = 'Woodbridge'>Woodbridge</option>
				<option value = 'Other'>Other</option>
                </select></td>
<% 
		if isDate(IntExpected) or IntExpected = "" then
		IntDateTime = IntExpected
		else
		IntDateTime = DateAdd("d",10,Date()) 
		end if
%>		
		 <td><input type="text" name='IntExpected' id='IntExpected' size='10' value='<% response.write IntDateTime %>' ></td>
		 <td><select name ='IntFrom'>
				<option value = "<% response.write IntFrom %>" ><% response.write IntFrom %></option> 
				<option value = 'Quest' >Quest</option>
				<option value = 'Cardinal'>Cardinal</option>
				<option value = 'TruLite'>TruLite</option>
				<option value = 'Woodbridge'>Woodbridge</option>
				<option value = 'Other'>Other</option>
                </select></td>
		</tr><table></li>	 
		
        <br>    
         <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
        <br>
<%

if PoNum = "" then
response.write "<li class='group'>No Current PO</li>"
else
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * from Z_GLASSDB where [PO] = '" & PoNum & "' order by ID DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

		response.write "<li class='group'>GLASS Entered on Current PO: " & PoNum & "</li>"
		response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
		response.write "<li><table border='1' class='sortable'><tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>Spacer</th><th>2 Mat</th><th>Input Date</th><th>Required Date</th><th>Department</th><th>Order By</th><th>Order For</th><th>PO</th><th>Gas</th><th>Ext Work #</th><th>Int Work #</th></tr>"

if rs.eof then
Response.write "<tr><td>No current Items</td></tr>"
end if		
do while not rs.eof
	response.write "<tr><td>" & RS("ID") & "</td><td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TAG") & "</td><td>" & RS("DIM X") & "''</td><td>" & RS("DIM Y") & "''</td><td>" & RS("1 MAT") & "</td><td>" & RS("1 SPAC") & "</td><td>" & RS("2 MAT") & "</td>" 
	response.write "<td>" & RS("INPUTDATE") & "</td><td>" & RS("REQUIREDDATE") & "</td><td>" & RS("DEPARTMENT") & "</td><td>" & RS("ORDERBY") & "</td><td>" & RS("ORDERFor") & "</td><td>" & RS("PO") & "</td><td>" & RS("AIR") & "</td><td>" & RS("ExtorderNum") & "</td><td>" & RS("IntOrderNum") & "</td></tr>"
	
	rs.movenext
loop
response.write "</table></li>"



rs.close
set rs = nothing
end if
rs5.close
set rs5 = nothing
rs6.close
set rs6 = nothing
DBConnection.close 
set DBConnection = nothing
%>
<br>
</ul>
            </form>
                
             
               
</body>
</html>
