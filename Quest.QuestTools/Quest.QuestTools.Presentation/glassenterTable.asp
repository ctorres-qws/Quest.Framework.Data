<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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

If ONESPAC = "-" Then
	SpacerNumber = 0
	ONESPAC = ""
Else
	SpacerNumber = CInt(ONESPAC)
End If

Set rs6 = Server.CreateObject("adodb.recordset")
strSQL6 = "Select * FROM XQSU_OTSpacer"
rs6.Cursortype = 2
rs6.Locktype = 3
rs6.Open strSQL6, DBConnection
response.write SpacerNumber
rs6.filter = "SPACER = '" & SpacerNumber & "'"

If rs6.eof Then
	OVERALLTHICKNESS = 0
Else
	OVERALLTHICKNESS = rs6("OTNUM")
End If

rs6.close
set rs6 = nothing

'Spandrel Colour
SPColour = REQUEST.QueryString("SPColour")

SPACERCOLOUR = REQUEST.QueryString("SPACERCOLOUR")
REQUIREDDATE = REQUEST.QueryString("REQUIREDDATE")
ORDERBY = REQUEST.QueryString("orderBy")
ORDERFor = REQUEST.QueryString("orderFor")

GlassFor = REQUEST.QueryString("GlassFor")
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

Select Case SPACERCOLOUR
	Case "Black"
		ONESPACCOLOUR = ONESPAC
	Case "Grey"
		ONESPACCOLOUR = ONESPAC & "G"
	Case "ALUM"
		ONESPACCOLOUR = ONESPAC & "A"
	Case Else
		ONESPACCOLOUR = ONESPAC
End Select


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

If NOT JOB = "" OR NOT TAG = "" Then

	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
			Process(true)
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select

'' Reads the Last Row to get the ID and then updates to the Barcode
'	Set rs2 = Server.CreateObject("adodb.recordset")
'	strSQL2 = "Select * from Z_GLASSDB"
'	rs2.Cursortype = GetDBCursorTypeInsert
'	rs2.Locktype = GetDBLockTypeInsert
'	rs2.Open strSQL2, DBConnection
'
'	Do While Not rs2.eof
'	rs2.movelast
'		BARCODE = "GT" & rs2.fields("ID")
'		rs2.fields("BARCODE") = BARCODE
'		rs2.fields("SPACER TEXT") = rs2.fields("ID")
'		if isdate(EXTExpected) then
'			rs2.fields("ExtExpected") = rs2.fields("ExtExpected")
'		end if
'		if isdate(IntExpected) then
'			rs2.fields("IntExpected") = rs2.fields("IntExpected")
'		end if
'		rs2.update
'	rs2.movenext
'	Loop

	Added= True 
Else
	Added = False

End If

Function Process(isSQLServer)

	DBOpen DBConnection, isSQLServer

	Dim str_ID

	Set rs = Server.CreateObject("adodb.recordset")
	'strSQL = "INSERT INTO Z_GLASSDB ([JOB], [FLOOR], [TAG], [DEPARTMENT], [DIM X], [DIM Y], [1 MAT], [2 MAT], [1 SPAC], [BARCODE], [INPUTDATE], [REQUIREDDATE], QTY, CUSTOMER, [ORDERBY], [ORDERfor], [PO], [Extordernum], [IntOrdernum], [ExtFrom],[IntFrom], [NOTES], [AIR], [EXTMethod], [INTMethod], [Condition], [OverallThickness], [GlassFor], [SPColour]) VALUES( '" & JOB & "', '" & FLOOR &  "', '" & TAG & "', '" & DEPARTMENT & "', '" & WIDTH & "', '" & HEIGHT & "', '" & ONEMAT2 & "', '" & TWOMAT2 & "', '" & ONESPACCOLOUR & "', '" & BARCODE & "', '" & INPUTDATE & "', '" & REQUIREDDATE & "', '" & QTY & "', '" & CUSTOMER & "', '" & ORDERBY & "', '" & ORDERfor & "', '" & PoNum & "', '" & ExtorderNum & "', '" & IntOrderNum & "', '" & ExtFrom & "', '" & IntFrom & "', '" & NOTES & "', '" & AIR & "', '" & EXTMethod & "', '" & INTMethod & "', '" & Condition & "', '" & OverallThickness & "', '" & GlassFor & "', '" & SPColour & "')"
	strSQL = "SELECT * FROM Z_GlassDB WHERE ID=-1"
	rs.Cursortype = GetDBCursorTypeInsert
	rs.Locktype = GetDBLockTypeInsert
	rs.Open strSQL, DBConnection

	rs.AddNew
	rs.Fields("JOB") = JOB
	rs.Fields("FLOOR") = FLOOR
	rs.Fields("TAG") = TAG
	rs.Fields("DEPARTMENT") = DEPARTMENT
	rs.Fields("DIM X") = WIDTH
	rs.Fields("DIM Y") = HEIGHT
	rs.Fields("1 MAT") = ONEMAT
	rs.Fields("2 MAT") = TWOMAT
	rs.Fields("1 SPAC") = ONESPACCOLOUR
	rs.Fields("BARCODE") = BARCODE
	rs.Fields("INPUTDATE") = INPUTDATE
	rs.Fields("REQUIREDDATE") = REQUIREDDATE
	rs.Fields("QTY") = QTY
	rs.Fields("CUSTOMER") = CUSTOMER
	rs.Fields("ORDERBY") = ORDERBY
	rs.Fields("ORDERFor") = ORDERFor
	rs.Fields("PO") = PoNum
	rs.Fields("ExtorderNum") = ExtorderNum
	rs.Fields("IntOrdernum") = IntorderNum
	rs.Fields("ExtFrom") = ExtFrom
	rs.Fields("IntFrom") = GLAIntFromSSTYPE
	rs.Fields("NOTES") = NOTES
	rs.Fields("AIR") = AIR
	rs.Fields("CONDITION") = Note4
	rs.Fields("EXTMethod") = ExtMethod
	rs.Fields("INTMethod") = IntMethod
	rs.Fields("SPColour") = SPColour

	If GetID(isSQLServer,1) <> "" Then rs.Fields("ID") = GetID(isSQLServer,1)
	rs.Update
	str_ID = rs.Fields("ID")
	Call StoreID1(isSQLServer, str_ID)

	strSQL = "SELECT * FROM Z_GlassDB WHERE ID=" & str_ID
	rs.Close
	rs.Open strSQL, DBConnection

	BARCODE = "GT" & str_ID
	rs.Fields("BARCODE") = BARCODE
	rs.Fields("SPACER TEXT") = str_ID

	If isdate(EXTExpected) Then
		rs.fields("ExtExpected") = rs.fields("ExtExpected")
	End If

	If isdate(IntExpected) Then
		rs.fields("IntExpected") = rs.fields("IntExpected")
	End If

	rs.Update

	DbCloseAll

End Function

If ONESPAC = "" or ONESPAC = "-" Then
	ONESPAC = 0
End if

Set DBConnection = Server.CreateObject("adodb.connection")
DBOpen DBConnection, isSQLServer

%>

              <form id="enter" title="Enter New Glass Form" class="panel" name="enter" action="glassentertable.asp" method="GET" target="_self" selected="true">
        <h2>Enter New Glass Information:</h2>
		
		 <ul id="Profiles" title="Enter Glass in Table Form" selected="true">
<%
If Added = TRUE then
Response.write "<li>Glass added for " & JOB & FLOOR & "-" & TAG & "</li>"
Else
response.write "<li>Please fill in all Fields to add Record </li>"
End If

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
			<option value="Testing">Testing</option>
			</select></td>
		<td><select name ='orderBy'>
				<option value = "<% response.write ORDERBY %>" ><% response.write ORDERBY %></option> 
				<option value = 'Yegor'>Yegor</option>
				<option value = 'Eric'>Eric</option>
				<option value = 'Joe'>Joe</option>
				<option value = 'Ariel'>Ariel</option>
				<option value = 'Hamid'>Hamid</option>
				<option value = 'Michael'>Michael</option>
				<option value = 'Kenny'>Kenny</option>
				<option value = 'Sasha'>Sasha</option>
				<option value = 'John'>John</option>
				<option value = 'WIS'>WIS</option>
				<option value = 'QC'>QC</option>
			</select></td>
		<td><select name ='orderfor'>
				<option value = "<% response.write ORDERfor %>" ><% response.write ORDERfor %></option> 
				<option value = 'Arten'>Artem</option>
				<option value = 'Daniel'>Daniel</option>
				<option value = 'Ellerton'>Ellerton</option>
				<option value = 'Eric'>Eric</option>
				<option value = 'George'>George</option>
				<option value = 'Hamlet'>Hamlet</option>
				<option value = 'Ivan'>Ivan</option>
				<option value = 'John'>John</option>
				<option value = 'Kenny'>Kenny</option>
				<option value = 'Rob'>Rob</option>
				<option value = 'Roman'>Roman</option>
				<option value = 'Vince'>Vince</option>
				<option value = 'Yegor'>Yegor</option>
				<option value = 'WIS'>WIS</option>
				<option value = 'QC'>QC</option>
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
		
		
		<tr><th>Ext - Method</th><th>Exterior Glass</th><th>Spacer</th><th>Black/Grey</th><th>INT - Method</th><th>Interior Glass</th></tr>
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
		
			<% 
			mat = mat1 
			entertype = "Both" 
			%>
			
			<!--#include file="QSU.inc"-->
			<% 
			' Coded this 3 times - to show Description again, despite collected value is TYPE ( USER CHOOSES DESCRIPTION, but SYSTEM NEEDS TYPE)
			rs5.filter = "Type = '" & ONEMAT & "'"
			if rs5.eof then
			else
			%>
			<option value = "<% response.write rs5("TYPE") %>" selected><%  response.write rs5("TYPE") & " - " & rs5("DESCRIPTION") %></option> 
			<%
			end if
			%>
			</select></td>
		<td><select name="ONESPAC">
			<% mat = spac1 %>
			
			<!--#include file="QSU2.inc"-->
			<%
			rs6.movefirst
			rs6.filter = "SPACER = '" & ONESPAC & "'"
			if rs6.eof then
			else
			%>
			<option value = "<% response.write rs6("SPACER") %>" selected><% response.write rs6("OT") %></option> 
			<%
			end if
			%>
			</select>
			</td>
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
				<option value = 'ALUM'>ALUM</option>
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
			<option value = "<% response.write rs5("TYPE") %>" selected><% response.write rs5("TYPE") & " - " & rs5("DESCRIPTION") %></option> 
			<%
			End if
			%>
			</select></td>
		</tr>
         </table></li>   
		<li><table  border='1'>
		<tr><th>Ext Expected Date</th><th>Ext Glass From</th><th>Int Expected Date</th><th>Int Glass From</th><th>Notes</th><th>Glass For</th><th>Spandrel Colour</th></tr>
		<tr>		
<% 
		if isDate(ExtExpected) or ExtExpected = "" then
		ExtDateTime = ExtExpected
		else
		ExtDateTime = DateAdd("d",10,Date()) 
		end if
%>		
		 <td><input type="text" class="NoMargin" name='ExtExpected' id='ExtExpected' size='10' value='<% response.write ExtDateTime %>'></td>
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
		 <td><input type="text" class="NoMargin" name='IntExpected' id='IntExpected' size='10' value='<% response.write IntDateTime %>' ></td>
		 <td><select name ='IntFrom'>
				<option value = "<% response.write IntFrom %>" ><% response.write IntFrom %></option> 
				<option value = 'Quest' >Quest</option>
				<option value = 'Cardinal'>Cardinal</option>
				<option value = 'TruLite'>TruLite</option>
				<option value = 'Woodbridge'>Woodbridge</option>
				<option value = 'Other'>Other</option>
                </select></td>
				<td><input  class="NoMargin" type="text" name='NOTES' id='NOTES' size='30'  value = "<% response.write NOTES %>" ></td>
		<td><select name ='GlassFor'>
				<option value = "<% response.write GlassFor %>" ><% response.write GlassFor %></option> 
					<option value = 'SU'>SU - Sealed Unit</option>
					<option value = 'OV'>OV - Operable Vent</option>
					<option value = 'SP'>SP - Spandrel</option>
					<option value = 'SB'>SB -Shadow Box</option>
					<option value = 'SBOC'>SBOC - Shadow Box Outside Corner</option>
					<option value = 'SBIC'>SBIC - Shadow Box Inside Corner</option>
					<option value = 'SW'>SW - Swing Door/option>
					<option value = 'SD'>SD - Sliding Door</option>
					<option value = 'Sunview Door'>Sunview Door</option>
					<option value = 'OC'>OC - Outside Corner Offset</option>
					<option value = 'IC'>IC - Inside Corner Offset</option>
					<option value = 'DOS'>DOS -Double Offset</option>
			</select></td>	
		<td><select name ='SPColour'>
				<option value = "<% response.write SPColour %>" ><% response.write SPColour %></option> 
				<option value = ''></option>
				<option value = 'SP1'>SP1</option>
				<option value = 'SP2'>SP2</option>
				<option value = 'SP3'>SP3</option>
				<option value = 'SP4'>SP4</option>
				<option value = 'SP5'>SP5</option>
				
				<option value = 'SB1'>SB1</option>
				<option value = 'SB2'>SB2</option>
				<option value = 'SB3'>SB3</option>
				<option value = 'SB4'>SB4</option>
				<option value = 'SB5'>SB5</option>
				
				<option value = 'GL1'>GL1</option>
				<option value = 'GL2'>GL2</option>
				<option value = 'GL3'>GL3</option>
				<option value = 'GL4'>GL4</option>
				<option value = 'GL5'>GL5</option>


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
		response.write "<li><table border='1' class='sortable'><tr><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>Spacer</th><th>2 Mat</th><th>Input Date</th><th>Required Date</th><th>Department</th><th>Order By</th><th>Order For</th><th>PO</th><th>Gas</th><th>Ext Work #</th><th>Int Work #</th><th>Glass For</th><th>Overall Thickness</th></tr>"

if rs.eof then
Response.write "<tr><td>No current Items</td></tr>"
end if		
do while not rs.eof
	response.write "<tr><td>" & RS("ID") & "</td><td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TAG") & "</td><td>" & RS("DIM X") & "''</td><td>" & RS("DIM Y") & "''</td><td>" & RS("1 MAT") & "</td><td>" & RS("1 SPAC") & "</td><td>" & RS("2 MAT") & "</td>" 
	response.write "<td>" & RS("INPUTDATE") & "</td><td>" & RS("REQUIREDDATE") & "</td><td>" & RS("DEPARTMENT") & "</td><td>" & RS("ORDERBY") & "</td><td>" & RS("ORDERFor") & "</td><td>" & RS("PO") & "</td><td>" & RS("AIR") & "</td><td>" & RS("ExtorderNum") & "</td><td>" & RS("IntOrderNum") & "</td><td>" & RS("GlassFor") & "</td><td>" & RS("OverallThickness") & "</td></tr>"
	
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
