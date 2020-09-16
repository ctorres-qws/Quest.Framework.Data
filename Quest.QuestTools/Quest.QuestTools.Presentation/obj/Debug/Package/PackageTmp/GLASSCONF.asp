<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Updated to include new Entry items and update them to the Database by Michael Bernholtz on request of Jody Cash-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />

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
	
	
'BARCODE has the Format GTXX where XX is the most recent ID

'Set rsCurrent = Server.CreateObject("adodb.recordset")
'strSQLC = "SELECT SCOPE_IDENTITY()"
'rsCurrent.Cursortype = 2
'rsCurrent.Locktype = 3
'rsCurrent.Open strSQLC, DBConnection

'ID = CINT(rsCurrent.fields("ID"))
'ID = ID+1

BARCODE = "GT" & ID

FLOOR = REQUEST.QueryString("FLOOR")
TAG = REQUEST.QueryString("TAG")
JOB = REQUEST.QueryString("PROJECT")
DEPARTMENT = REQUEST.QueryString("DEPARTMENT")
WIDTH = REQUEST.QueryString("WIDTH")
HEIGHT = REQUEST.QueryString("HEIGHT")
NOTES = REQUEST.QueryString("NOTES")
ONEMAT = REQUEST.QueryString("ONEMAT")
TWOMAT = REQUEST.QueryString("TWOMAT")
ONESPAC = REQUEST.QueryString("ONESPAC")
REQUIREDDATE = REQUEST.QueryString("REQUIREDDATE")
ORDERBY = REQUEST.QueryString("orderBy")
ORDERFor = REQUEST.QueryString("orderFor")
PoNum = REQUEST.QueryString("PoNum")
ExtorderNum = REQUEST.QueryString("ExtOrderNum")
IntorderNum = REQUEST.QueryString("IntorderNum")
AIR = REQUEST.QueryString("AIR")
ExtExpected = REQUEST.QueryString("ExtExpected")
IntExpected = REQUEST.QueryString("IntExpected")
ExtFrom = REQUEST.QueryString("ExtFrom")
IntFrom = REQUEST.QueryString("IntFrom")
SPColour = REQUEST.QueryString("SPColour")
' Added April 2015 for IVAN and SASHA
GlassFor = REQUEST.QueryString("GlassFor")

ExtMethod = REQUEST.QueryString("EXTMethod")
if EXTMethod = "ALREADY-HAVE" then
	ONEMAT = "X" & ONEMAT & "X"
end if

IntMethod = REQUEST.QueryString("INTMethod")
if INTMethod = "ALREADY-HAVE" then
	TWOMAT = "X" & TWOMAT & "X"
end if

SPACERCOLOUR = REQUEST.QueryString("SPACERCOLOUR")
if SPACERCOLOUR = "Grey" then 
	ONESPAC = ONESPAC & "G"
End if

'Added at Request of Sasha For Note 4 - TMP/HS Notice
' If only one is Tempered then Note 4 receives that Material
Note4 = ""
If Instr(1, ONEMAT,"TMP") <> 0 AND Instr(1, TWOMAT,"TMP") = 0 Then
	Note4 = ONEMAT
End IF
IF Instr(1, ONEMAT,"TMP") = 0 AND Instr(1, TWOMAT,"TMP") <> 0 Then
	Note4 = TWOMAT
End IF

' If only one is Heat Strengthened then Note 4 receives that Material
If Instr(1, ONEMAT,"HS") <> 0 AND Instr(1, TWOMAT,"HS") = 0 Then
	Note4 = ONEMAT
End IF
IF Instr(1, ONEMAT,"HS") = 0 AND Instr(1, TWOMAT,"HS") <> 0 Then
	Note4 = TWOMAT
End IF

' If both are Tempered Note 4 receives TMP/TMP
If Instr(1, ONEMAT,"TMP") <> 0 AND Instr(1, TWOMAT,"TMP") <> 0 Then
	Note4 = "TMP/TMP"
End IF

INPUTDATE = Date
CUSTOMER=JOB
QTY = 1

	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
			Process(true)
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select

Function Process(isSQLServer)

	DBOpen DBConnection, isSQLServer

	Dim str_ID

	Set rs = Server.CreateObject("adodb.recordset")
	'strSQL = "INSERT INTO Z_GLASSDB ([JOB], [FLOOR], [TAG], [DEPARTMENT], [DIM X], [DIM Y], [1 MAT], [2 MAT], [1 SPAC], [BARCODE], [INPUTDATE], [REQUIREDDATE], QTY, CUSTOMER, [ORDERBY], [ORDERFor], [PO], [ExtorderNum], [IntOrdernum], [ExtFrom],[IntFrom], [NOTES], [AIR], [CONDITION], [EXTMethod], [INTMethod], [SPColour]) VALUES( '" & JOB & "', '" & FLOOR &  "', '" & TAG & "', '" & DEPARTMENT & "', '" & WIDTH & "', '" & HEIGHT & "', '" & ONEMAT & "', '" & TWOMAT & "', '" & ONESPAC & "', '" & BARCODE & "', '" & INPUTDATE & "', '" & REQUIREDDATE & "', '" & QTY & "', '" & CUSTOMER & "', '" & ORDERBY & "', '" & ORDERFor & "', '" & PoNum & "', '" & ExtorderNum & "', '" & IntorderNum & "', '" & ExtFrom & "', '" & IntFrom & "', '" & NOTES & "', '" & AIR & "', '" & Note4 & "', '" & ExtMethod & "', '" & IntMethod & "', '" & SPColour & "')"
	'rs.Open strSQL, DBConnection
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
	rs.Fields("1 SPAC") = ONESPAC
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
	rs.Fields("GlassFor") = GlassFor 

	If GetID(isSQLServer,1) <> "" Then rs.Fields("ID") = GetID(isSQLServer,1)
	rs.Update
	str_ID = rs.Fields("ID")
	Call StoreID1(isSQLServer, str_ID)

	strSQL = "SELECT * FROM Z_GlassDB WHERE ID=" & str_ID
	rs.Close
	rs.Cursortype = GetDBCursorTypeInsert
	rs.Locktype = GetDBLockTypeInsert
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

'' Reads the Last Row to get the ID and then updates to the Barcode
'Set rs2 = Server.CreateObject("adodb.recordset")
'strSQL2 = "Select * from Z_GLASSDB"
'rs2.Cursortype = 2
'rs2.Locktype = 3
'rs2.Open strSQL2, DBConnection
'
'do while not rs2.eof
'	rs2.movelast
'	BARCODE = "GT" & rs2.fields("ID")
'	rs2.fields("BARCODE") = BARCODE
'	rs2.fields("SPACER TEXT") = rs2.fields("ID")
'	if isdate(EXTExpected) then
'		rs2.fields("ExtExpected") = rs2.fields("ExtExpected")
'	end if
'	if isdate(IntExpected) then
'		rs2.fields("IntExpected") = rs2.fields("IntExpected")
'	end if
'	rs2.update
'	rs2.movenext
'loop

'CHANGED from a read and addrow to an Insert command.

'Notes about Added Fields  
'
' DEPARTMENT = Department
' BARCODE used to be = JOB & FLOOR & TAG & "#3SV  BUT IS NOW GTXX where XX is ID

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="GLASSENTER.asp" target="_self">Enter Glass</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="Report" title="Added" selected="true">
	<li><% response.write "JOB " & JOB %></li>
	<li><% response.write "FLOOR " & FLOOR %></li>
	<li><% response.write "TAG " & TAG %></li>
	<li><% response.write "DEPARTMENT " & DEPARTMENT %></li>
	<li><% response.write "NOTES " & NOTES %></li>
	<li><% response.write "WIDTH " & WIDTH %></li>
	<li><% response.write "HEIGHT " & HEIGHT %></li>
	<li><% response.write "INPUT DATE " & INPUTDATE %></li>
	<li><% response.write "REQUIRED DATE " & REQUIREDDATE %></li>
	<li><% response.write "BARCODE " & BARCODE %></li>
	<li><% response.write "Ordered by: " & ORDERBY & " for " & ORDERfor %></li>
	<li><% response.write "Window PO: " & PoNum %></li>
	<li><% response.write "External Glass Work order: " & ExtorderNum & " From " & ExtFrom %></li>
	<li><% response.write "Internal Glass Work order: " & IntorderNum & " From " & IntFrom %></li>
	<li><% response.write "Glass For: " & GlassFor %></li>
</ul>

<%

'rs.close
'set rs=nothing

'DBConnection.close
'set DBConnection = nothing
%>

</body>
</html>



