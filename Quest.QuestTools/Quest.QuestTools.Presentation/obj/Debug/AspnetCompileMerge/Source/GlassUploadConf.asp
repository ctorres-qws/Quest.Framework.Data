<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Table Format of the Glass Report - Based on Template Report from  Production-->
<!-- First level Duplicate page of Glass Report Production glassreportProduction.asp- exact duplicate except for the SQL STRING-->
<!-- Created December 6th, by Michael Bernholtz - Reports split into 3 departments - SQL string does the filter-->
<!-- Updated February 2015, by Michael Bernholtz - Update to new system and show confirmation page of update -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Import</title>
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

<%
ExcelFile = gstr_FolderUploadRecords & Request.Querystring("ItemName")

%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass</a>
        </div>
        <ul id="Profiles" title="Glass Report - Commercial" selected="true">
        <li>Successful Import to Glass Tool</li>
		<li> <% response.write ExcelFile %> </li>
		<li><table border='1' class='sortable'><tr><th>Job</th><th>Floor</th><th>Tag</th><th>PO</th><th>Exterior Work Order</th><th>Interior Work Order</th><th>Entry Date</th><th>Notes</th></tr>

<%

	'First Open Access
	If gi_Mode = c_MODE_HYBRID Then
		DBConnection.Close()
		Set DBConnection = Server.CreateObject("adodb.connection")
		DBConnection.Open GetConnectionStr(false)
	End If

	Dim DBConnSQL
	Set DBConnSQL = Server.CreateObject("adodb.connection")

	SQL = "SELECT * FROM [Sheet1$] WHERE JOB <> ''"
	Set ExcelConnection = Server.createobject("ADODB.Connection")
	ExcelConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & ExcelFile & "';Extended Properties='Excel 8.0;HDR=YES;IMEX=1';"

	SET RS = Server.CreateObject("ADODB.Recordset")
	rs.Cursortype = 2
	rs.Locktype = 3
	RS.Open SQL, ExcelConnection

	' Reads the Last Row to get the ID and then updates to the Barcode
	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "Select * from Z_GLASSDB"
	rs2.Cursortype = GetDBCursorTypeInsert
	rs2.Locktype = GetDBLockTypeInsert
	rs2.Open strSQL2, DBConnection

	If gi_Mode = c_MODE_HYBRID Then
		DBConnSQL.Open GetConnectionStr(true)
		Set RS3 = Server.CreateObject("adodb.recordset")
		strSQL2 = "Select * from Z_GLASSDB WHERE ID=-1"
		RS3.Cursortype = GetDBCursorTypeInsert
		RS3.Locktype = GetDBLockTypeInsert
		RS3.Open strSQL2, DBConnSQL
	End If

	Dim str_ID
	Dim str_Barcode

	Do While Not RS.eof
' Commented out the ability to enter multiple items with the same Tag	
		Entrytimes = RS.Fields("QTY") 

		If (Entrytimes = "") or (isNull(Entrytimes)) or (Entrytimes = " ")  or (ISNumeric(Entrytimes) = FALSE) then
			Entrytimes = 1
		end if

		
		i =0

		DO Until i = CINT(Entrytimes)
			i = i +1
			Response.write "<tr>"
			Response.write "<td>" & Rs.Fields("Job") & "</td>"
			Response.write "<td>" & Rs.Fields("Floor") & "</td>"
			Response.write "<td>" & Rs.Fields("Tag") & "</td>"
			Response.write "<td>" & Rs.Fields("PO") & "</td>"
			Response.write "<td>" & RS.Fields("Work Order External") & "</td>"
			Response.write "<td>" & RS.Fields("Work Order Internal") & "</td>"
			Response.write "<td>" & Date() & "</td>"
			Response.write "<td>" & Rs.Fields("Notes") & "</td>"
			Response.write "<td>" & Rs.Fields("Glass For") & "</td>"
			Response.write "<td>" & Rs.Fields("Overall Thickness") & "</td>"
			Response.write "</tr>"

			RS2.Addnew
			RS2("JOB") = RS.Fields("Job")
			RS2("Customer") = RS.Fields("Job")
			RS2("FLOOR") = RS.Fields("Floor")
			RS2("TAG") = RS.Fields("Tag")
			RS2("Department") = RS.Fields("Department")
			RS2("DIM X") = RS.Fields("W")
			RS2("DIM Y") = RS.Fields("H")
			RS2("1 MAT") = RS.Fields("Glass 1")
			RS2("2 MAT") = RS.Fields("Glass 2")
			RS2("1 SPAC") = RS.Fields("Spacer")
			RS2("InputDate") = Date()
			RS2("OrderBy") = RS.Fields("Order By")
			RS2("OrderFor") = RS.Fields("Order For")
			RS2("GlassFor") = RS.Fields("Glass For")
			RS2("OverallThickness") = RS.Fields("Overall Thickness")
			RS2("SPCOlour") = RS.Fields("SP Colour")
			RS2("NOTES") = RS.Fields("Notes")
			RS2("ExtExpected") = RS.Fields("External Expected Date")

			If isDate(RS.Fields("Required Date")) Then
				ReqDateTime = RS.Fields("Required Date")
			Else
				ReqDateTime = DateAdd("d",10,Date()) 
			End If

			RS2("RequiredDate") = ReqDateTime
			RS2("IntExpected") = RS.Fields("Internal Expected Date")

			If RS.Fields("PO") = "" Then
				RS2("PO") = RS.Fields("Work Order External")
			Else
				RS2("PO") = RS.Fields("PO")
			End If

			RS2("ExtOrderNum") = RS.Fields("Work Order External")
			RS2("IntOrderNum") = RS.Fields("Work Order Internal")
			RS2("QTY") = 1

			'Special Code added to mark all TMP/HS items in a seperate field (Note 4)
			Note4 = ""
			If Instr(1, RS2("1 MAT"),"TMP") <> 0 AND Instr(1, RS2("2 MAT"),"TMP") = 0 Then
				Note4 = RS2("1 MAT")
			End If

			If Instr(1, RS2("1 MAT"),"TMP") = 0 AND Instr(1, RS2("2 MAT"),"TMP") <> 0 Then
				Note4 = RS2("2 MAT")
			End If

		' If only one is Heat Strengthened then Note 4 receives that Material
			If Instr(1, RS2("1 MAT"),"HS") <> 0 AND Instr(1, RS2("2 MAT"),"HS") = 0 Then
				Note4 = RS2("1 MAT")
			End If

			If Instr(1, RS2("1 MAT"),"HS") = 0 AND Instr(1, RS2("2 MAT"),"HS") <> 0 Then
				Note4 = RS2("2 MAT")
			End IF

		' If both are Tempered Note 4 receives TMP/TMP
			If Instr(1, RS2("1 MAT"),"TMP") <> 0 AND Instr(1, RS2("2 MAT"),"TMP") <> 0 Then
				Note4 = "TMP/TMP"
			End If

			RS2("Condition") = Note4

			If GetID(isSQLServer,1) <> "" Then RS2.Fields("ID") = GetID(isSQLServer,1)
			RS2.Update
			str_ID = RS2.Fields("ID")
			Call StoreID1(isSQLServer, str_ID)

			str_Barcode = "GT" & str_ID
			'RS2.Filter = "ID=" & str_ID
			'RS2("BARCODE") = str_Barcode
			'RS2("Spacer Text") = str_ID
			'RS2.Update

			RS2("BARCODE") = str_Barcode
			RS2("Spacer Text") = str_ID
			Rs2.Update

			If gi_Mode = c_MODE_HYBRID Then

				Rs3.Addnew
				RS3("JOB") = RS.Fields("Job")
				RS3("Customer") = RS.Fields("Job")
				RS3("FLOOR") = RS.Fields("Floor")
				RS3("TAG") = RS.Fields("Tag")
				RS3("Department") = RS.Fields("Department")
				RS3("DIM X") = RS.Fields("W")
				RS3("DIM Y") = RS.Fields("H")
				RS3("1 MAT") = RS.Fields("Glass 1")
				RS3("2 MAT") = RS.Fields("Glass 2")
				RS3("1 SPAC") = RS.Fields("Spacer")
				RS3("InputDate") = Date()
				RS3("OrderBy") = RS.Fields("Order By")
				RS3("OrderFor") = RS.Fields("Order For")
				RS3("GlassFor") = RS.Fields("Glass For")
				RS3("OverallThickness") = RS.Fields("Overall Thickness")
				RS3("SPCOlour") = RS.Fields("SP Colour")
				RS3("NOTES") = RS.Fields("Notes")
				RS3("ExtExpected") = RS.Fields("External Expected Date")
				RS3("RequiredDate") = ReqDateTime
				RS3("IntExpected") = RS.Fields("Internal Expected Date")

				If RS.Fields("PO") = "" Then
					RS3("PO") = RS.Fields("Work Order External")
				Else
					RS3("PO") = RS.Fields("PO")
				End If

				RS3("ExtOrderNum") = RS.Fields("Work Order External")
				RS3("IntOrderNum") = RS.Fields("Work Order Internal")
				RS3("QTY") = 1
				RS3("Condition") = Note4
				RS3("ID") = str_ID
				RS3("BARCODE") = str_Barcode
				RS3("Spacer Text") = str_ID
				RS3.Update

			End If

		Loop
		rs.movenext
	Loop

rs.close
set rs = nothing
rs2.close
set rs2 = nothing
rs3.close
set rs3 = nothing
DBConnection.close 
set DBConnection = nothing
ExcelConnection.close

%>
        </table>
    </ul>

</body>
</html>
