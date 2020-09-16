<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Injection Tool</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />

  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>
  

<% 
	
JOB = UCASE(REQUEST.QueryString("JOB"))
FLOOR = UCASE(REQUEST.QueryString("FLOOR"))
TAG = UCASE(REQUEST.QueryString("TAG"))
	if left(TAG,1) = "-" then
	else 
	TAG = "-" & TAG
	end if
POS = UCASE(REQUEST.QueryString("POS"))
TYPEG = UCASE(REQUEST.QueryString("TYPEG"))
EMPLOYEE = REQUEST.QueryString("EMPLOYEE")
DEPARTMENT = REQUEST.QueryString("DEPT")
INJECTDATE = REQUEST.QueryString("INJECTDATE")
if isDate(INJECTDATE) then
	INJECTDAY = DAY(INJECTDATE)
	INJECTMONTH = MONTH(INJECTDATE)
	INJECTYEAR = YEAR(INJECTDATE)
	INJECTWEEK = DatePart("ww", INJECTDATE)
end if
TIMESTAMP = Now
INJECTTIME = Time()
BARCODE = JOB & FLOOR & TAG & POS & TYPEG

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		If gstr_ErrMsg="" Then Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

Function Process(isSQLServer)
DBOpen DBConnection, isSQLServer

if JOB ="" or EMPLOYEE = "" then
	IsError = TRUE
	Error = "ERROR: Please fill in all the Data to add Glass"
	gstr_ErrMsg="Err"
else	
	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM X_BARCODEGA"
	rs.Cursortype = GetDBCursorTypeInsert
	rs.Locktype = GetDBLockTypeInsert
	rs.Open strSQL, DBConnection

	rs.Filter = "BARCODE = '" & BARCODE & "'"
	if NOT rs.eof then
		IsError = TRUE
		Error = "ERROR: Barcode Already Exists"
		gstr_ErrMsg="Error"
	else
	rs.Filter =""
	rs.AddNew
		rs.Fields("BARCODE") = BARCODE
		rs.Fields("JOB") = JOB
		rs.Fields("FLOOR") = FLOOR
		rs.Fields("TAG") = TAG
		rs.Fields("Position") = POS
		rs.Fields("TYpe") = TYPEG
		rs.Fields("DEPT") = DEPARTMENT
		rs.Fields("DATETIME") = TIMESTAMP
		rs.Fields("DAY") = INJECTDAY
		rs.Fields("MONTH") = INJECTMONTH
		rs.Fields("YEAR") = INJECTYEAR
		rs.Fields("TIME") = INJECTTIME
		rs.Fields("WEEK") = INJECTWEEK
		rs.Fields("Last") = "Injected"

		If GetID(isSQLServer,1) <> "" Then rs.Fields("ID") = GetID(isSQLServer,1)
		rs.update

		Call StoreID1(isSQLServer, rs.Fields("ID"))

	end if
	rs.close
	set rs = nothing
end if

DbCloseAll

End Function
%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="GlassInjectionForm.asp" target="_self">Glass Form</a>
    </div>


    
<ul id="Report" title="Added" selected="true">
<%
if IsError = TRUE then
%>
	<li><% response.write "Glass Not Added" %></li>
	<li><% response.write ERROR %></li>
<%
else
%>

	<li><% response.write "Glass Added:" %></li>
	<li><% response.write "Barcode: " & BARCODE %></li>
	<li><% response.write "Added Date: " & TIMESTAMP %></li>
	<li><% response.write "Injected Day: " & INJECTDAY  %></li>
	<li><% response.write "Injected Month: " & INJECTMONTH %></li>
    <li><% response.write "Injected Year: " & INJECTYEAR %></li>
    <li><% response.write "Employee: " & EMPLOYEE %></li>
    <li><% response.write "Department: " & DEPARTMENT %></li>
	<li><% response.write "JOB: " & JOB %></li>
	<li><% response.write "FLOOR:  " & FLOOR %></li>
    <li><% response.write "TAG: " & TAG %></li>
<%
end if
%>
	
    
  
  <a class="whiteButton" href="GlassInjectionForm.asp">Add Another Glass</a>
</ul>



</body>
</html>

<% 

'rs.close
'set rs=nothing

'DBConnection.close
'set DBConnection=nothing
%>

