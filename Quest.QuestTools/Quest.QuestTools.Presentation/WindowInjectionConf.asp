<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Window Injection Tool</title>
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
If left(TAG,1) = "-" Then
Else 
	TAG = "-" & TAG
End If
EMPLOYEE = REQUEST.QueryString("EMPLOYEE")
DEPARTMENT = REQUEST.QueryString("DEPT")
INJECTDATE = REQUEST.QueryString("INJECTDATE")
If isDate(INJECTDATE) Then
	INJECTDAY = DAY(INJECTDATE)
	INJECTMONTH = MONTH(INJECTDATE)
	INJECTYEAR = YEAR(INJECTDATE)
	INJECTWEEK = DatePart("ww", INJECTDATE)
End If
TIMESTAMP = Now
INJECTTIME = Time()
BARCODE = JOB & FLOOR & TAG

If JOB ="" or EMPLOYEE = "" Then
	IsError = TRUE
	Error = "ERROR: Please fill in all the Data to add a Window"
Else

	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
			Process(true)
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select

End If

Function Process(isSQLServer)

	DBOpen DBConnection, isSQLServer

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM X_BARCODE"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection

	rs.Filter = "BARCODE = '" & BARCODE & "'"
	If NOT rs.eof Then
		IsError = TRUE
		Error = "ERROR: Barcode Already Exists"
	Else
		rs.Filter =""
		rs.AddNew
		rs.Fields("BARCODE") = BARCODE
		rs.Fields("JOB") = JOB
		rs.Fields("FLOOR") = FLOOR
		rs.Fields("TAG") = TAG
		rs.Fields("DEPT") = DEPARTMENT
		rs.Fields("EMPLOYEE") = EMPLOYEE
		rs.Fields("DATETIME") = TIMESTAMP
		rs.Fields("DAY") = INJECTDAY
		rs.Fields("MONTH") = INJECTMONTH
		rs.Fields("YEAR") = INJECTYEAR
		rs.Fields("TIME") = INJECTTIME
		rs.Fields("WEEK") = INJECTWEEK
		rs.update

	End If
	rs.close
	set rs = nothing

	DbCloseAll

End Function

%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="WindowInjectionForm.asp" target="_self">Win Form</a>
    </div>

<ul id="Report" title="Added" selected="true">
<%
If IsError = TRUE Then
%>
	<li><% response.write "Window Not Added" %></li>
	<li><% response.write ERROR %></li>
<%
Else
%>

	<li><% response.write "Window Added:" %></li>
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
End If
%>

  <a class="whiteButton" href="WindowInjectionForm.asp">ADD Another Window</a>
</ul>

</body>
</html>

<%
'rs.close
'set rs=nothing

'DBConnection.close
'set DBConnection=nothing
%>
