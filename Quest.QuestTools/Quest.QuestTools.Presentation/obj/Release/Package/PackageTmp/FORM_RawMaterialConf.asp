<!--#include file="QCdbpath.asp"--> 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 
<!-- Testing Results stored in the system - Designed for Daniel Zalcman - March 2017, Michael Bernholtz-->
<!-- Confirm Page - Entered from FORM_RawMaterialEnter.asp -->

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



<%
InDate = REQUEST.QueryString("Date")
PO= REQUEST.QueryString("PO")
BOL= REQUEST.QueryString("BOL")
If BOL = "on" then
	BOL = TRUE
Else
	BOL = FALSE
End If
Damaged= REQUEST.QueryString("Damaged")
If Damaged = "on" then
	Damaged = TRUE
Else
	Damaged = FALSE
End If
CheckedBy = REQUEST.QueryString("CheckedBy")

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

DBOpenQC DBConnection, isSQLServer

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM FORM_RawMaterial WHERE ID=-1"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.addnew
if isDate(InDate) then
	rs("Date") = InDate
else
	rs("Date") = Now
End if
rs("PO") = PO
rs("BOL") = BOL
rs("Damaged") = Damaged
rs("Checkedby") = CheckedBy

rs.update

'rs.close
'set rs=nothing

'DBConnection.close
'set DBConnection = nothing

DbCloseAll

End Function

%>
	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="FORM_RawMaterialENTER.asp" target="_self">Back</a>

    </div>


    
<ul id="Report" title="Added" selected="true">
	
    <li><% response.write "Date: " & InDate %></li>
	<li><% response.write "PO: " & PO %></li>
    <li><% response.write "Bill of Lading Confirmed?: " & BOL %></li>
	<li><% response.write "Damaged: " & Damaged %></li>
	<li><% response.write "Checked By: " & CheckedBy %></li>
	<a class="whiteButton" href="Form_RawMaterial.asp" target="_self">Confirm and Return</a>
</ul>

</body>
</html>



