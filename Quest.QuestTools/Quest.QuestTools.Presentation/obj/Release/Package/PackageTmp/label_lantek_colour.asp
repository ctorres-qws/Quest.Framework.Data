<%
Country =request.querystring("Country")
 
if	Country = "USA" then
	'Connect to QWDALLANPP for Texas Database
	Set DBConnection2 = Server.CreateObject("adodb.connection")
	DSN2 = "DRIVER={SQL Server}; "
	DSN2 = DSN2 & "SERVER=QWDALLANPP.quest.local\LANTEK; Database=LANTEK-DALLAS; UID=sa;"
	' New Connection March 2019
else 
	'Connect to LASSARD for "QUEST Database
	Set DBConnection2 = Server.CreateObject("adodb.connection")
	DSN2 = "DRIVER={SQL Server}; "
	'DSN2 = DSN2 & "SERVER=HANH-PC\LANTEK;UID=QUEST\jcash;PWD=0liv3rmill3R;Database=HANH-PC\LANTEK"
	DSN2 = DSN2 & "SERVER=LANTEKLAB-PC\LANTEK; Database=LANTEK; UID=sa;"
	' ITWORKS moved Server onto Seperate VM - LANTEKLAB-PC, May 4th, 2016
end if

DBConnection2.Open DSN2
%>



<%
'Value collected by Ariel/Han to collect Label Data from lantek Job
job_name=request.querystring("job_name")
' The include for the connection file below '

Set rs3 = Server.CreateObject("adodb.recordset")
strSQL = "select M.DIS_JobRef, M.PrdREF, M.IOrder, P.DIS_width, P.DIS_length, P.dis_udata2_prt, P.dis_udata3_prt, P.dis_udata4_prt, P.dis_udata5_prt, P.dis_udata6_prt, P.dis_udata7_prt From MMNN_MMOO_00000300 as M Inner Join PPRR_PPRR_00000100 as P on M.PrdRef = P.PrdRef where M.DIS_JobRef = '" & job_name & "' order by M.IOrder"

rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL, DBConnection2


%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Content-Language" content="en-us">
<title>Panel Label</title>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">

body { margin-top: -5px; }

body,td,th {
	font-family: Verdana, Geneva, sans-serif;
	font-weight: normal;
	font-size: 17px;
}

@media print
{
table {page-break-after:always}

}

</style>

</head>

<body  align= "center" valign ="top" link="#000000" vlink="#C0C0C0" alink="#F6F000">
<%if rs3.eof then
response.write "Empty File"
end if %> 
<% Do while not rs3.eof
%>
<table align= "center" valign ="top" width="500" cellspacing="1" cellpadding="1" Style="font-size: 190%;">
	<tr>

	<td Style="font-size: 90%;" align = 'center' valign = 'top'  ><b><%response.write RS3("DIS_UData4_PRT") %> </b></td>
	</tr>
	<td Style="font-size: 90%;" align = 'center' valign = 'top'  ><%response.write RS3("DIS_UData7_PRT") %></td>
	</tr>


	</table>
<%
rs3.MoveNext %>
<% loop %>




<%
rs3.close
set rs3=nothing
DBConnection2.close
set DBConnection2=nothing
%>

