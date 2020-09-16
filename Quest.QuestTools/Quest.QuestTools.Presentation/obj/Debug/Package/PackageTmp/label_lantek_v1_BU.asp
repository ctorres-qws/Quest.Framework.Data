<%
'Connect to LASSARD for "QUEST Database
Set DBConnection2 = Server.CreateObject("adodb.connection")
DSN2 = "DRIVER={SQL Server}; "
'DSN2 = DSN2 & "SERVER=HANH-PC\LANTEK;UID=QUEST\jcash;PWD=0liv3rmill3R;Database=HANH-PC\LANTEK"
DSN2 = DSN2 & "SERVER=LANTEKLAB-PC\LANTEK; Database=LANTEK; UID=sa;"
' ITWORKS moved Server onto Seperate VM - LANTEKLAB-PC, May 4th, 2016
'Lantek updated to change Job / FLoor locations for copy - March 2017 - For Hanh Nguyen
'This is the old version, saved in case of changes again and need to go back to old version

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

body { margin-top: 5px; }

body,td,th {
	font-family: Verdana, Geneva, sans-serif;
	font-weight: normal;
	font-size: 18px;
}

<!--@media print
{
table {page-break-after:always}
}
-->
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
	<td Style="font-size: 80%;" align = 'center' valign = 'top' colspan = "3" ><b><%response.write RS3("DIS_UData7_PRT") %> </b></td>
	</tr>
	<% 
	BarcodeSTR = RS3("PrdRef")
	Barcode = Left (BarcodeSTR, Len(BarcodeSTR)-4)
	NewBarcode= Barcode
	Job = Left(Barcode,3)
	CommChange = FALSE
	if Left(Barcode,4) = "RVT3" then
	CommChange = TRUE 
	Job = Left(Barcode,4)
	end if
	
	if(instr(1,Barcode,"-")) >1 then
	If COMMChange = FALSE then 
	Floor = Mid(Barcode,4,Instr(1,Barcode,"-") - 4)
	ELse
	Floor = Mid(Barcode,4,Instr(1,Barcode,"-") - 4)
	end if
	
	'Break te code
	'if Floor = "9" then
	'	Floor = "14"
	'	NewTag = Right(NewBarcode, Len(NewBarcode)-4)
	'	NewBarcode = JOB & Floor & NewTag
	'	
	'end if
	
	
	
	
	else
	Floor = "C1"
	JOB = "COM"
		if len(RS3("DIS_UData6_PRT"))<2 then
			NewBarcode = JOB & Floor & "-" & Barcode
		else
			NewBarcode = JOB & Floor & "-" & RS3("DIS_UData6_PRT")
		end if
	End if
	%>
	<tr><td>Job: <b><%response.write Job %></b></td><td align = 'center' valign = 'center'>Floor: <b><%response.write Floor %></b></td>
	
	<%
			if len(Barcode) >12 then
				if COMMChange = FALSE then
				Shortened = Right(Barcode, len(barcode)-(len(joB)+len(Floor)+1))
				else
				Shortened = Right(Barcode, len(barcode)-(len(joB)+len(Floor)))
				end if
			
			else 
			Shortened = barcode
			end if
	%>
	
	
<td Style="font-size: 45%"><% response.write UCASE(Shortened) %></td>
	
<!-- TO break the code:    <td Style="font-size: 45%"><% response.write UCASE(NewBarcode) %></td>-->
	
	</tr>

	<tr><td>H: <%response.write Round(RS3("DIS_Width")/25.4,2) %></td><td align = 'center' valign = 'center'>W: <%response.write Round(RS3("DIS_Length")/25.4,2) %></td><td>BC<%response.write RS3("DIS_UData2_PRT") %> : <%response.write RS3("DIS_UData3_PRT") %></td></tr>
<!-- Add back ROund( ,2)-->
<!--	<tr><td>H: <%response.write RS3("DIS_Udata4_PRT") %></td><td align = 'center' valign = 'center'>W: <%response.write RS3("DIS_UData5_PRT") %></td><td>BC<%response.write RS3("DIS_UData2_PRT") %> : <%response.write RS3("DIS_UData3_PRT") %></td></tr>
-->
	
	<tr> 
		<td Style="font-size: 85%;" align= "center" ><b> Tag </b></td>
		<td align = 'center' valign = 'top' rowspan ="2" ><img align = 'center'  src="http://chart.apis.google.com/chart?cht=qr&chs=100x100&chl=<% response.write UCASE(NewBarcode) %>&chld=H|0" alt="Barcode" /></td>
		<td Style="font-size: 85%;" align= "center" ><b> Panel # </b></td>
	</tr>	
	<tr>
	
	<%
	TagFont =135
	if len(RS3("DIS_UData6_PRT")) >5 then
		TagFont = 75
	end if
	
	%>
	<td Style=" font-size: <% response.write TagFont%>%"  align = 'center' valign = 'top' ><b><%response.write RS3("DIS_UData6_PRT") %> </b></td>

	<td align = 'center' valign = 'top'  Style="font-size: 115%;"><b><%response.write RS3("IOrder") %></b></td>
	</tr>


	</table>
	<div style="page-break-after: always;"></div>
	<div>&nbsp;
	</div>
<%
rs3.MoveNext %>
<% loop %>




<%
rs3.close
set rs3=nothing
DBConnection2.close
set DBConnection2=nothing
%>

