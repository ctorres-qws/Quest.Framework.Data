                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Reads files as entered by Sergey for the Itemizer program and creates labels -->
<!-- Created by Michael Bernholtz for Ariel Aziza, April 2016 -->
<!-- Requires file Sergey's folder -->


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
	<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>  

	<body  align= "center" valign ="top" link="#000000" vlink="#C0C0C0" alink="#F6F000">
	<%if RS.eof then
		response.write "Empty File"
	end if %> 
    
	
	<%
ExcelFile = "UploadRecords/" & Request.Querystring("ItemName")

 
Set ExcelConnection = Server.createobject("ADODB.Connection")
ExcelConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(ExcelFile) & ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1';"


Set cat = Server.CreateObject("ADOX.Catalog")
cat.ActiveConnection = ExcelConnection
'Response.Write "Found " & cat.Tables.count & " tables in that XLS file.<p>"
For tnum = 0 To cat.Tables.count - 1
    Set tbl = cat.Tables(tnum)
   ' Response.Write "Table " & tnum & " is named " & tbl.Name & "<br/>" & vbNewLine

   SQL = "SELECT * FROM " & tbl.Name & " WHERE JOB <> ''"
SET RS = Server.CreateObject("ADODB.Recordset")
		rs.Cursortype = 2
		rs.Locktype = 3
		
		
		
RS.Open SQL, ExcelConnection





%>

	

<% Do while not RS.eof
%>
<table align= "center" valign ="top" width="420" cellspacing="1" cellpadding="1" Style="font-size: 175%;">
	<tr>
	<td Style="font-size: 100%;" align = 'center' valign = 'top' colspan = "3" ><b><%response.write RS("DIS_UData7_PRT") %> </b></td>
	</tr>
	<tr><td>Job: <b><%response.write RS("DIS_UData4_PRT") %></b></td><td align = 'center' valign = 'center'>Floor: <b><%response.write RS("DIS_UData5_PRT") %></b></td>
	<% 
	BarcodeSTR = RS("PrdRef")
	Barcode = Left (BarcodeSTR, Len(BarcodeSTR)-4)
	%>
	<td Style="font-size: 65%"><% response.write UCASE(Barcode) %></td>
	</tr>

	<tr><td>H: <b><%response.write Round(RS("DIS_Width")/25.4,2) %></b></td><td align = 'center' valign = 'center'>W: <b><%response.write Round(RS("DIS_Length")/25.4,2) %></b></td><td>BC<b><%response.write RS("DIS_UData2_PRT") %> : <%response.write RS("DIS_UData3_PRT") %></b></td></tr>

	
	<tr> 
		<td Style="font-size: 85%;" align= "center" ><b> Tag </b></td>
		<td align = 'center' valign = 'top' rowspan ="2" ><img align = 'center'  src="http://chart.apis.google.com/chart?cht=qr&chs=80x80&chl=<% response.write UCASE(Barcode) %>&chld=H|0" alt="Barcode" /></td>
		<td Style="font-size: 85%;" align= "center" ><b> Panel # </b></td>
	</tr>	
	<tr>
	<td Style="font-size: 175%;" align = 'center' valign = 'top' ><b><%response.write RS("DIS_UData6_PRT") %> </b></td>

	<td align = 'center' valign = 'top' rowspan ="2" Style="font-size: 225%;"><b><%response.write RS("IOrder") %></b></td>
	</tr>


	</table>
<%
RS.MoveNext %>
<% loop 

Next
%>




<%
ExcelConnection.close
set ExcelConnection=nothing
DBConnection.close
set DBConnection=nothing
%>
