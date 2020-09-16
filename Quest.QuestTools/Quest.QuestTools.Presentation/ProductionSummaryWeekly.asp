<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

			
			<!--BARCODERTV-email.aspx - is an E-mail report of BARCODERTV information in webmail form, it must be updated when the code here gets changed-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>

 

<!--#include file="todayandyesterday.asp"-->

</head>
<body>

	<ul id="screen1" title="Quest Dashboard" selected="true">
	<li><table>
	<tr>
	<TH>Date</TH><TH>Shipping Scan</TH><TH>Openned Trucks</TH><TH>Closed Trucks</TH>
	<TH>Full Glaze</TH><TH>Partial Glaze</TH><TH>Assembly</TH><TH>Panel</TH><TH>Awning</TH>
	<TH>Forel</TH><TH>Willian</TH><TH>Red Zipper</TH><TH>Blue Zipper</TH>	
	</tr>
	<tbody align="center">
	
	
	
<% 
	currentDate = Now
	WindowScanned = 0
	TruckOpen = 0
	TruckClosed = 0
	GlazingFull = 0
	GlazingPartial = 0
	ASSEMBLYTotal = 0
	PANEL = 0
	AwningGlaze = 0
	FOREL = 0
	Willian = 0 
	ZipperRED = 0
	ZipperBLUE = 0
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM V_REPORT1 ORDER BY ID DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

rs.filter = "DAY = " & cDay & " AND Month = " & cMonth & " AND YEAR = " & cYear
TOdayID = RS("ID")
WeekID = TodayID - 7
rs.filter = "ID > " & WeekID




Do While not rs.eof
	'Shipping
	WindowScanned = WindowScanned + rs("ShipScan")
	TruckOpen = TruckOpen + rs("TruckOpen")
	TruckClosed = TruckClosed + rs("TruckClose")
	'GLAZING
	GlazingFull = GlazingFull + rs("GlazingFull")
	GlazingPartial = GlazingPartial + rs("GlazingPartial")
	ASSEMBLYTotal = AssemblyTotal + rs("Assembly")
	PANEL = Panel + rs("Panel")
	AwningGlaze = AwningGlaze + rs("Awning")
	FOREL = Forel + RS("Forel")
	Willian = Willian + rs("Willian")
	ZipperRED = ZipperRed + rs("ZipperRED")
	ZipperBLUE = ZipperBlue + rs("ZipperBLUE")
	
%>
<tr>
<td><%response.write RS("DAY") & "/" & RS("MONTH") & "/" & RS("YEAR")%></td>
<td><%response.write rs("Shipscan")%></td>
<td><%response.write rs("TruckOpen")%></td>
<td><%response.write rs("TruckClose")%></td>
<td><%response.write rs("GlazingFull")%></td>
<td><%response.write rs("GlazingPartial")%></td>
<td><%response.write rs("Assembly")%></td>
<td><%response.write rs("Panel")%></td>
<td><%response.write rs("Awning")%></td>
<td><%response.write rs("Forel")%></td>
<td><%response.write rs("Willian")%></td>
<td><%response.write rs("ZipperRed")%></td>
<td><%response.write rs("ZipperBlue")%></td>
</tr>

<%	
	
	
	
rs.movenext
loop


%>	
	
	<tr>
<td><b>Week Total</b></td>
<TD><B><%response.write WindowScanned %></B></TD>
<TD><B><%response.write TruckOpen %></B></TD>
<TD><B><%response.write TruckClosed%></B></TD>
<TD><B><%response.write GlazingFull%></B></TD>
<TD><B><%response.write GlazingPartial%></B></TD>
<TD><B><%response.write AssemblyTotal%></B></TD>
<TD><B><%response.write Panel%></B></TD>
<TD><B><%response.write AwningGlaze%></B></TD>
<TD><B><%response.write Forel%></B></TD>
<TD><B><%response.write Willian%></B></TD>
<TD><B><%response.write ZipperRed%></B></TD>
<TD><B><%response.write ZipperBlue%></B></TD>

</tr>

	
	
	</tbody></table></li>
	<br>








			<b><u>Shipping Statistics</u></b>
		<li>Windows Scanned: <%response.write WindowScanned %> </li>
		<li>Openned Trucks: <%response.write TruckOpen %> </li>
		<li>Closed Trucks: <%response.write TruckClosed %> </li>
			<b><u>Glazing Statistics</u></b>
		<li>Glazing Full: <%response.write GlazingFull %> </li>
		<li>Glazing Partial: <%response.write GlazingPartial %> </li>
			<b><u>Assembly Statistics</u></b>
		<li>Windows Assembled: <%response.write AssemblyTotal %> </li>
		<li>Panels Bent: <%response.write Panel %> </li>
		<li>Awnings Glazed: <%response.write AwningGlaze %> </li>
			<b><u>Glass Prepared</u></b>
		<li>Forel Glass: <%response.write Forel %> </li>
		<li>Willian Glass: <%response.write Willian %> </li>
			<b><u>Rolled on Zipper</u></b>		
		<li>Rolled Red: <%response.write ZipperRed %> </li>
		<li>Rolled Blue: <%response.write ZipperBlue %> </li>
		
		
	
        </ul>
  
<% 

rs.close
set rs=nothing
DBConnection.close
set DBConnection=nothing

%>


</body>
</html>
