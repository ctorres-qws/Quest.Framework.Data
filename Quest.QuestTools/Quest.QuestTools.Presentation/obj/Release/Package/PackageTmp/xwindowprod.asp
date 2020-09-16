<!--#include file="dbpath.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

		 <!-- This code was Deleted by a data connection loss and rewritten on February 26th 2015 -->
		 <!-- Job View Shows all Jobs Worked on in the last 6 weeks - Broken down into Assembly / Glazing / Glazing2-->
		 <!-- Rewritten 2017 by Harj Sandhu to speed up the program and introduce SQL capabilities -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1200" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript"> iui.animOn = true; </script>
  <script src="sorttable.js"></script>
<!-- DataTables CSS -->
<link rel="stylesheet" type="text/css" href="../DataTables-1.10.2/media/css/jquery.dataTables.css">
  
<!-- jQuery -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.js"></script>
  
<!-- DataTables -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.dataTables.js"></script>

  <script type="text/javascript">
	$(document).ready( function () {
		$('#Job').DataTable({
			"iDisplayLength": 25
		});
	});
  
  </script>
  
  
</head>

<body>
<!--#include file="todayandyesterday.asp"-->
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">

	              <li class="group">LAST 12 MONTH'S ACTIVITY</li>
				  
				 <li><table border='1' class='Job' id ='Job' ><thead><tr><th>Job</th><th>Floor</th><th>SqFt</th><th>Total Window</th><th>Glazed</th><th>Date</th><th>Assembled</th><th>Date</th><th>Forel</th><th>Date</th><th>Willian</th><th>Date</th><th>Ship</th><th>Date</th></tr></thead><tbody>
<%		
SixWeek = Date()-42
'SixWeek = DateAdd("yyyy", -1, Date())
'SixWeek = DateAdd("m", -3, Date())
SixWeek = DateAdd("m", -12, Date())
Set rs = Server.CreateObject("adodb.recordset")
Dim str_SQL
str_SQL = ""
str_SQL = str_SQL & "SELECT xWP.DateStamp, xWP.JOB, xWP.Floor, xWP.TotalWin, xWP.TotalSqFt, A.MaxDate as GlazeDate, A.Count as GlazeCount, B.MaxDate as AssembleDate, B.Count as AssembleCount, C.ShipDate as ShipDate, C.Count as ShipCount, D.ForelCount, D.ForelDate, D.WillianCount, D.WillianDate From  "
str_SQL = str_SQL & "(((X_WIN_PROD xWP "
str_SQL = str_SQL & "LEFT JOIN "
str_SQL = str_SQL & "( "
str_SQL = str_SQL & "SELECT Max(DateTime) as MaxDate, COUNT(*) as [COUNT], Job, [Floor] From X_GLAZING WHERE DEPT ='GLAZING' AND FIRSTCOMPLETE = 'TRUE' GROUP BY JOB, [Floor] "
str_SQL = str_SQL & ") A ON A.Job = xWP.Job AND A.Floor = xWP.Floor "
str_SQL = str_SQL & ") "
str_SQL = str_SQL & "LEFT JOIN "
str_SQL = str_SQL & "( "
str_SQL = str_SQL & "SELECT Max(DateTime) as MaxDate, COUNT(*) as [COUNT], Job, [Floor] From X_BARCODE WHERE DEPT ='ASSEMBLY' GROUP BY JOB, [Floor] "
str_SQL = str_SQL & ") B ON B.Job = xWP.Job AND B.Floor = xWP.Floor "
str_SQL = str_SQL & ") "
str_SQL = str_SQL & "LEFT JOIN "
str_SQL = str_SQL & "( "
str_SQL = str_SQL & "SELECT Max(MaxShipDate) as ShipDate, Count(*) as [Count], Job, Floor FROM (SELECT Max(ShipDate) as MaxShipDate, Job, Floor, Tag FROM x_Shipping GROUP BY Job, Floor, Tag) T GROUP BY Job, Floor "
str_SQL = str_SQL & ") C ON C.Job = xWP.Job AND C.Floor = xWP.Floor "
str_SQL = str_SQL & ") "
str_SQL = str_SQL & " LEFT JOIN ( "

If b_SQL_Server Then
	str_SQL = str_SQL & " SELECT Job, [Floor], SUM(Case When Dept = 'Forel' Then 1 Else 0 End) as ForelCount, SUM(Case When Dept = 'Willian' Then 1 Else 0 End) as WillianCount, MAX(Case When Dept = 'Forel' Then [DateTime] Else NULL End) ForelDate, MAX(Case When Dept = 'Willian' Then [DateTime] Else NULL End) WillianDate From X_BARCODEGA GROUP BY JOB, [Floor] "
Else
	str_SQL = str_SQL & " SELECT Job, [Floor], SUM(IIF(Dept = 'Forel', 1, 0)) as ForelCount,SUM(IIF(Dept = 'Willian', 1, 0)) as WillianCount, MAX(IIF(Dept = 'Forel', [DateTime], NULL)) as ForelDate,MAX(IIF(Dept = 'Willian', [DateTime], NULL)) as WillianDate From X_BARCODEGA GROUP BY JOB, [Floor] "
End If

str_SQL = str_SQL & " ) D ON D.Job = xWP.Job AND D.Floor = xWP.Floor "
str_SQL = str_SQL & " "
str_SQL = str_SQL & " WHERE xWP.DATESTAMP > #" & SixWeek & "# ORDER BY xWP.DATESTAMP DESC"

'strSQL = FixSQL("SELECT DateStamp, JOB, Floor,TotalWin, TotalSqFt From X_WIN_PROD where DATESTAMP > #" & SixWeek & "# order by DATESTAMP DESC")
strSQL = FixSQL(str_SQL)

'Response.Write(strSQL)

Set rs = GetDisconnectedRS(strSQL, DBConnection)

Do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("JOB") & " </td>"
	response.write "<td>" & RS("Floor") & " </td>"
	response.write "<td>" & RS("TotalSqFT") & " ft<sup>2</sup></td>"
	response.write "<td>" & RS("TotalWin") & " </td>"

' **********************************************  X_GLAZING
' Now find each floor in X_Glazing
	Glaze = 0
	GlazeDate = #01/01/1999#

	If rs("GlazeDate") & "" <> "" Then
		If rs("GlazeCount") > 0 Then
			GLAZEDate = rs("GlazeDate")
			GLAZE = rs("GlazeCount")
		End If
	End If

	response.write "<td>" & Glaze & " </td>"
	if GlazeDate = #01/01/1999# then 
		response.write "<td></td>"
	else
		response.write "<td>" & FormatDateTime(GlazeDate,2) & " </td>"
	end if

' **********************************************  X_BARCODE
' Now find each floor in X_Barcode
	Assemble = 0
	AssembleDate = #01/01/1999#

	If rs("AssembleDate") & "" <> "" Then
		If rs("AssembleCount") > 0 Then
			AssembleDate = rs("AssembleDate")
			Assemble = rs("AssembleCount")
		End If
	End If

	response.write "<td>" & Assemble & " </td>"
	if AssembleDate = #01/01/1999# then 
		response.write "<td></td>"
	else
		response.write "<td>" & FormatDateTime(AssembleDate,2) & " </td>"
	end if

' **********************************************  X_BARCODEGA
' Now find each floor in X_BarcodeGA for Window Sealing

	Forel = 0
	ForelDate = #01/01/1999#
	Willian = 0
	WillianDate = #01/01/1999#

	If rs("ForelDate") & "" <> "" Then
		ForelDate = rs("ForelDate")
		Forel = rs("ForelCount")
	End If

	If rs("WillianDate") & "" <> "" Then
		WillianDate = rs("WillianDate")
		Willian = rs("WillianCount")
	End If

	response.write "<td>" & Forel & " </td>"
		if ForelDate = #01/01/1999# then 
		response.write "<td></td>"
	else
		response.write "<td>" & FormatDateTime(ForelDate,2) & " </td>"
	end if	
	response.write "<td>" & Willian & " </td>"
		if WillianDate = #01/01/1999# then 
		response.write "<td></td>"
	else
		response.write "<td>" & FormatDateTime(WillianDate,2) & " </td>"
	end if	

' **********************************************  X_SHIPPING
	Dim str_ShipCount, str_ShipDate
	str_ShipCount = "0": str_ShipDate = ""
	str_ShipCount = 0
	str_ShipDate = #01/01/1999#

	If rs("ShipDate") & "" <> "" Then
		If rs("ShipCount") > 0 Then
			str_ShipDate = rs("ShipDate")
			str_ShipCount = rs("ShipCount")
		End If
	End If

	response.write "<td>" & str_ShipCount & " </td>"
	if str_ShipDate = #01/01/1999# then 
		response.write "<td></td>"
	else
		response.write "<td>" & FormatDateTime(str_ShipDate,2) & " </td>"
	end if

	response.write "</tr>"

rs.movenext
loop

rs.close
set rs=nothing
'rs_Ship.Close: Set rs_Ship = Nothing
'rs_Glazing.Close: Set rs_Glazing = Nothing
'rs_Barcode.Close: Set rs_Barcode = Nothing
'rs_BarcodeGA.Close: Set rs_BarcodeGA = Nothing

DBConnection.close
set DBConnection = nothing

%>
</tbody></table>
</ul>

</body>
</html>
