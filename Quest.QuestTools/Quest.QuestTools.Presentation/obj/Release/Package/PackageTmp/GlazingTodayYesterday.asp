<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- BARCODE SHIFT page changed from Direct Database call to V_REPORT3 call - Original code saved as BarcodeShiftBackup.asp-->
<!-- Michael Bernholtz, July 28th, at Request of Shaun Levy and Jody Cash -->
		 
		 
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="120" >
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

<!--#include file="todayandyesterday.asp"-->


<% 

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_GLAZING WHERE (DAY = " & cDAY & " and MONTH = " & cMonth & " and YEAR = " & cYear & ") OR (DAY = " & cYesterday & " and MONTH = " & cMonthy & " and YEAR = " & cYeary & ")"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection
EmptyFile = FALSE
Allscans = 0
Allscansy = 0


fullglaze = 0
partialglaze = 0


fullglazey = 0
partialglazey = 0

rs.filter = " DAY = '" & cDAY & "' and MONTH = '" & cMonth & "' and YEAR = '" & cYear & "'"
Do while not rs.eof
Allscans = Allscans + 1
	if rs("FirstComplete") = "TRUE" then
		fullglaze = fullglaze + 1
	else
		partialglaze = partialglaze + 1
	end if
rs.movenext
loop
if rs.eof then
EmptyFile = TRUE
else
rs.movefirst

rs.filter = " DAY = '" & cYesterday & "' and MONTH = '" & cMonthy & "' and YEAR = '" & cYeary & "'"
Do while not rs.eof
EmptyFile = FALSE
Allscansy = Allscansy + 1
	if rs("FirstComplete") = "TRUE" then
		fullglazey = fullglazey + 1
	else
		partialglazey = partialglazey + 1
	end if
rs.movenext
loop
End if

rs.close
set rs=nothing

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "Select * FROM X_GLAZING WHERE (DAY = " & cDAY & " and MONTH = " & cMonth & " and YEAR = " & cYear & ") OR (DAY = " & cYesterday & " and MONTH = " & cMonthy & " and YEAR = " & cYeary & ") Order by Barcode"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection

ScanWindow = 0
ScanCompleteWindow = 0


ScanWindowy = 0
ScanCompleteWindowy = 0

rs2.filter = " DAY = '" & cDAY & "' and MONTH = '" & cMonth & "' and YEAR = '" & cYear & "'"
if rs2.eof then
EmptyFile = TRUE
else
Do while not rs2.eof
EmptyFile = FALSE
OldBarcode = Barcode
Barcode = rs2("Barcode")
		
	if OldBarcode = Barcode then
		if rs2("FirstComplete") = "TRUE" then
			ScanCompleteWindow = ScanCompleteWindow + 1
			ScanWindow = ScanWindow - 1
		end if
	else
		if rs2("FirstComplete") = "TRUE" then
			ScanCompleteWindow = ScanCompleteWindow + 1
		else
			ScanWindow = ScanWindow + 1
		end if
	end if
rs2.movenext
loop

rs2.movefirst

rs2.filter = " DAY = '" & cYesterday & "' and MONTH = '" & cMonthy & "' and YEAR = '" & cYeary & "'"
Do while not rs2.eof
OldBarcodey = Barcodey
Barcodey = rs2("Barcode")
		
	if OldBarcodey = Barcodey then
		if rs2("FirstComplete") = "TRUE" then
			ScanCompleteWindowy = ScanCompleteWindowy + 1
			ScanWindowy = ScanWindowy - 1
		end if
	else
		if rs2("FirstComplete") = "TRUE" then
			ScanCompleteWindowy = ScanCompleteWindowy + 1
		else
			ScanWindowy = ScanWindowy + 1
		end if
	end if
rs2.movenext
loop

End if

rs2.close
set rs2=nothing


DBConnection.close
set DBConnection=nothing
%>  


</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

	<ul id="screen1" title="Glazing Today" selected="true">

		<% if EmptyFile = True then 
			response.write "<li> No Glazing occured Today or Yesterday </li>"
			end if 
		%>
		<li class="group">Glazing Today</li>
			<li><% response.write "All Scans Today (Completed / Partial): " & Allscans & " ( " & fullglaze & " / " & partialglaze %>) </li>
			<li><% response.write "# of Windows Scanned Partially: " & ScanWindow %></li>
			<li><% response.write "# of Windows Scanned Complete: " & ScanCompleteWindow%></li>
		<li class="group">Glazing Yesterday</li>
			<li><% response.write "All Scans Yesterday (Completed / Partial): " & Allscansy & " ( " & fullglazey & " / " & partialglazey %>) </li>
			<li><% response.write "# of Windows Scanned Partially: " & ScanWindowy %></li>
			<li><% response.write "# of Windows Scanned Complete: " & ScanCompleteWindowy %></li>   
        
	</ul>
        
        





</body>
</html>



