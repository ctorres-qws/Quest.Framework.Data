<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
          
		  
<!-- Edited	September 2014 for Night Shift availability -->	  
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<!-- Request from Jody Cash on January 9th 2014 to change the Auto Refresh from 120 to 90 -->
<!-- Lev's Office runs Planttv.asp not this one -->
  <meta http-equiv="refresh" content="90" >
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

'Create a Query
  '  SQL = "Select * FROM X_BARCODE ORDER BY DATETIME DESC"
'Get a Record Set
  '  Set RS = DBConnection.Execute(SQL)

Set rs = Server.CreateObject("adodb.recordset")
SQL = "Select * FROM X_BARCODE WHERE DEPT = 'ASSEMBLY' AND DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear & " ORDER BY DATETIME DESC"
rs.Cursortype = 0
rs.Locktype = 1
rs.Open SQL, DBConnection

Set rs1 = Server.CreateObject("adodb.recordset")
SQL1 = "Select * FROM X_GLAZING WHERE DEPT = 'GLAZING' and FIRSTCOMPLETE ='TRUE' AND DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear & " ORDER BY DATETIME DESC"
rs1.Cursortype = 0
rs1.Locktype = 1
rs1.Open SQL1, DBConnection
	
Set rs6 = Server.CreateObject("adodb.recordset")
strSQL6 = "SELECT * From X_BARCODEGA WHERE DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear & " ORDER BY DATETIME DESC"
rs6.Cursortype = 0
rs6.Locktype = 1
rs6.Open strSQL6, DBConnection

Set rs7 = Server.CreateObject("adodb.recordset")
strSQL7 = "SELECT * From X_PRODTARGETS ORDER BY ID ASC"
rs7.Cursortype = 0
rs7.Locktype = 1
rs7.Open strSQL7, DBConnection

'Set rs8 = Server.CreateObject("adodb.recordset")
'strSQL8 = "SELECT * From X_BARCODEP ORDER BY DATETIME DESC"
'rs8.Cursortype = 2
'rs8.Locktype = 3
'rs8.Open strSQL8, DBConnection

'Total /Day / Night
totalg = 0
totalgd = 0
totalgn = 0

'Total /Day / Night
totala = 0
totalad = 0
totalan = 0

'Total /Day / Night
totalsu = 0
totalsud = 0
totalsun= 0

'Total /Day / Night
totalsuForel = 0
totalsuForeld = 0
totalsuForeln = 0

'Total /Day / Night
totalsuWillian = 0
totalsuWilliand = 0
totalsuWilliann= 0

'Panel Line added to the Report - from X_BARCODEP - New Table same as X_BARCODEGA
' totalP = 0
' totalPd = 0
' totalPn = 0




if chour < 3 then
	cDay = cYesterday
	cMonth = cMonthy
	cYear = cYeary
end if ' before three am code



rs.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
Do while not rs.eof

	if cDate(rs("TIME")) < cDate("6:00:00 PM") then
		if cDate(rs("TIME")) > cDate("3:00:00 AM") then
			totalad = totalad + 1
		end if
	else
		totalan = totalan + 1
	end if

rs.movenext
loop

rs.close
set rs = nothing

rs1.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
Do while not rs1.eof

	if cDate(rs1("TIME")) < cDate("6:00:00 PM") then
		if cDate(rs1("TIME")) > cDate("3:00:00 AM") then
			totalgd = totalgd + 1
		end if
	else
		totalgn = totalgn + 1
	end if

rs1.movenext
loop

rs1.close
set rs1 = nothing



rs6.filter = "DAY = " & cDay & " AND MONTH = " & cMonth & " AND YEAR = " & cYear
Do while not rs6.eof
'for the latter months or not

 'Forel Line - of glass Scanned - Added March 10th 2014


if cDate(rs6("TIME")) < cDate("6:00:00 PM") then

	'Forel Line - of glass Scanned  - Added March 10th 2014
	IF rs6("DEPT") = "Forel"  then
		totalsuForeld = totalsuForeld + 1
		totalsud = totalsud + 1
	end if

else 

	'Forel Line - of glass Scanned  - Added March 10th 2014
	IF rs6("DEPT") = "Forel"  then
		totalsuForeln = totalsuForeln + 1
		totalsun = totalsun + 1
	end if

end if




'SP is not reported in this page
'IF rs6("DEPT") = "Forel" AND rs6("TYPE") = "SP" then
'totalspForel = totalspForel + 1
'totalsp = totalsp + 1
'end if

if cDate(rs6("TIME")) < cDate("6:00:00 PM") then

	'Willain Line - of glass Scanned  - Added March 10th 2014
	IF rs6("DEPT") = "Willian" then
		totalsuWilliand = totalsuWilliand + 1
		totalsud = totalsud + 1
	end if

else 

	'Willain Line - of glass Scanned  - Added March 10th 2014
	IF rs6("DEPT") = "Willian" then
		totalsuWilliann = totalsuWilliann + 1
		totalsun = totalsun + 1
	End if
end if
 
 
  rs6.movenext
loop

rs6.close
set rs6 = nothing

DO WHILE NOT RS7.EOF

IF RS7("Type") = "DAYG" then
gfactor = rs7("Target") / 7.5
end if 

IF RS7("Type") = "DAYA" then
afactor = rs7("Target") / 7.5
end if

IF RS7("Type") = "DAYIG" then
sufactor = rs7("Target") / 7.5
end if

IF RS7("Type") = "NIGHTG" then
gfactorn = rs7("Target") / 6.5
end if 

IF RS7("Type") = "NIGHTA" then
afactorn = rs7("Target") / 6.5
end if

IF RS7("Type") = "NIGHTIG" then
sufactorn = rs7("Target")/ 6.5
end if

RS7.MOVENEXT
LOOP

rs7.close
set rs7 = nothing
'chour = 20

if chour < 18 then
totalgok = (chour - 7.5) * gfactor * 0.9
totalggood = (chour - 7.5) * gfactor 
totalaok = (chour - 7.5) * afactor * 0.9
totalagood = (chour - 7.5) * afactor 
totalsuok = (chour - 7.5) * sufactor * 0.9
totalsugood = (chour - 7.5) * sufactor
else
totalgok = (chour - 15.5) * gfactorn * 0.9
totalggood = (chour - 15.5) * gfactorn 
totalaok = (chour - 15.5) * afactorn * 0.9
totalagood = (chour - 15.5) * afactorn 
totalsuok = (chour - 15.5) * sufactorn * 0.9
totalsugood = (chour - 15.5) * sufactorn
end if

DBConnection.close
set DBConnection=nothing
%>
</head>
<body onload="startTime()" >
    

  <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Quest Window Systems" selected="true">


		

		<%
		if cDate(Time) < cDate("6:00:00 PM") then
			totalg = totalgd
			totala = totalad
			totalsu = totalsud
			totalsuForel = totalsuForeld
			totalsuWillian = totalsuWilliand
			response.write "<li class='group'>Day Shift Stats</li>"
		else
			totalg = totalgn
			totala = totalan
			totalsu = totalsun
			totalsuForel = totalsuForeln
			totalsuWillian = totalsuWilliann
			response.write "<li class='group'>Night Shift Stats</li>"
		end if
		
		
		%>
		<li class="libig"> <% response.write "GLAZING: " & totalg 
		if totalg > totalgok then
			if totalg > totalggood then
			response.write " " & "<img src='bigg.gif' alt='' width='20' height='20' />"
			else
			response.write " " & "<img src='bigy.gif' alt='' width='20' height='20' />"
			end if
		else
			response.write " " & "<img src='bigr.gif' alt='' width='20' height='20' />"
		end if
		%></li>
  		<li class="libig"><% response.write "ASSEMBLY: " & totala
		if totala > totalaok then
			if totala > totalagood then
			response.write " " & "<img src='bigg.gif' alt='' width='20' height='20' />"
			else
			response.write " " & "<img src='bigy.gif' alt='' width='20' height='20' />"
			end if
		else
			response.write " " & "<img src='bigr.gif' alt='' width='20' height='20' />"
		end if
		%></li>
        <li class="libig"><% response.write "IG UNITS: " & totalsu
		if totalsu > totalsuok then
			if totalsu > totalsugood then
			response.write " " & "<img src='bigg.gif' alt='' width='20' height='20' />"
			else
			response.write " " & "<img src='bigy.gif' alt='' width='20' height='20' />"
			end if
		else
			response.write " " & "<img src='bigr.gif' alt='' width='20' height='20' />"
		end if
		response.write " (Forel: " & totalsuforel & " - Willian: " & totalsuwillian & " ) "
		%></li>
		<%
		response.write "<li>Day Glazing " &totalgd & "</li>"
		response.write "<li>Night Glazing " & totalgn& "</li>"
		response.write "<li>Day Assembly " & totalad& "</li>"
		response.write "<li>Night Assembly " &totalan & "</li>"
		response.write "<li>Day Glass " &totalsud & "</li>"
		response.write "<li>Night Glass " & totalsun& "</li>"
%>
		
<%

wcount=0
JFCHECKID=0

%>

</body>
</html>


