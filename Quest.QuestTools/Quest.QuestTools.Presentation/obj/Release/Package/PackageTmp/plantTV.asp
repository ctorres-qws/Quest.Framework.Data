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
<!-- Request from Jody Cash on January 9th 2014 to change the Auto Refresh - not working in Lev's office -->
<!-- Lev's Office runs this one not Planttvm.asp-->
<!-- I have commented out the Meta Reload and set up a Javascript Reload, Michael Bernholtz, January 10 2014 -->
<!--  <meta http-equiv="refresh" content="9"; URL="http://172.18.13.31:8081/planttv.asp#_screen1" /> -->
  <link rel="stylesheet" href="/iui/iuitv.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
 




  <!--Delay Redirect Function added  {90000 = 90 Seconds} (Samsung Smart TV do not have a function for Auto Refresh as confirm in contact with a tech support representative at Samsung, January 10 2014)-->
 <script> 
delayRedirect('http://172.18.13.31:8081/planttv.asp');
function delayRedirect(url)
 {
 var Timeout = setTimeout("window.location='" + url + "'",90000);
 }
  
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


<% 

Set rs = Server.CreateObject("adodb.recordset")
SQL = "Select * FROM X_BARCODE WHERE DEPT = 'ASSEMBLY' ORDER BY DATETIME DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open SQL, DBConnection

Set rs1 = Server.CreateObject("adodb.recordset")
SQL1 = "Select * FROM X_GLAZING WHERE DEPT = 'GLAZING' and FIRSTCOMPLETE ='TRUE' ORDER BY DATETIME DESC"
rs1.Cursortype = 2
rs1.Locktype = 3
rs1.Open SQL1, DBConnection

Set rs6 = Server.CreateObject("adodb.recordset")
strSQL6 = "SELECT * From X_BARCODEGA ORDER BY DATETIME DESC"
rs6.Cursortype = 2
rs6.Locktype = 3
rs6.Open strSQL6, DBConnection

Set rs7 = Server.CreateObject("adodb.recordset")
strSQL7 = "SELECT * From X_PRODTARGETS ORDER BY ID ASC"
rs7.Cursortype = 2
rs7.Locktype = 3
rs7.Open strSQL7, DBConnection



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

%>
<!--#include file="todayandyesterday.asp"-->
<%

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
'gfactor is glazing per 17 hours

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
gfactorn = rs7("Target") / 7.5
end if 

IF RS7("Type") = "NIGHTA" then
afactorn = rs7("Target") / 7.5
end if

IF RS7("Type") = "NIGHTIG" then
sufactorn = rs7("Target")/ 7.5
end if

RS7.MOVENEXT
LOOP

rs7.close
set rs7 = nothing

DBConnection.close
set DBConnection = nothing

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

%>
</head>
<body onload="startTime();" >

  <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_webapp">Reports</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>

<ul id="screen1" title="Quest Window Systems" selected="true">


		<li class="group">Stats</li>
		<%
		if cDate(Time) < cDate("6:00:00 PM") then
			totalg = totalgd
			totala = totalad
			totalsu = totalsud
			totalsuForel = totalsuForeld
			totalsuWillian = totalsuWilliand
		else
			totalg = totalgn
			totala = totalan
			totalsu = totalsun
			totalsuForel = totalsuForeln
			totalsuWillian = totalsuWilliann
		end if
		%>
		
		<li class="libig"><% response.write "GLAZING: " & totalg 
		if totalg > totalgok then
			if totalg > totalggood then
			response.write " " & "<img src='bigg.gif' alt='' width='120' height='120' />"
			else
			response.write " " & "<img src='bigy.gif' alt='' width='120' height='120' />"
			end if
		else
			response.write " " & "<img src='bigr.gif' alt='' width='120' height='120' />"
		end if
		%></li>
  		<li class="libig"><% response.write "ASSEMBLY: " & totala
		if totala > totalaok then
			if totala > totalagood then
			response.write " " & "<img src='bigg.gif' alt='' width='120' height='120' />"
			else
			response.write " " & "<img src='bigy.gif' alt='' width='120' height='120' />"
			end if
		else
			response.write " " & "<img src='bigr.gif' alt='' width='120' height='120' />"
		end if
		%></li>
        <li class="libig"><% response.write "IG UNITS: " & totalsu
		if totalsu > totalsuok then
			if totalsu > totalsugood then
			response.write " " & "<img src='bigg.gif' alt='' width='120' height='120' />"
			else
			response.write " " & "<img src='bigy.gif' alt='' width='120' height='120' />"
			end if
		else
			response.write " " & "<img src='bigr.gif' alt='' width='120' height='120' />"
		end if
		response.write " (Forel: " & totalsuforel & " - Willian: " & totalsuwillian & " ) "
		%></li>
            
<%

wcount=0
JFCHECKID=0

%>

</body>
</html>
