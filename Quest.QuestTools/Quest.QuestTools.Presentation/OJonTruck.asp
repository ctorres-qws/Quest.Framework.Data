<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Counts Windows scanned to any truck and writes openings to the x shipping table for tallying by date range -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Detailed Stats</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<!--<meta http-equiv="refresh" content="120" >-->
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


</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="ojdaterange.asp#_screen1" target="_self">OJ Date Range</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>
<!--#include file="todayandyesterday.asp"-->

<ul id="screen1" title="Quest Dashboard" selected="true">
	
<li class="group">Week's Stats</li>
    
    
<% 

datevar1 = request.querystring("D1")
datevar2 = request.querystring("D2")
'response.write "<BR><BR><BR>" & datevar1
'response.write "<BR><BR><BR>" & datevar2
'datevar1 = "2019-02-02"
'datevar2 = "2019-02-09"
weekvar1 = DatePart("ww", datevar1)
weekvar2 = DatePart("ww", datevar2)
Response.write Datevar1 & " to " & datevar2
'Response.write "(Week " & weekvar1 & " to " & weekvar2 & ")"

Set rs = Server.CreateObject("adodb.recordset")
'strSQL = "Select * FROM X_SHIPPING WHERE ID > 145955 ORDER BY JOB, FLOOR, ID ASC"
strSQL = "Select * FROM X_SHIPPING WHERE Format([ShipDate],'yyyy-mm-dd')>='" & datevar1 & "' AND Format([ShipDate],'yyyy-mm-dd')<='" & datevar2 & "' ORDER BY JOB, FLOOR, TAG ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

totalo = 0
totalj = 0
loopo = 0
loopj = 0
loopg = 0
loopa = 0
gtotalo = 0
gtotalj = 0
wcount = 0



response.write "<li><table border ='1'>"
response.write "<tr><th>Job</th><th>Windows</th><th>Openings</th><th>Openings per Window</th></tr>"
lastjob = rs("Job")

Do while not rs.eof
tablename = rs("Job")
tag = "-" & rs("Tag")
floor = rs("Floor")
job = rs("Job")
loopo = 0
loopa = 0

' Only reload table when job changes
If strJob <> rs("Job") AND strFloor <> rs("Floor") Then

	'Dim rsscroll, intRow, rsArray
    	'strSQLscroll = "SELECT * FROM [" & tablename & "]"
	'Set rsscroll = GetDisconnectedRS(strSQLscroll, DBConnection)
	'
	'SQLStr3="SELECT * FROM INFORMATION_SCHEMA.TABLES"
	'Set rs3 = GetDisconnectedRS(strSQL3, DBConnection)
	'response.write RS3("TABLE_NAME")

	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL =  "SELECT * From  [" & tablename & "] ORDER by JOB, FLOOR, TAG ASC"
	Set rs2 = GetDisconnectedRS(strSQL, DBConnection)

	

    	'rsscroll.close: set rsscroll = nothing 
       
	

End If

strJob = rs("Job")
' added floor to deal with subfloors not showing
'strFloor = rs("Floor")

rs2.filter = "[Floor] = '" & floor & "' AND [Tag] = '" & tag & "'"

if rs2.bof then
	'response.write "No Records"
	response.write "No Opening Data for: " & Job & " " & Floor & " " & Tag & "<BR>"
	'trying to figure out why these dont share data doesn't seem to be string length mismatch
	'response.write len(Job) & len(Floor) &len(Tag) &"<BR>"
	'stylez = rs2("style")
	rs("Openings") = Left(stylez,1)
	rs("First") = stylez
	'rs("Openings") = 0
	'rs("First") = 0
	
	
		
	else
		if rs("Openings") > 0 then
			wcount = wcount + 1
			stylez = rs2("style")
			'response.write rs2("Style")
			'response.write "Records"
	
			rs("Openings") = Left(stylez,1)
			rs("First") = stylez
	
				if rs("Window") = "Window" then
				loopo = Left(stylez,1)
				loopa = 1
		
				'response.write "window yes"
				'response.write rs2("Style")
				end if
		
		rs.update

		else
			wcount = wcount + 1
			stylez = rs2("style")

			if rs("Window") = "Window" then
				loopo = Left(stylez,1)
				loopa = 1

			rs("Openings") = Left(stylez,1)
			rs("First") = stylez
			RS.UPDATE
			end if


		'response.write "  already opening recorded"
		'response.write rs("Openings")

		end if
end if



if job = lastjob then
	totalo = totalo + loopo
		
	else
	'floor = rs("Floor")
	response.write "<tr><td>" & lastjob & "</td>" ' Job
	'response.write "<td>" & Floor & "</td>" 
	response.write "<td>" & wcount & "</td>"  ' Windows
	response.write "<td>" & totalo & "</td>"  ' Openings
	if aJob >0 then
		response.write "<td>" & round((totalo/aJob),1) & "</td>" ' Openings Per Window
	else
		response.write "<td>0 (No activity)</td>" ' Openings Per Window
	end if
	
	

	aJob = 0
	totalo = loopo +0


end if


 lastjob = job
 
'These two variables add up each time an opening or joint is tallyed in the loop
gtotalo = gtotalo + loopo

'These totals look at total glazed or total assembled for aggregate use in a summary stat
atotal = atotal + loopa

aJob = aJob + loopa
wcount = aJob

 
 'rs2.close
 'set rs2 = nothing
rs.movenext
loop

'this generates the last row no if statement because all looping is declared complete

response.write "<tr><td>" & lastjob & "</td>" ' Job
aJob = aJob + loopa
wcount = aJob
	'response.write "<td>" & Floor & "</td>" 
	response.write "<td>" & wcount& "</td>"  ' Windows
	response.write "<td>" & totalo & "</td>"  ' Openings
	if aJob >0 then
		response.write "<td>" & round((totalo/aJob),1) & "</td>" ' Openings Per Window
	else
		response.write "<td>0 (No activity)</td>" ' Openings Per Window
	end if
	response.write "</table></Li>"

%>

<li class="group">Week's Total Stats</li>
    <%
'gtotalo = gtotalo + loopo
atotal = atotal + loopa
wcount = aJob
	response.write "<li>TOTAL OPENINGS: " & gtotalo & "</li>"
	if atotal = 0 then
	response.write "<li>O/W: 0</li>" 
	else
	response.write "<li>O/W: " & round((gtotalo/atotal),1) & "</li>" 
	end if
	response.write "<li>Total Windows in Trucks: " & atotal & "</li>"
	
'Number of 955 is correct
'Number of 347 windows is correct
'Seems to be skipping the last window


rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>
</ul>

</body>
</html>



