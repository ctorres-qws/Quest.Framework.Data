<!--#include file="dbpath.asp"-->

<!-- Shift Hardware Corner Post Kit View for Hardware Inventory to add materials to kit-->
<!-- Logic Designed by Ariel Aziza , 8 X 5 Cart to be filled Container by container -->
<!-- Each container holds the Shift Hardware for unique JobFloorTagOpening-->
<!-- Breakdown Cart piece by piece -->  
<!-- View All - Shows whole Floor of Trolley, Bin, Cart, Container -->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Shift Corner Post Kit - Full View</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />

  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
	
	
  </script>
  <script src="sorttable.js"></script>
  
  
  
  
  
  
  
<% 
JOB = request.querystring("Job")
Floor = request.querystring("Floor")

PositionX = request.querystring("PositionX")
PositionY = request.querystring("PositionY")
PositionI = request.querystring("PositionI")
SIDE = request.querystring("SIDE")


Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_SHIFTHARDWARE WHERE JOB = '" & JOB & "' AND FLOOR = '" & FLOOR & "' ORDER BY TAG ASC, OPENING ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

	TopCart = 0
	TopBin = 0
	RecordCount = 0
	Do while not rs.eof
	RecordCount = RecordCount + 1
		if rs("Cart")+0  >= TopCart then
			TopCart = rs("Cart") + 0
			if rs("bin")+0 >= TopBin then
				TopBin = rs("Bin") + 0
			end if
		end if
	rs.movenext
	Loop
	TotalTrolleyNum = RecordCount/80
	if TotalTrolleyNum > INT(TotalTrolleyNum) then
		TotalTrolleyNum = INT(TotalTrolleyNum) + 1
	end if
	

'Trolley Count used to create 3D Array FrontTrolley(Container x,Container y,TrolleyNum)
Dim FrontTrolley(8,5,6)
Dim BackTrolley(8,5,6)
firstunitFront = "False"
firstunitBack = "False"

i = 1
rs.movefirst

Do Until i > TotalTrolleyNum
	Counter = RecordNumber
	ContainerColour = "Blue"
	x = 0
	y = 0
	 
	Do Until y = 5
		x =0
		Do Until x = 8
				if not rs.eof then
					if firstunitFront = "False" then
						firstunitFront = "True"
					else
						rs.movenext
					end if
					if not rs.eof then
						FrontTrolley(x,y,i) = rs("ID")
					else
						FrontTrolley(x,y,i) = "0"
					end if
				else
					FrontTrolley(x,y,i) = "0"
					
				end if
		x = x+1
		Loop
	y = y + 1
	Loop	
		

	x = 0
	y = 0
	Do Until y = 5
		x =0
		Do Until x = 8
				if not rs.eof then
					if firstunitBack = "False" then
						firstunitBack = "True"
					else
						rs.movenext
					end if
					if not rs.eof then
						BackTrolley(x,y,i) = rs("ID")
					else
						BackTrolley(x,y,i) = "0"
					end if
				else
					BackTrolley(x,y,i) = "0"	
				end if
		x = x+1
		Loop
	y = y + 1
	Loop	
	 	 
	
i = i+1
Loop
%>

</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
                <a class="button leftButton" type="cancel" href="ShiftHardwareCornerJobFloor.asp?Job=<%response.write JOB%>&Floor=<%response.write FLOOR%>" target="_self">Shift</a>

    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">


		<li class="group">Colour</li>
		
<%


response.write "<li> Last Bin/Cart Position of the Job: " & TopBin & " /" & TopCart & "</li>"
response.write "<li> Trolley Numbers: " & TotalTrolleyNum & "</li>"

i = 1
Do Until i > TotalTrolleyNum
	Response.write "<LI>Trolley Number " & i & "</LI>"
	response.write "<table border='1'>"
	response.write "<TR><TH>1</TH><TH>2</TH><TH>3</TH><TH>4</TH><TH>5</TH><TH>6</TH><TH>7</TH><TH>8</TH></TR>"
	x = 0
	y = 0
	Do Until y = 5
		x =0
		Response.write "<TR>"
		Do Until x = 8
			Response.write "<TD bgcolor='cyan' width = '200' height ='20' >" &  FrontTrolley(x,y,i)  & "</TD>"
			rs.filter = ""
		x = x+1
		Loop
		Response.write"</TR>"
	y = y + 1
	Loop

	Response.write "</table>"

		response.write "<table border='1'>"
	response.write "<TR><TH>1</TH><TH>2</TH><TH>3</TH><TH>4</TH><TH>5</TH><TH>6</TH><TH>7</TH><TH>8</TH></TR>"
	x = 0
	y = 0
	Do Until y = 5
		x =0
		Response.write "<TR>"
		Do Until x = 8
			Response.write "<TD bgcolor='cyan' width = '200' height ='20' >" &  BackTrolley(x,y,i)  & "</TD>"
			rs.filter = ""
		x = x+1
		Loop
		Response.write"</TR>"
	y = y + 1
	Loop

	Response.write "</table>"

i = i+1
loop

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>


</body>
</html>

