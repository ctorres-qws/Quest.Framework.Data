<!--#include file="dbpath.asp"-->

<!-- Shift Hardware Kit View for Hardware Inventory to add materials to kit-->
<!-- Logic Designed by Ariel Aziza , 12 X 10 Cart to be filled Container by container -->
<!-- Each container holds the Shift Hardware for unique JobFloorTagOpening-->
<!-- Breakdown Cart into 3,4,6 spacing based on MOST SHIFT PANELS in a window per Floor -->  
<!-- View All - Shows whole Floor of Buggy, Bin, Cart, Container -->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Shift Kit - Full View</title>
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
'JOB = "AAA"
'Floor = "99"
'JOB = "ARI"
'Floor = "8"

PositionX = request.querystring("PositionX")
PositionY = request.querystring("PositionY")
PositionI = request.querystring("PositionI")
SIDE = request.querystring("SIDE")

if (PositionX + PositionY) > 1 then

'Run Code to Print current
' Include Run print and come back here
'Move 1 forward


end if 

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_SHIFTHARDWARE WHERE JOB = '" & JOB & "' AND FLOOR = '" & FLOOR & "' ORDER BY TAG ASC, OPENING ASC"

rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection



if RecordNumber < 1 then
	MaxRecords = 0
	CurrentRecords = 0
	CurrentTAG = "X"
	PreviousTAG ="X"
	TopCart = 0
	TopBin = 0
	Do while not rs.eof
		PreviousTAG = CurrentTAG
		CurrentTAG = rs("TAG")
		if PreviousTAG = CurrentTAG then
				CurrentRecords = CurrentRecords + 1
		else
				if MaxRecords < CurrentRecords then
					MaxRecords = CurrentRecords
				end if
				CurrentRecords = 1
		end if
	if rs("Cart")+0  >= TopCart then
		TopCart = rs("Cart") + 0
		if rs("bin")+0 >= TopBin then
			TopBin = rs("Bin") + 0
		end if
	end if
	
	
	
	rs.movenext
	Loop
		'Last Record Check
		if MaxRecords < CurrentRecords then
			MaxRecords = CurrentRecords
		end if


	If MaxRecords <=3 then
		RecordNumber = 3
		FBINSTART = 1
		FBINEND = 40
		FCART = 1
		BBINSTART = 1
		BBINEND = 40
		BCart =2
		TotalBuggyNum = TopCart/2
		if TotalBuggyNum > INT(TotalBuggyNum) then
			TotalBuggyNum = INT(TotalBuggyNum) + 1
		end if
	end if	
	If MaxRecords =4 then
		RecordNumber = 4
		FBINSTART = 1
		FBINEND = 30
		FCart = 1
		BBINSTART = 31
		BBINEND = 40
		BCart = 1
		TotalBuggyNum = TopCart
	end if	
	If MaxRecords >=5 then
		RecordNumber = 6
		FBINSTART = 1
		FBINEND = 20
		FCart = 1
		BBINSTART = 21
		BBINEND = 40
		BCart =1
		TotalBuggyNum = TopCart
	end if	
end if 'RecordNumber already calculated


'Buggy Count used to create 3D Array FrontBuggy(Container x,Container y,BuggyNum)
Dim FrontBuggy(12,10,6)
Dim BackBuggy(12,10,6)
i = 1
Do Until i > TotalBuggyNum


	'All Front Containers Declared for the Cart and set to "0"
	x = 0
	y = 0
	Do Until y = 10
		 x=0
		Do Until x = 12
			FrontBuggy(x,y,i) = "0"
			
		x = x+1
		Loop
	y = y + 1
	Loop

	'All Back Containers Declared for the Cart and set to "0"

	x = 0
	y = 0
	Do Until y = 10
		x =0
		Do Until x = 12
			BackBuggy(x,y,i) = "0"
			
		x = x+1
		Loop
	y = y + 1
	Loop

i = i +1
loop

' Each Opening will Require the Following information:
' Blue / Yellow
' JobFloorTagOpening
' H131LS,H131L,H131RS,H131R,H132,H132S
'Sample "Blue,AAA99-001#1,0,1,0,1,0,1"

' Step 1 is to fill the locations
' Step 2 is to create the view







'All Front Containers Set to Blue or Yellow
' Using RecordNumber

i = 1
Do Until i > TotalBuggyNum
	Counter = RecordNumber
	ContainerColour = "Yellow"
			

	x = 0
	y = 0
	BinNumber = 1
	
	if BCART > FCART AND i > 1 then
		CartNumber = (FCART * i ) + (i-1)
	else
		CartNumber = FCART * i 
	end if
	rs.filter = ""
	rs.filter = "CART = '" & CartNumber & "'"
	if not rs.eof then
		rs.movefirst
		BinCount = rs("Bin")
		CartCount  = rs("Cart")
	end if
	 
	Do Until y = 10
		x =0
		Do Until x = 12

			if (BinCount + 0 = BinNumber + 0) AND (CartCount + 0 = CartNumber + 0) then
				FrontBuggy(x,y,i) = ContainerColour & " " & rs("ID")

				if not rs.eof then
					rs.movenext
					if not rs.eof then
						BinCount = rs("BIN") + 0
						CartCount = rs("CART") + 0
					else
						BinCount = 99
					end if
					
				end if
			else
				FrontBuggy(x,y,i) = ContainerColour

			end if
			Counter = Counter - 1	
				
				
			if Counter = 0 AND ContainerColour = "Yellow"  then
				Counter = RecordNumber
				if x=11 And (RecordNumber <> 4) then
				else
				ContainerColour = "Blue"
				end if
				BinNumber = BinNumber + 1
			end if
			if Counter = 0 AND ContainerColour = "Blue" then
				Counter = RecordNumber
				
				if x=11 AND (RecordNumber <> 4)  then
				else
				ContainerColour = "Yellow"
				end if
				BinNumber = BinNumber + 1
			end if
		x = x+1
		Loop
	y = y + 1
	Loop	
		

	x = 0
	y = 0
	BinNumber = BBINSTART
	CartNumber = BCART * i
	rs.filter = ""
	rs.filter = "CART = '" & CartNumber & "' AND BIN >= '" & BBINSTART & "'"
	if not rs.eof then
		rs.movefirst
		BinCount = rs("Bin")
		CartCount  = rs("Cart")
	end if
	 	 
	Do Until y = 10
		x =0
		Do Until x = 12

			if (BinCount + 0 = BinNumber + 0) AND (CartCount + 0 = CartNumber + 0) then
				BackBuggy(x,y,i) = ContainerColour & " " & rs("ID")

				if not rs.eof then
					rs.movenext
					if not rs.eof then
						BinCount = rs("BIN") + 0
						CartCount = rs("CART") + 0
					else
						BinCount = 99
					end if
					
				end if
			else
				BackBuggy(x,y,i) = ContainerColour

			end if
			Counter = Counter - 1	
				
				
			if Counter = 0 AND ContainerColour = "Yellow"  then
				Counter = RecordNumber
				if x=11 And (RecordNumber <> 4) then
				else
				ContainerColour = "Blue"
				end if
				BinNumber = BinNumber + 1
			end if
			if Counter = 0 AND ContainerColour = "Blue" then
				Counter = RecordNumber
				
				if x=11 AND (RecordNumber <> 4)  then
				else
				ContainerColour = "Yellow"
				end if
				BinNumber = BinNumber + 1
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
                <a class="button leftButton" type="cancel" href="ShiftHardwareJobFloor.asp?Job=<%response.write JOB%>&Floor=<%response.write FLOOR%>" target="_self">Shift</a>

    </div>

<ul id="screen1" title="Quest Dashboard" selected="true">


		<li class="group">Colour</li>
		
<%


response.write "<li> Containers Per Tag: " & RecordNumber & "</li>"
response.write "<li> Last Bin/Cart Position of the Job: " & TopBin & " /" & TopCart & "</li>"
response.write "<li> Buggy Numbers: " & TotalBuggyNum & "</li>"

i = 1
Do Until i > TotalBuggyNum

	response.write "<li> Bin: " & FBINSTART & " - " & FBINEND & " </li>"
	if BCART > FCART AND i > 1 then
		
		response.write "<li> Cart: " & (FCart * i) + (i-1) & "</li>"
	else	
		response.write "<li> Cart: " & FCart * i & "</li>"
	end if
	response.write "<table border='1'>"
	response.write "<TR><TH>1</TH><TH>2</TH><TH>3</TH><TH>4</TH><TH>5</TH><TH>6</TH><TH>7</TH><TH>8</TH><TH>9</TH><TH>10</TH><TH>11</TH><TH>12</TH></TR>"
	x = 0
	y = 0
	Do Until y = 10
		x =0
		Response.write "<TR>"
		Do Until x = 12
		
			BackgroundColour = "Black"
			if Left(FrontBuggy(x,y,i),6) = "Yellow" then
				BackgroundColour = "Yellow"
				FrontBuggy(x,y,i)= Right(FrontBuggy(x,y,i), Len(FrontBuggy(x,y,i)) - 6 ) 
			end if 
			if Left(FrontBuggy(x,y,i),4) = "Blue" then
				BackgroundColour = "Cyan"
				FrontBuggy(x,y,i)= Right(FrontBuggy(x,y,i), Len(FrontBuggy(x,y,i)) -4 ) 
			end if 
		
			CellContent = ""
			if len(FrontBuggy(x,y,i)) >2 then
				rs.filter = "ID = " & trim(FrontBuggy(x,y,i))
				CellContent = rs("Barcode")
			end if
			Response.write "<TD bgcolor= " & BackgroundColour& " width = '100' height ='20' >" &  CellContent & "</TD>"
			rs.filter = ""
			
		x = x+1
		Loop
		Response.write"</TR>"
	y = y + 1
	Loop

	Response.write "</table></li>"

	response.write "<li> Bin: " & BBINSTART & " - " & BBINEND & " </li>"
	response.write "<li> Cart: " & BCART* i & "</li>"
	response.write "<table border='1'>"
	response.write "<TR><TH>1</TH><TH>2</TH><TH>3</TH><TH>4</TH><TH>5</TH><TH>6</TH><TH>7</TH><TH>8</TH><TH>9</TH><TH>10</TH><TH>11</TH><TH>12</TH></TR>"

	x = 0
	y = 0
	Do Until y = 10
		x =0
		Response.write "<TR>"
		Do Until x = 12
			
			BackgroundColour = "Black"
			if Left(BackBuggy(x,y,i),6) = "Yellow" then
				BackgroundColour = "Yellow"
				BackBuggy(x,y,i)= Right(BackBuggy(x,y,i), Len(BackBuggy(x,y,i)) - 6 ) 
			end if 
			if Left(BackBuggy(x,y,i),4) = "Blue" then
				BackgroundColour = "Cyan"
				BackBuggy(x,y,i)= Right(BackBuggy(x,y,i), Len(BackBuggy(x,y,i)) -4 )
			end if 
			
			CellContent = ""
			if len(BackBuggy(x,y,i)) >2 then
				rs.filter = "ID = " & trim(BackBuggy(x,y,i))
				CellContent = rs("Barcode")
			end if
			Response.write "<TD bgcolor= " & BackgroundColour& " width = '100' height ='20' >" & CellContent & "</TD>"
			rs.filter =""
			
		x = x+1
		Loop
		Response.write"</TR>"
	y = y + 1
	Loop

	Response.write "</table></li>"

i = i+1
loop

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>


</body>
</html>

