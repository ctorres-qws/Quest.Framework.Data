<!--#include file="dbpath.asp"-->

<!-- Shift Hardware Kit View for Hardware Inventory to add materials to kit-->
<!-- Logic Designed by Ariel Aziza , 12 X 10 Cart to be filled Container by container -->
<!-- Each container holds the Shift Hardware for unique JobFloorTagOpening-->
<!-- Breakdown Cart into 3,4,6 spacing based on MOST SHIFT PANELS in a window per Floor -->  
<!-- View1 Shows all portions - Split in Quest tools to Frame and Sash Kits -->
<!-- Printer Code runs to ShiftHardwareLabel.asp January 2019-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Shift Kit - Entry</title>
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
Ticket = request.querystring("Ticket")
JOB = request.querystring("Job")
Floor = request.querystring("Floor")
'JOB = "AAA"
'Floor = "99"
'JOB = "ARI"
'Floor = "8"

Jump = request.querystring("JUMP")
'X and Y values are 0-11 instead of 1-12, so if Jump Command is used X and Y must move back by 1
if Jump = "Jump" then
	PositionX = request.querystring("PositionX")-1
	PositionY = request.querystring("PositionY")-1
else
	PositionX = request.querystring("PositionX")+0
	PositionY = request.querystring("PositionY")+0
end if
PositionI = request.querystring("PositionI")+0
SIDE = request.querystring("SIDE")

' PositionX = 0
' PositionY = 4
' PositionI = 1
' Side = "Front"


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
	FirstID= rs("ID")
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
	
	'Response.write MaxRecords
	
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
                <a class="button leftButton" type="cancel" href="ShiftHardwareJobFloor.asp?Job=<%response.write JOB%>&Floor=<%response.write FLOOR%>" target="_self">JobFloor</a>
				
    </div>

<ul id="screen1" selected="true">


		<li class="group">Colour</li>
		
<%


 response.write "<li> " & Side & " of  Buggy number: " & PositionI & " Buggy Row " & PositionY + 1 & "</li>"
 'FrontBuggy(PositionX,PositionY,PositionI) = Right(FrontBuggy(PositionX,PositionY,PositionI), Len(FrontBuggy(PositionX,PositionY,PositionI))-7)
 'response.write "<li> Container: " & FrontBuggy(PositionX,PositionY,PositionI) & "</li>"

 'rs.filter = "ID = " & FrontBuggy(PositionX,PositionY,PositionI)

 'response.write "<li> H131L: " & rs("H-131L") & "</li>"

 rs.filter =""

if side = "Front" then

i = PositionI

'	response.write "<li> Bin: " & FBINSTART & " - " & FBINEND & " </li>"
'	if BCART > FCART AND i > 1 then
'		
'		response.write "<li> Cart: " & (FCart * i) + (i-1) & "</li>"
'	else	
'		response.write "<li> Cart: " & FCart * i & "</li>"
'	end if
	response.write "<table border='1'>"
	response.write "<TR><TH>1</TH><TH>2</TH><TH>3</TH><TH>4</TH><TH>5</TH><TH>6</TH><TH>7</TH><TH>8</TH><TH>9</TH><TH>10</TH><TH>11</TH><TH>12</TH></TR>"
	x = 0
	y = PositionY
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
			if x+0 = PositionX+0 then 

				BackgroundColour = "Green"
			end if
		
			if FrontBuggy(x,y,i) = "" then		
				Response.write "<TD bgcolor= " & BackgroundColour & " width = '100' height ='20' >" & FrontBuggy(x,y,i) & "</TD>"
			else
				rs.filter = "ID = " & FrontBuggy(x,y,i)
				if not rs.eof then
					CurrentID = rs("ID")
					response.write "<TD bgcolor= " & BackgroundColour& " width = '100' height ='20' >"
					response.write "<B>" & rs("BARCODE") & "</B>"
					response.write "</TD>"
				end if
			end if
		x = x+1
		Loop
		Response.write"</TR>"

	Response.write "</table></li>"
end if

If Side = "Back" then

i = PositionI

'	response.write "<li> Bin: " & BBINSTART & " - " & BBINEND & " </li>"
'	response.write "<li> Cart: " & BCART* i & "</li>"
	response.write "<table border='1'>"
	response.write "<TR><TH>1</TH><TH>2</TH><TH>3</TH><TH>4</TH><TH>5</TH><TH>6</TH><TH>7</TH><TH>8</TH><TH>9</TH><TH>10</TH><TH>11</TH><TH>12</TH></TR>"

	x = 0
	y= positionY
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
			
			if x+0 = PositionX+0 then 
				BackgroundColour = "Green"
			end if
			
			if BackBuggy(x,y,i) = "" then		
				Response.write "<TD bgcolor= " & BackgroundColour& " width = '100' height ='20' >" & BackBuggy(x,y,i) & "</TD>"
			else
				rs.filter = "ID = " & BackBuggy(x,y,i)
				if not rs.eof then
					CurrentID = rs("ID")
					response.write "<TD bgcolor= " & BackgroundColour& " width = '100' height ='20' >"
					response.write "<B>" & rs("BARCODE") & "</B>"
					response.write "</TD>"
				end if
			end if
			
		x = x+1
		Loop
		Response.write"</TR>"

	Response.write "</table></li>"

	end if	

 response.write "<li> Row: " & PositionY + 1 & "</li>"
 response.write "<li> Column: " & PositionX + 1 & "</li>"
 response.write "<li> Buggy Number: " & PositionI & "</li>"
 response.write "<li> Buggy Side: " & Side & "</li>"
 
	if Side = "Front" then
		ItemCheck = FrontBuggy(PositionX,PositionY,PositionI)
	else
		ItemCheck = BackBuggy(PositionX,PositionY,PositionI)
	End if
	response.write "<table border='1'>"
	response.write "<TR><TH>Part Name</TH><TH>Part Picture</TH><TH>Qty</TH><TH>Size</TH><TH>Part Name</TH><TH>Part Picture</TH><TH>Qty</TH><TH>Size</TH></TR>"

	
	if ItemCheck = "" then
		currentBarcode = "Empty"
		response.write "<TR><TD>Empty Container</TD></TR>"
	else
		rs.filter = "ID = " & ItemCheck
		CurrentID = rs("ID")
		currentBarcode = rs("Barcode")
					RealBin = rs("Bin")
					RealCart = rs("Cart")
		
		if rs("H-131LS") > 0 or rs("H-131RS") > 0 then
			response.write "<TR>"
			response.write "<TD>H-131LS</TD>"
			response.write "<TD><img src='/HardwarePics/H-131LS.png'/></TD>"
			response.write "<TD><B><font size='12'>" & rs("H-131LS") & "</font></B></TD>"
			response.write "<TD></TD>"
		

			response.write "<TD>H-131RS</TD>"
			response.write "<TD><img src='/HardwarePics/H-131RS.png'/></TD>"
			response.write "<TD><B><font size='12'>" & rs("H-131RS") & "</font></B></TD>"
			response.write "<TD></TD>"
			response.write "</TR>"	
		end if
		
		if rs("H-131L") > 0 or rs("H-131R") > 0 then
			response.write "<TR>"
			response.write "<TD>H-131L</TD>"
			response.write "<TD><img src='/HardwarePics/H-131L.png'/></TD>"
			response.write "<TD><B><font size='12'>" & rs("H-131L") & "</font></B></TD>"
			response.write "<TD></TD>"

			response.write "<TD>H-131R</TD>"
			response.write "<TD><img src='/HardwarePics/H-131R.png'/></TD>"
			response.write "<TD><B><font size='12'>" & rs("H-131R") & "</font></B></TD>"
			response.write "<TD></TD>"
			response.write "</TR>"	
		end if
		if rs("H-132") > 0 then
			response.write "<TR>"
			response.write "<TD>H-132</TD>"
			response.write "<TD><img src='/HardwarePics/H-132.png'/></TD>"
			response.write "<TD><B><font size='12'>" & rs("H-132") & "</font></B></TD>"
			response.write "<TD></TD>"
			response.write "</TR>"	
		end if
		if rs("H-132S") > 0 then
			response.write "<TR>"
			response.write "<TD>H-132S</TD>"
			response.write "<TD><img src='/HardwarePics/H-132S.png'/></TD>"
			response.write "<TD><B><font size='12'>" & rs("H-132S") & "</font></B></TD>"
			response.write "<TD></TD>"
			response.write "</TR>"	
		end if
		if rs("7505") > 0 then
			response.write "<TR>"
			response.write "<TD>7505</TD>"
			response.write "<TD><img src='/HardwarePics/7505.png'/></TD>"
			response.write "<TD><B><font size='12'>" & rs("7505") & "</font></B></TD>"
			response.write "<TD></TD>"
			response.write "</TR>"	
		end if
		if rs("7506") > 0 then
			response.write "<TR>"
			response.write "<TD>7506</TD>"
			response.write "<TD><img src='/HardwarePics/7506.png'/></TD>"
			response.write "<TD><B><font size='12'>" & rs("7506") & "</font></B></TD>"
			response.write "<TD></TD>"
			response.write "</TR>"	
		end if
		if rs("7507") > 0 then
			response.write "<TR>"
			response.write "<TD>7507</TD>"
			response.write "<TD><img src='/HardwarePics/7507.png'/></TD>"
			response.write "<TD><B><font size='12'>" & rs("7507") & "</font></B></TD>"
			response.write "<TD></TD>"
			response.write "</TR>"	
		end if
		if rs("DeadloadQTY") > 0 then
			response.write "<TR>"
			response.write "<TD>Deadload</TD>"
			
			DeadloadName = rs("DeadloadSize")
			Select Case DeadloadName
				Case 1.25
					response.write "<TD><img src='/HardwarePics/Deadload125.png'/></TD>"
				Case 3
					response.write "<TD><img src='/HardwarePics/Deadload3.png'/></TD>"
				Case 4
					response.write "<TD><img src='/HardwarePics/Deadload4.png'/></TD>"
				Case 6
					response.write "<TD><img src='/HardwarePics/Deadload6.png'/></TD>"
				Case 10
					response.write "<TD><img src='/HardwarePics/Deadload10.png'/></TD>"
				Case Else
					response.write "<TD>Error</TD>"
			End Select	
			
			response.write "<TD><B><font size='12'>" & rs("DeadloadQTY") & "</font></B></TD>"
			response.write "<TD><B><font size='12'>" & rs("DeadloadSize") & "</font></B></TD>"
			response.write "</TR>"	
		end if
		
		if rs("H-32") > 0 then
			response.write "<TR>"
			response.write "<TD>H-32</TD>"
			response.write "<TD><img src='/HardwarePics/H-32.png'/></TD>"
			response.write "<TD><B><font size='12'>" & rs("H-32") & "</font></B></TD>"
			response.write "<TD></TD>"
			response.write "</TR>"
		end if
		if rs("H-32S") > 0 then
			response.write "<TR>"
			response.write "<TD>H-32S</TD>"
			response.write "<TD><img src='/HardwarePics/H-32S.png'/></TD>"
			response.write "<TD><B><font size='12'>" & rs("H-32S") & "</font></B></TD>"
			response.write "<TD></TD>"
			response.write "</TR>"
		end if
	End if
Response.write "</table></li>"
response.write "<li> Barcode: " & currentBarcode & "</li>"

%>

<li>

<FORM METHOD="GET" ACTION="ShiftHardwareLabel.asp" target="_self">
<input type="hidden" name="Job" value= <%response.write Job%>>   
<input type="hidden" name="Floor" value= <%response.write Floor%>>   
<input type="hidden" name="PositionX" value= <%response.write PositionX %>>   
<input type="hidden" name="PositionY" value= <%response.write PositionY %>>   
<input type="hidden" name="PositionI" value= <%response.write PositionI %>> 
<input type="hidden" name="Bin" value= <%response.write RealBin %>>   
<input type="hidden" name="Cart" value= <%response.write RealCart %>>  
<input type="hidden" name="Barcode" value= <%response.write currentBarcode %>>    
<input type="hidden" name="Side" value= <%response.write Side%>>   
<input type="hidden" name="Ticket" value= "View">  
<INPUT TYPE="submit" align="Right" value="Next Position" target="_self"  style="font-size : 15px; height:50px;width:100px" ></FORM>
</li>

<%
response.write "<li> Last Bin/Cart Position of the Job: " & TopBin & " /" & TopCart & "</li>"
%>
<li>

<FORM METHOD="GET" ACTION="ShiftHardwareView1.asp" target="_self">
<input type="hidden" name="Job" value= <%response.write Job%>>
<input type="hidden" name="Floor" value= <%response.write Floor%>>  
<input type="hidden" name="Jump" value= "Jump" > 

Column<select name="PositionX" id="PositionX">

<% 
Num = 1
Do Until Num = 13 
	Response.write "<option value='"
	Response.write Num
	Response.write "'"
		If PositionX + 1 = Num then
			Response.write " selected "
		end if
	Response.write ">" & Num & "</option>"
	
Num = Num + 1
Loop
%>
</select>

Row<select name="PositionY" id="PositionY">

<% 
Num = 1
Do Until Num = 11
	Response.write "<option value='"
	Response.write Num
	Response.write "'"
		If PositionY + 1 = Num then
			Response.write " selected "
		end if
	Response.write ">" & Num & "</option>"
	
Num = Num + 1
Loop
%>
</select>

Buggy<select name="PositionI" id="PositionI">

<% 
Num = 1
Do Until Num = TotalBuggyNum + 1
	Response.write "<option value='"
	Response.write Num
	Response.write "'"
		If PositionI = Num then
			Response.write " selected "
		end if
	Response.write ">" & Num & "</option>"
	
Num = Num + 1
Loop
%>
</select>

Side<select name="Side" id="Side">
  <option value="Front"<% if Side = "Front"  then response.write "selected"  end if%> >Front</option>
  <option value="Back" <% if Side = "Back"  then response.write "selected"  end if%> >Back</option>
</select>

<INPUT TYPE="submit" align="Right" value="Jump to" target="_self" ></FORM>




<li>
<FORM METHOD="GET" ACTION="ShiftHardwareLabel.asp" target="_self">

<%

rs.filter = ""
First = request.querystring("First")
rs.filter = "ID >= " & CurrentID
if FirstID = CurrentID and First = "" then
else
rs.Movenext
end if
if rs.eof then
Response.write "<P> Last Item, No Jump to Next Available</P>"
else

NextID = rs("ID")
CheckX = 0
CheckY = 0
CheckI = 1
FoundX = 0
FoundY = 0
FoundI = 0
FoundSide = ""
Found = ""
	Do Until CheckI > TotalBuggyNum
		CheckY =0
		Do Until CheckY = 10
			CheckX =0
			Do Until CheckX = 12
			
				if Trim(FrontBuggy(CheckX,CheckY,CheckI)) & "" = NextID&"" OR  Trim(FrontBuggy(CheckX,CheckY,CheckI)) = "Blue " & NextID OR Trim(FrontBuggy(CheckX,CheckY,CheckI)) & "" = "Yellow " & NextID then
					
					FoundX = CheckX
					FoundY = CheckY
					FoundI = CheckI
					FoundSide = "Front"
				end if
				
			CheckX = CheckX+1
			Loop
		CheckY = CheckY+1
		Loop
		
		CheckY =0
		Do Until CheckY = 10
			CheckX =0
			Do Until CheckX = 12
			
				if Trim(BackBuggy(CheckX,CheckY,CheckI)) & "" = NextID&"" OR  Trim(BackBuggy(CheckX,CheckY,CheckI)) = "Blue " & NextID OR Trim(BackBuggy(CheckX,CheckY,CheckI)) & "" = "Yellow " & NextID then
					
					FoundX = CheckX
					FoundY = CheckY
					FoundI = CheckI
					FoundSide = "Back"
					
				end if
				
			CheckX = CheckX+1
			Loop
		CheckY = CheckY+1
		Loop
	CheckI= CheckI+1	
	Loop

%>
<input type="hidden" name="Job" value= <%response.write Job%>>   
<input type="hidden" name="Floor" value= <%response.write Floor%>>   
<input type="hidden" name="Barcode" value= <%response.write currentBarcode %>>    
<input type="hidden" name="PositionX" value= <%response.write PositionX %>>   
<input type="hidden" name="PositionY" value= <%response.write PositionY %>>   
<input type="hidden" name="PositionI" value= <%response.write PositionI %>>  
<input type="hidden" name="Side" value= <%response.write Side%>>  
<input type="hidden" name="NextX" value= <%response.write FoundX %>>   
<input type="hidden" name="NextY" value= <%response.write FoundY %>>   
<input type="hidden" name="NextI" value= <%response.write FoundI %>>  
<input type="hidden" name="NextSide" value= <%response.write FoundSide%>>  
<input type="hidden" name="First" value= <%response.write First%>>  
<input type="hidden" name="NextType" value="Jump">  
<input type="hidden" name="Ticket" value= "View1" >   
<INPUT TYPE="submit" align="Right" value="Next Shift" target="_self" style="font-size : 30px; height:100px;width:200px" >

<!--
<p> <% Response.write FoundX & ":" & FoundY & ":" & FoundI & ":" & FoundSide %></P>
<p> <% Response.write FirstID & ":" & NextID & ":" & CurrentID %></P>
<p> <% Response.write ":" & TRIM(FrontBuggy(9,0,1)) & ":" & ":" & CurrentID & ":"  %></P>
-->
<%
end if
%>

</FORM>
</li>


<%

	
rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>


</body>
</html>

