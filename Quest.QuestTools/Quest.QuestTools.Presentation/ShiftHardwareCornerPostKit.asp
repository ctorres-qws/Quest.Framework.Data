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
  <title>Shift Corner Post Kit - Corner Post</title>
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
IF UCASE(RIGHT(Floor,2))="CP" then
else
Floor = Floor & "CP"
end if



if JOB = "" or Floor = "" then
EmptyLoad = TRUE
else

	Jump = request.querystring("JUMP")

	'X and Y values are 0-7 instead of 1-8, so if Jump Command is used X and Y must move back by 1
	if Jump = "Jump" then
		PositionX = request.querystring("PositionX")-1
		PositionY = request.querystring("PositionY")-1
	else
		PositionX = request.querystring("PositionX")+0
		PositionY = request.querystring("PositionY")+0
	end if
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
	if rs.bof then
	EmptyFile = TRUE
	else
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
	End if 'Empty Record?
end if
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

if EmptyLoad = TRUE or EmptyFile = TRUE then 
	if EmptyLoad = TRUE then
		response.write "<li> No Job or Floor entered </li>"
		response.write "<li> Please Go Back and Enter a Job AND Floor to Continue </li>"
	else
		response.write "<li>" & Job & Floor & " have no Corner Post items</li>"
		response.write "<li> Please Check that Job/Floor has been Processed Correctly </li>"
	end if
else
		
	Response.write "<li> " & Side & " of  Trolley number: " & PositionI & " / " & TotalTrolleyNum & " Showing Row " & PositionY + 1 & "</li>"
	
	Response.write "<table border='1'>"
	Response.write "<TR><TH>Row</TH><TH>1</TH><TH>2</TH><TH>3</TH><TH>4</TH><TH>5</TH><TH>6</TH><TH>7</TH><TH>8</TH></TR>"
	
	if Side = "Front" then
	rs.filter = ""
	x = 0
	y = PositionY
	i = PositionI

		Response.write "<TR><TD>" & PositionY + 1 & "</TD>"
		Do Until x = 8
			BackgroundColour = "Cyan"
			if x+0 = PositionX+0 then 
				BackgroundColour = "Lime"
			end if
		
			if FrontTrolley(x,y,i) = "" then		
				Response.write "<TD bgcolor= " & BackgroundColour & " width = '100' height ='20' >" & FrontTrolley(x,y,i) & "</TD>"
			else
				rs.filter = "ID = " & FrontTrolley(x,y,i)
				if not rs.eof then
					CurrentID = rs("ID")
					Response.write "<TD bgcolor= " & BackgroundColour& " width = '100' height ='20' >"
					Response.write "<B>" & rs("BARCODE") & "</B>"
					Response.write "</TD>"
				else
					Response.write "<TD bgcolor= " & BackgroundColour& " width = '100' height ='20' >"
					Response.write "<B>0</B>"
					Response.write "</TD>"
				end if
			end if
		x = x+1
		Loop
		Response.write"</TR>"
		Response.write "</table></li>"
	end if
	if Side = "Back" then
	rs.filter = ""
	x = 0
	y = PositionY
	i = PositionI

		Response.write "<TR><TD></TD>"
		Do Until x = 8
			BackgroundColour = "Cyan"
			if x+0 = PositionX+0 then 
				BackgroundColour = "Lime"
			end if
		
			if BackTrolley(x,y,i) = "" then		
				Response.write "<TD bgcolor= " & BackgroundColour & " width = '100' height ='20' >" & BackTrolley(x,y,i) & "</TD>"
			else
				rs.filter = "ID = " & BackTrolley(x,y,i)
				if not rs.eof then
					CurrentID = rs("ID")
					Response.write "<TD bgcolor= " & BackgroundColour& " width = '100' height ='20' >"
					Response.write "<B>" & rs("BARCODE") & "</B>"
					Response.write "</TD>"
				else
					Response.write "<TD bgcolor= " & BackgroundColour& " width = '100' height ='20' >"
					Response.write "<B>0</B>"
					Response.write "</TD>"
				end if
			end if
		x = x+1
		Loop
		Response.write"</TR>"
		Response.write "</table></li>"
	end if

	Response.write "<li> Row: " & PositionY + 1 & "</li>"
 Response.write "<li> Column: " & PositionX + 1 & "</li>"
 
	if Side = "Front" then
		ItemCheck = FrontTrolley(PositionX,PositionY,PositionI)
	else
		ItemCheck = BackTrolley(PositionX,PositionY,PositionI)
	End if
	
	if ItemCheck = "" or ItemCheck = 0 then
		currentBarcode = "Empty"
		Response.write "<TR><TD>Empty Container</TD></TR>"
		Response.write "<TR><TD>Deplete Job/Floor</TD></TR>"
		%>
		<!--#include file="ShiftHardwareDeplete.inc"-->
		<%
	else
		rs.filter = "ID = " & ItemCheck
		CurrentID = rs("ID")
		currentBarcode = rs("Barcode")
		Response.write "<table border='2'>"
					RealBin = rs("Bin")
					RealCart = rs("Cart")
		
		Response.write "<TR>"
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("V-32") & "</font></CENTER></B>"
				Response.write "<center>V-32</center><br>"
				Response.write "<img src='/HardwarePics/V-32.png'/>"
			Response.write "</TD>" 
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("H-32") & "</font></CENTER></B>"
				Response.write "<center>H-32</center><br>"
				Response.write "<img src='/HardwarePics/H-32.png'/>"
			Response.write "</TD>" 
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("H-32S") & "</font></CENTER></B>"
				Response.write "<center>H-32S</center><br>"
				Response.write "<img src='/HardwarePics/H-32S.png'/>"
			Response.write "</TD>" 
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("H-503") & "</font></CENTER></B>"
				Response.write "<center>H-503</center><br>"
				Response.write "<img src='/HardwarePics/H-503.png'/>"
			Response.write "</TD>" 
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("H-503B") & "</font></CENTER></B>"
				Response.write "<center>H-503B</center><br>"
				Response.write "<img src='/HardwarePics/H-503B.png'/>"
			Response.write "</TD>" 
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("H-503L") & "</font></CENTER></B>"
				Response.write "<center>H-503L</center><br>"
				Response.write "<img src='/HardwarePics/H-503L.png'/>"
			Response.write "</TD>" 
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("7506") & "</font></CENTER></B>"
				Response.write "<center>7506</center><br>"
				Response.write "<img src='/HardwarePics/7506.png'/>"
				
			Response.write "</TD>" 
		Response.write "</TR>"
		Response.write "<TR><TD></TD><TD></TD><TD></TD><TD></TD><TD></TD><TD></TD>"
				Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("7507") & "</font></CENTER></B>"
				Response.write "<center>7507</center><br>"
				Response.write "<img src='/HardwarePics/7507.png'/>"
				
			Response.write "</TD>" 

		Response.write "</TR>"
		Response.write "<TR>"
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("H-31CL") & "</font></CENTER></B>"
				Response.write "<center>H-31CL</center><br>"
				Response.write "<img src='/HardwarePics/H-31CL.png'/>"
			Response.write "</TD>" 
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("H-31CR") & "</font></CENTER></B>"
				Response.write "<center>H-31CR</center><br>"
				Response.write "<img src='/HardwarePics/H-31CR.png'/>"
				
			Response.write "</TD>" 
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("H-201-3") & "</font></CENTER></B>"
				Response.write "<center>H-201-3</center><br>"
				Response.write "<img src='/HardwarePics/H-201-3.png'/>"
				
			Response.write "</TD>" 
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("H-31CLS") & "</font></CENTER></B>"
				Response.write "<center>H-31CLS</center><br>"
				Response.write "<img src='/HardwarePics/H-31CLS.png'/>"
				
			Response.write "</TD>" 
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("H-31CRS") & "</font></CENTER></B>"
				Response.write "<center>H-31CRS</center><br>"
				Response.write "<img src='/HardwarePics/H-31CRS.png'/>"
				
			Response.write "</TD>" 
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("H-201-15") & "</font></CENTER></B>"
				Response.write "<center>H-201-15</center><br>"
				Response.write "<img src='/HardwarePics/H-201-15.png'/>"
				
			Response.write "</TD>" 
			Response.write "<TD>" 
				Response.write "<B><CENTER><font size='12'>" & rs("7505") & "</font></CENTER></B>"
				Response.write "<center>7505</center><br>"
				Response.write "<img src='/HardwarePics/7505.png'/>"
				
			Response.write "</TD>" 
		Response.write "</TR>"
						
		
	
	End if
Response.write "</table></li>"
	
%>	
<li>

<FORM METHOD="GET" ACTION="ShiftHardwareLabel.asp" target="_self">
<input type="hidden" name="Job" value= <%Response.write Job%>>   
<input type="hidden" name="Floor" value= <%Response.write Floor%>>   
<input type="hidden" name="PositionX" value= <%Response.write PositionX %>>   
<input type="hidden" name="PositionY" value= <%Response.write PositionY %>>   
<input type="hidden" name="PositionI" value= <%Response.write PositionI %>> 
<input type="hidden" name="Bin" value= <%Response.write RealBin %>>   
<input type="hidden" name="Cart" value= <%Response.write RealCart %>>  
<input type="hidden" name="Barcode" value= <%Response.write currentBarcode %>>    
<input type="hidden" name="Side" value= <%Response.write Side%>>   
<input type="hidden" name="Ticket" value= "CORNER">  
<INPUT TYPE="submit" align="Right" value="Next Position" target="_self"  style="font-size : 15px; height:50px;width:100px" ></FORM>
</li>
<li>

<FORM METHOD="GET" ACTION="ShiftHardwareCornerPostKit.asp" target="_self">
<input type="hidden" name="Job" value= <%Response.write Job%>>
<input type="hidden" name="Floor" value= <%Response.write Floor%>>  
<input type="hidden" name="Jump" value= "Jump" > 

Column<select name="PositionX" id="PositionX">

<% 
Num = 1
Do Until Num = 9
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
Do Until Num = 6
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

Trolley<select name="PositionI" id="PositionI">

<% 
Num = 1
Do Until Num = TotalTrolleyNum + 1
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
  <option value="Front"<% if Side = "Front"  then Response.write "selected"  end if%> >Front</option>
  <option value="Back" <% if Side = "Back"  then Response.write "selected"  end if%> >Back</option>
</select>

<INPUT TYPE="submit" align="Right" value="Jump to" target="_self" ></FORM>
</li>


<%
	rs.close
	set rs=nothing
end if
DBConnection.close
set DBConnection=nothing
%>


</body>
</html>

