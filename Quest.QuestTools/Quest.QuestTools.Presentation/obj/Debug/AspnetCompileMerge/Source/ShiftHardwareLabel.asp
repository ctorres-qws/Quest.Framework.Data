<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- Created January 2019 Michael Bernholtz - Label to be Printed for Shift Buggy System-->
<!-- Next Button Forces Print and then Redirects back to ShiftHardwareSashKit or ShiftHardwareFrameKit-->

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ko" lang="ko">
<head>
<title>Label Printer</title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta name="viewport" content="width=device-width,initial-scale=1,user-scalable=no" />

</head>
<body>
<%
Container ="Buggy"
Ticket = UCASE(Request.Querystring("Ticket"))
SELECT CASE Ticket
	CASE "SASH"
		ReturnSite = "ShiftHardwareSashKit.asp"
		Container ="Buggy"
	CASE "FRAME"
		ReturnSite = "ShiftHardwareFrameKit.asp"
		Container ="Buggy"
	CASE "CORNER"
		ReturnSite = "ShiftHardwareCornerPostKit.asp"
		Container ="Trolley"
	CASE ELSE
		ReturnSite = "ShiftHardwareView1.asp"
End SELECT

 PositionX = request.querystring("PositionX")
 PositionY = request.querystring("PositionY")
 PositionI = request.querystring("PositionI")
 Side = request.querystring("Side")
 NextType = request.querystring("NextType")
 
 

 Side = request.querystring("Side")
 Job = request.querystring("Job")
 Floor = request.querystring("Floor")
 Barcode = request.querystring("Barcode")
 Bin = request.querystring("Bin")
 Cart = request.querystring("Cart")

if NextType = "Jump" then
	First = "YES"
	NextX = request.querystring("NextX")
	NextY = request.querystring("NextY")
	NextI = request.querystring("NextI")
	NextSide = request.querystring("NextSide")
else 
	NextX = PositionX + 1
	NextY = PositionY
	NextI = PositionI
	NextSide = SIDE
	if Ticket = "CORNER" then
	if NextX = 8 then
			NextX = 0
			NextY = PositionY + 1
			if NextY = 5 then
				NextY = 0
				if Side = "Front" then
					NextSide = "Back"
				end if
				if Side = "Back" then
					NextSide = "Front"
					NextI = PositionI + 1
				end if 
			end if
		end if
	else
		if NextX = 12 then
			NextX = 0
			NextY = PositionY + 1
			if NextY = 10 then
				NextY = 0
				if Side = "Front" then
					NextSide = "Back"
				end if
				if Side = "Back" then
					NextSide = "Front"
					NextI = PositionI + 1
				end if 
			end if
		end if
	end if
end if
 
if barcode = "Empty" then
else
%> 
 
<table align= "center" frame="box" width="300px" cellspacing="1" cellpadding="1">
	
	<tr>
		<td align = 'center' style="font-size: 75%;"> <b><%response.write BARCODE %></b></td>
		<td align = 'Left' style="font-size:75%;"> <b><%response.write Container & " " & PositionI %><b></td>
	</tr>
	
	<tr>
		<td Rowspan = "4" align = 'center'><img src="http://chart.apis.google.com/chart?cht=qr&chs=65x65&chl=<% response.write PositionX +1 &"-"& PositionY + 1&"-"& PositionI &"-"& Side &"-"& Job&"-"& Floor %>&chld=H|0" alt="Barcode" /></td>
		<td align = 'Left' style="font-size: 75%;"> <b>Row <%response.write PositionY + 1%><b></td>
	</tr>
	<tr>
		<td align = 'Left' style="font-size: 75%;"> <b>Column <%response.write PositionX + 1%><b></td>
	</tr>
	<tr>
		<td align = 'Left' style="font-size: 75%;"> <b>Side <%response.write Side %><b></td>
	</tr>
	<tr>
		<td align = 'Left' style="font-size: 75%;"> <b> <%response.write  Ticket  %> KIT<b></td>
	</tr>

	
</table>
<br>

<script>
print()
</script>
<p>
 
</p>
<p>&nbsp;</p>
<%
end if
%>
<meta http-equiv="refresh" content="0;url=<%response.write ReturnSite & "?Job=" & JOB & "&Floor=" & FLOOR & "&PositionX=" & NEXTX & "&PositionY=" & NEXTY & "&PositionI=" & NEXTI & "&Side=" & NEXTSIDE & "&First=" & First & "&Ticket=" & Ticket %>" />
</body>
</html>

