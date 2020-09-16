 <!--#include file="dbpath.asp"-->             
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
           
			<!-- Matching Program Scan Hbar individual Labels into large Labels-->
			<!-- Designed to ensure every piece of HBar gets scanned before going to Shipping-->
			<!--ScanHbar.asp - ScanHbarMatching.asp -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>HBAR Matching </title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
  <meta http-equiv="refresh" content="1000" >
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
 
  </script>
 
     <!--#include file="TodayAndYesterday.asp"-->
     <%
HBarLabel = trim(request.Querystring("HBarLabel")) ' Big Label
'Remove Truck tag 11
if Left(HBarLabel,2) = "11" then 
	HBarLabel = Right(HBarLabel, Len(HBarLabel)-2)
end if  

Label = HBarLabel

ScanDate = Date

'Break Down Label into Job, Floor, Opening, Label Number 

LabelNum = Right(Label, 1)
Label = Left(Label, Len(Label)-2)
LabelJob = Left(Label,3)
Label = Right(Label, Len(Label)-3)
LabelFloor = Left(Label,  inStr(1, label, "-", 0) - 1)
LabelOpening = Right(Label, Len(label)- inStr(1, Label, "-", 0))

LabelMin = (LabelNum * 50) - 49
LabelMax = LabelNum * 50




HbarScan = request.Querystring("HBarScan") ' Little Label
HBarScan = Replace(HBarScan,"%2F", "/")
'Testing HbarScan = ""
Scan = HBarScan


IsError = False
IDFound = False
ErrorDetail =""

if Len(HbarScan) < 10 then
	IsError = True
	ErrorDetail = "Please Scan an H-Bar Label - Nothing Entered"
else
	ScanJob = Left(Scan, 3)
	Scan = Right(Scan, Len(Scan)-3)
	ScanFloor = Left(Scan,  inStr(1, Scan, "-", 0) - 1)
	ScanOpeningFull = Right(Scan, Len(Scan)- inStr(1, Scan, "-", 0))
	
	FirstSlash = inStr(1, ScanOpeningFull, "/", 0)
	SecondSlash = inStr(FirstSlash+1, ScanOpeningFull, "/", 0)
	ScanNum = MID(ScanOpeningFull, FirstSlash+1, SecondSlash-FirstSlash -1)
	ScanOpening = Left(ScanOpeningFull, FirstSlash - 1)

	'Check to Match Big Label to Little Label
	'If Little Label has Same Job / Floor / Opening / Num within limits
	
	if (LabelJob = ScanJob) AND (LabelFloor = ScanFloor) And (LabelOpening = ScanOpening) then
		if ScanNum+0 >= LabelMin+0 AND ScanNum+0 <= labelMax+0 then
		IsError= False
	
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM [HBAR_" & ScanJob & ScanFloor & "]"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		rs.filter = "TAG = '" & trim(ScanOpeningFull) & "'"

		if rs("BundleDate")  > 1 then
			IsError = True
			ErrorDetail = "Already Scanned - This H-Bar has already been scanned."
		else
			rs("BundleDate") = Now
			rs.update
		End if
		rs.close
		set rs = nothing
		

	
	
		else
			' Correct Job Floor Tag BUT Wrong Label Num
			IsError = True
			ErrorDetail = "Wrong Label Number - This H-Bar goes to another Label - Nothing Entered"
		end if
	else
		' Label <> Scan for Job Floor Tag
		IsError = True
		ErrorDetail = "HBAR does not Match Label - Nothing Entered"
	end if 
end if


 %>
</head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="ScanHbar.asp" target="_self">Bundle Label</a>
        </div>

    <form id="igline" title="H-BAR MATCHING" class="panel" name="igline" action="ScanHbarMatching.ASP" method="GET" selected="true">
        <H1 align ="center" > Scan H-Bar for: <% response.write HBarLabel%> </h1>
		<h2 align ="center" > Job: <%response.write LabelJob %> Floor: <%response.write LabelFloor %> Opening: <%response.write LabelOpening %> </h2>
		<h2 align = "center" > Number: <%response.write LabelNum %> Only Scan HBAR between <%response.write LabelMin %> : <%response.write LabelMax %></h2>
			<div class="row">
                <label>
					<% 
					if IsError = False then
						response.write HBarLabel & " - SCAN SUCCESSFUL - " & ScanOPeningFull
				' & DetailsID
					else
						response.write ErrorDetail
					end if	
					%>
					</label>
              
            </div>

        <fieldset>
       
            
         <div class="row">

                        <div class="row">
                <label>HBAR</label>
                <input type="text" name='HBARScan' id='inputbcw' >
				<input type="Hidden" name='HBarLabel' Value =' <%response.write HBarLabel%>' >
            </div>
            
                              	



            
            
                  <script type="text/javascript">
				  
				 		  
            function callback1(barcode) {
                var barcodeText = "BARCODE:" + barcode;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
        
            }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("inputbcw");
    
    
        textbox.value = barcode;
		igline.submit();
    
}
        </script>
        
        
           
         
 
        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:igline.submit()">Submit</a>  
            
<br>
<hr>
<br>

<table border = "1" align = 'Center'>
	<tr><th>Saw ID</TH><th>HBAR TAG</TH><th>Cut Time</TH><th>Matched Time</TH></tr>
<%

		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM [HBAR_" & LabelJob & LabelFloor & "]"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		
Completed = 0
Total = 0
rs.filter = ""
rs.filter = "BundleLabel = '" & trim(Right(HBarLabel, (Len(HBarLabel)-(Len(LabelJob)+Len(LabelFloor)+1)))) & "'"
'BundleLabel is the Opening and Truck, so need to remove Job, floor, comma

Do while not rs.eof

	Response.write "<tr>"
	Response.write "<td>" & rs("ID") & "</td>"
	Response.write "<td>" & rs("Tag") & "</td>"
	Response.write "<td>" & rs("ddate") & " " & rs("dtime") & "</td>"
	Response.write "<td>" & rs("BundleDate") & "</td>"
	Response.write "</tr>"
	Total = Total + 1
	
	if rs("BundleDate")  > 1 then
		Completed = Completed + 1
	end if
rs.movenext
loop

if Completed = Total then
	Response.write "<tr>"
	Response.write "<td colspan = '4' align = 'center'><font size='20' >DONE</Font></td>"
	Response.write "</tr>"
end if
%>
</table>
            
            </form>
			

	<%
	DBConnection.close
	set DBConnection=nothing
%>	

</body>
</html>