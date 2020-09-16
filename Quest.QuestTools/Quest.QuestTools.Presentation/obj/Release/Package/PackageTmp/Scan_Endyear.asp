<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
             <!--#include file="dbpath.asp"-->
			<!-- Designed December 2019 as a Table location for end of year Windows completed but not shipped. -->
			<!-- Scan_Endyear.asp is scanner-->
			<!-- X_SHIP_ENDYEAR is Database table-->
			<!-- EndYear_Report.asp is the Report to view it-->
			
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>End Year Scan</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
 
 <!--#include file="TodayAndYesterday.asp"-->    
<%      
	ERROR = "Already Scanned - Not Sent"
	
	bc = request.querystring("window")
	ScanCheck = 0
	
	if not bc ="" then
		Label = bc
		jobname = Left(bc, 3)
		Label = Right(Label, Len(Label)-3)
		floor = Left(Label,  inStr(1, label, "-", 0) - 1)
		tag = Right(Label, Len(label)- inStr(1, Label, "-", 0))
	end if
if floor >= "0" then

	sizecheckid = 0

	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select

end if

Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_SHIP_EndYear WHERE Barcode='" & bc & "'"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

rs.filter = "BARCODE = '" & bc & "'"
  
if RS.EOF then
		rs.addnew 
		rs.fields("BARCODE") = bc
		rs.fields("ScanDate") = Now				
		RS.UPDATE
		ScanCheck = 1
Else
sizeCheckID= 0
	error = "Barcode Already Scanned"
	gstr_ErrMsg = "Err"
	ScanCheck = 2
End if

DbCloseAll

End Function

 %>

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle">End Year</h1>
		<a class="button leftButton" type="cancel" href="index.html#_Ship" target="_self">Shipping</a>
        </div>
   
   
   
    <form id="assy1" title="End Year" class="panel" name="assy1" action="Scan_Endyear.asp" method="GET" selected="true">

 				
			<div class="row">
                <label>
				<% 
				if ScanCheck = 0 then
				else
					if ScanCheck = 1 then
					response.write bc & " - Sent <BR>" 
					else
					response.write error & "-" & bc
					end if	
				end if
				%>
				</label>
              
            </div>

        <h2>Window Scan</h2>
        <fieldset>

 
            
        <div class="row">
            <label>Window</label>
            <input type="text" name='window' id='inputbcw' >
        </div>

 
        </fieldset>
		
		<script type="text/javascript">
				  
				 		  
            function callback1(barcode) {
                var barcodeText = "BARCODE:" + barcode;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
        
            }
            
		function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
			var textbox = document.getElementById("inputbcw");
			textbox.value = barcode;
			assy1.submit();
    
}
			
        </script>
        <BR>
        <a class="whiteButton" href="javascript:assy1.submit()">Submit</a>
            
            
            
            </form>
</body>
</html>
