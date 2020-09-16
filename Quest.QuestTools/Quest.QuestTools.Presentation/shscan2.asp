                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            <!--#include file="connect_barcodeqc.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>iUI Theaters</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>
 
     
     <% 
	 
	 tablename = request.querystring("jobname")
	 multi = request.querystring("multi")
	 bc = request.querystring("window")
	  inputbcw = request.querystring("inputbcw")
	 
	

Set rs7 = Server.CreateObject("adodb.recordset")
'strSQL7 = "Select * From [SHIP_" & jobname & floor & "]"
strSQL7 = "Select * From " & tablename & " ORDER BY BC ASC"
rs7.Cursortype = 2
rs7.Locktype = 3
rs7.Open strSQL7, DBConnection





'
'if floor > "0" then
'
'sizecheckid = 0
'
'  Do while not rs.eof
'  if rs("BARCODE") = bc AND rs("DEPT") = "ASSEMBLY" then
'  response.write "&nbsp;<img src='images/square.gif'>"
'  sizecheckid = rs("ID")
'  end if
'  rs.movenext
'  loop
'  
'if sizecheckid = 0 then
'
''rs2.filter = "NUMBER = " & EMPLOYEE
''IF RS2("NUMBER") = "" THEN
''ELSE
''FIRST = RS2("FIRST")
''LAST = RS2("LAST")
''END IF
'
'if Len(employee) = 4 AND Len(bc) > 5 then
'
'rs.addnew 
'rs.fields("BARCODE") = bc
'rs.fields("JOB") = jobname
'rs.fields("FLOOR") = floor
'rs.fields("TAG") = tag
'rs.fields("DEPT") = DEPTVAR
'rs.fields("EMPLOYEE") = EMPLOYEE
'rs.fields("DATETIME") = STAMPVAR
'rs.fields("FIRST") = FIRST
'rs.fields("LAST") = LAST
'RS.UPDATE
'
'else 
'error = "Wrong Barcode, Try Again"
'end if
'
'end if
'
'end if
  
 %>

     

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="shnew.asp" target="_webapp">Job</a>
        </div>
   
   
   
    <form id="ship" title="Scan Multiple Windows" class="panel" name="ship" action="shscan2.asp" method="GET" target="_self" selected="true">

        <h2>Scan for <% response.write tablename %></h2>
        <fieldset>
       
         
      <div class="row">
             <label>Window</label>
                <input type="text" name='window' id='inputbcw' >
                <input name="ws" type="hidden" value="1" />
                <input name="jobname" type="hidden" value="<% response.write tablename %>" />
                <input name="window" type="hidden" value="<% response.write bc %>" />
            </div>
            
                              	
                <% 
				
				


			rs7.movefirst
					do while not rs7.eof
					if rs7.Fields("bc") = inputbcw then
					rs7.Fields("sStatus") = 1
					rs7.update		
					end if
			rs7.movenext
			loop
	



'a=Split(multi)
'for each x in a
'document.write(x & "<br />")
'		do while not rs7.eof
'				response.write x & "<BR>"
'				if rs7.Fields("bc") = x then
'				rs7.Fields("sStatus") = 1
'				rs7.update
'				end if
'		rs7.movenext
'		loop
'
'next
'
' end if



%>
            
            
                     <script type="text/javascript">
				  
				 		  
            function callback1(barcode) {
                var barcodeText = "BARCODE:" + barcode;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
				
        
            }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("inputbcw");
       textbox.value = barcode;
		 ship.submit();
    
}
			
			

    

        </script>
        </script>
        
        
           
         
 
        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:ship.submit()">Submit</a>
            
            
            
            </form>
            
            
                
</body>
</html>
