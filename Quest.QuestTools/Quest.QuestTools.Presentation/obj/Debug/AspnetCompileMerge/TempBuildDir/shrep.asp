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
	 
Const adSchemaTables = 20
Set rs3 = DBConnection.OpenSchema(adSchemaTables)


'bc = request.querystring("barcode")
'
'bc = request.querystring("window")
'jobname = Left(bc, 3)
'if inStr(1, bc, "-", 0) = 5 then
'floor = Mid(bc, 4, 1)
'tag = Mid(bc, 5, 5)
'else
'floor = Mid(bc, 4, 2)
'tag = Mid(bc, 6, 5)
'end if
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
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_webapp">Reports</a>
        </div>
   
   
   
    <form id="ship" title="Ship Report" class="panel" name="ship" action="shreportb.asp" method="GET" target="_self" selected="true">
        <h2>Select Job / Floor </h2>
  
                        <div class="row">
                                <fieldset>
<!--<FORM METHOD="GET" ACTION="su_process.asp">-->
<select name="tn">
<option selected name=jobname value="-">-
<!--#include file="ship_tables.inc"-->
</select>
            </div>




       
            
        

            
                              	
                <% 'response.write window & "<Br>" & employeeID %>
            
            
                  <script type="text/javascript">
				  
				 		  
            function callback1(barcode) {
                var barcodeText = "BARCODE:" + barcode;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
        
            }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("inputbcw");
    
  
        textbox.value = barcode;
    
}

     
			
			

    

        </script>
        
        
           
         
 
        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:ship.submit()">Submit</a>
            
            
            
            </form>
            
            
                
</body>
</html>
