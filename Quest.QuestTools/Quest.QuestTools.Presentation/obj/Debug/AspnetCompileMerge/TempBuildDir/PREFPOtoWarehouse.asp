<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
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


<% 
'
''Create a Query
'    SQL = "Select * FROM Y_INV ORDER BY PART ASC"
''Get a Record Set
'    Set RS = DBConnection.Execute(SQL)
	
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV ORDER BY PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

'part = request.QueryString("part")

'id = REQUEST.QueryString("ID")
'aisle = REQUEST.QueryString("aisle")


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html">Home</a>
    </div>
    
        
    
              <form id="wreceive" title="Receive Stock" class="panel" name="wreceive" action="PREFPOtoWarehouseSplit.asp" method="GET" target="_self" selected="true" > 
			  <h2>Receive to Warehouse and PREF by PO</h2>
  
   

<fieldset>
     <div class="row">
                <label>PO</label>
                <input type="text" name='PO' id='PO'>
            </div>
            
</fieldset>


        <BR>
        <a class="whiteButton" href="javascript:wreceive.submit()">Process PO</a><BR>
            


            
            </form> </form>
            
 <form id="conf" title="Edit Stock" class="panel" name="conf" action="stock.asp#_remove" method="GET" target="_self">
        <h2>Stock Edited</h2>
  

            
            </form>
			
  <script type="text/javascript">
				  
				 		  
            function callback1(barcode) {
                var barcodeText = "BARCODE:" + barcode;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
        
            }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("PO");
    
   // if ( barcode.length == 4 ) {
   //     textbox = document.getElementById("inputbce");
   //   }
    
        textbox.value = barcode;
    
}
			
			

    

        </script>
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set conntemp=nothing
%>

