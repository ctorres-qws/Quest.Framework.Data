<% bctarget = request.querystring("bc")
bc = request.querystring("bc") %>
<html>
  <head>
  
    <title>iUI Theaters</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="no" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>
  
  
    <style>
      * {
          color:#7F7F7F;
          font-family:Arial,sans-serif;
          font-size:12px;
          font-weight:normal;
      }    
      #config{
          overflow: auto;
          margin-bottom: 10px;
      }
      .config{
          float: left;
          width: 200px;
          height: 250px;
          border: 1px solid #000;
          margin-left: 10px;
      }
      .config .title{
          font-weight: bold;
          text-align: center;
      }
      .config .barcode2D,
      #miscCanvas{
        display: none;
      }
      #submit{
          clear: both;
      }
      #barcodeTarget,
      #canvasTarget{
        margin-top: 20px;
      }        
    </style>
    
    <script type="text/javascript" src="/javascript/jquery-1.3.2.min.js"></script>
    <script type="text/javascript" src="/javascript/jquery-barcode.js"></script>
    
	
		<script type="text/javascript">
	
      function generateBarcode(){
        var value = $("#barcodeValue").val();
        var btype = $("input[name=btype]:checked").val();
        var renderer = $("input[name=renderer]:checked").val();
        
		var quietZone = false;
        if ($("#quietzone").is(':checked') || $("#quietzone").attr('checked')){
          quietZone = true;
        }
		
        var settings = {
          output:renderer,
          bgColor: $("#bgColor").val(),
          color: $("#color").val(),
          barWidth: $("#barWidth").val(),
          barHeight: $("#barHeight").val(),
          moduleSize: $("#moduleSize").val(),
          posX: $("#posX").val(),
          posY: $("#posY").val(),
          addQuietZone: $("#quietZoneSize").val()
        };
        if ($("#rectangular").is(':checked') || $("#rectangular").attr('checked')){
          value = {code:value, rect: true};
        }
        if (renderer == 'canvas'){
          clearCanvas();
          $("#<% response.write bctarget %>").hide();
          $("#canvasTarget").show().barcode(value, btype, settings);
        } else {
          $("#canvasTarget").hide();
          $("#<% response.write bctarget %>").html("").show().barcode(value, btype, settings);
        }
      }
          
      function showConfig1D(){
        $('.config .barcode1D').show();
        $('.config .barcode2D').hide();
      }
      
      function showConfig2D(){
        $('.config .barcode1D').hide();
        $('.config .barcode2D').show();
      }
      
      function clearCanvas(){
        var canvas = $('#canvasTarget').get(0);
        var ctx = canvas.getContext('2d');
        ctx.lineWidth = 1;
        ctx.lineCap = 'butt';
        ctx.fillStyle = '#FFFFFF';
        ctx.strokeStyle  = '#000000';
        ctx.clearRect (0, 0, canvas.width, canvas.height);
        ctx.strokeRect (0, 0, canvas.width, canvas.height);
      }
      
      $(function(){
        $('input[name=btype]').click(function(){
          if ($(this).attr('id') == 'datamatrix') showConfig2D(); else showConfig1D();
        });
        $('input[name=renderer]').click(function(){
          if ($(this).attr('id') == 'canvas') $('#miscCanvas').show(); else $('#miscCanvas').hide();
        });
        generateBarcode();
		
      });
  
    </script>
  </head>
  <body>
  
      <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Scan" target="_webapp">Scan Tools</a>
        </div>
        
  
      <form id="igline" title="Reprint" class="panel" name="igline" action="bcprint.asp" method="GET" selected="true">
  

<input type="hidden" id="barcodeValue" value="<% response.write bc %>">
   

          <input type="radio" name="btype" id="code128" value="code128" checked="checked" style="display:none;"><label for="code128"></label> 
       

            
        
            <input type="hidden" id="bgColor" value="#FFFFFF" size="7"> 
          <input type="hidden" id="color" value="#000000" size="7"> 

            <input type="hidden" id="barWidth" value="2" size="3"> 
           <input type="hidden" id="barHeight" value="150" size="3"> 

  
            <input type="hidden" id="moduleSize" value="5" size="3"> 
           <input type="hidden" id="quietZoneSize" value="1" size="3"> 
            <input type="hidden" name="rectangular" id="rectangular"><label for="rectangular"></label> 
     
    
            <input type="hidden" id="posX" value="0" size="3"> 
            <input type="hidden" id="posY" value="0" size="3"> 
     
            

       
          <input type="hidden" id="css" name="renderer" value="css" checked="checked"><label for="css"></label> 

  
     <!-- <div id="submit">
        <input type="button" onClick="generateBarcode();" value="barcode">
      </div>-->
        

    
    <div id="<% response.write bctarget %>" class="barcodeTarget"></div>
    

         <% if bc = "" then
		 response.write ""
		 else %>
			<div class="row">
                <label><% if sizecheckid = 0 then
				response.write "" 
				else
				response.write error
				end if				
				%></label>
              
            </div>
            <% 			
			end if %>
        <h2>Scan Barcode</h2>
        <fieldset>
       
            
         <div class="row">

                        <div class="row">
                <label>Window</label>
                <input type="text" name='bc' id='inputbcw' >
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
		igline.submit();
    
}
 </script>
        
			
		
        
        
           
         
 
        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:igline.submit()">Submit</a>
        
        <a class="whiteButton" href="" onClick="window.print()">Print</a>
        <br><Br>
     

            <a href="safari://bcprint.asp" target="_new">jody</a>
            
            </form>
    	
        
  
