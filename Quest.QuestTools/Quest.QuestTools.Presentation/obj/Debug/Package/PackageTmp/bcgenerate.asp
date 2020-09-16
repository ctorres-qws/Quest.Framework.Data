
<html>
  <head>
    <style>
      * {
          color:#000000;
          font-family:Arial,sans-serif;
           font-weight:heavy;
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
  

<input type="hidden" id="barcodeValue" value="<% response.write id %>">
   

          <input type="radio" name="btype" id="code128" value="code128" checked="checked" style="display:none;"><label for="code128"></label> 
       

            
        
            <input type="hidden" id="bgColor" value="#FFFFFF" size="7"> 
          <input type="hidden" id="color" value="#000000" size="7"> 

            <input type="hidden" id="barWidth" value="3" size="3"> 
           <input type="hidden" id="barHeight" value="60" size="3"> 

  
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
    	
  
