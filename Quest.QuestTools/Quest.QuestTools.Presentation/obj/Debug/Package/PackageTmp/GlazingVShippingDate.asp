
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Report Designed February 2018 for Jody Cash, Antonio Colalillo, Shaun Levy -->
<!-- Report shows Shipping and Glazing scans for a day -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Shipping Glazing Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
  <script src="sorttable.js"></script>
  
  



</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Report</a>
    </div>


    <form id="GlazeShip" title="Glaze vs Ship" class="panel" name="GlazeShip" action="GlazingVShipping.asp" method="GET" selected="true">

        <h2>Choose Day, Month, Year</h2>
        <fieldset>
       

         <div class="row">
                       
            <label>Day </label>
            <input type="number" name='sDay' id='sDay' value='<%response.write CINT(Day(Date))%>'>
		</div>
		<div class="row">
                       
            <label>Month </label>
			<select name='sMonth' id='sMonth'>
				
				
				
				<option value="1" <% if Month(Date) =  1 then response.write "Selected" end if%>>January</option>
				<option value="2" <% if Month(Date) =  2 then response.write "selected" end if%>>February</option>
				<option value="3" <% if Month(Date) =  3 then response.write "selected" end if%>>March</option>
				<option value="4" <% if Month(Date) =  4 then response.write "selected" end if%>>April</option>
				<option value="5" <% if Month(Date) =  5 then response.write "selected" end if%>>May</option>
				<option value="6" <% if Month(Date) =  6 then response.write "selected" end if%>>June</option>
				<option value="7" <% if Month(Date) =  7 then response.write "selected" end if%>>July</option>
				<option value="8" <% if Month(Date) =  8 then response.write "selected" end if%>>August</option>
				<option value="9" <% if Month(Date) =  9 then response.write "selected" end if%>>September</option>
				<option value="10" <% if Month(Date) =  10 then response.write "selected" end if%>>October</option>
				<option value="11" <% if Month(Date) =  11 then response.write "selected" end if%>>November</option>
				<option value="12" <% if Month(Date) =  12 then response.write "selected" end if%>>December</option>
			</select>
		</div>		
         <div class="row">
                       
            <label>Year</label>
			<select name="sYear" id="sYear">
				<option value="<%response.write(Year(Now))%>"><%response.write(Year(Now))%></option>
				<option value="<%response.write(Year(Now)-1)%>"><%response.write(Year(Now)-1)%></option>
				<option value="<%response.write(Year(Now)-2)%>"><%response.write(Year(Now)-2)%></option>
				<option value="<%response.write(Year(Now)-3)%>"><%response.write(Year(Now)-3)%></option>

				
			</select>
		</div>
		
		<a class="whiteButton" onClick=" GlazeShip.submit()">View Glazing V Shipping Report</a><BR>
        </fieldset>	
	</form>

<%
DBConnection.close
set DBConnection=nothing
%>


</body>
</html>

