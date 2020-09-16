<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!--Quality Control Button Needs Password to access - Michael Bernholtz, Reuqested by Victor -->
<!-- Password is coded with basic ASP protection, it is simple to crack for a programmer -->
<!-- I do not think more complicated security is required for this particular activity-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Delete QC Item - Password Protected</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
   <script src="sorttable.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

	<script type="text/javascript">
 
    function EnterPassword(pwd)
    {
		if (UCASE(pwd) == "JODY")
		{
			document.getElementById(adminpass).style.display = block;
			document.getElementById(home).style.display = block;
		}
		else
		{
			document.getElementById(adminpass).style.display = block;
			document.getElementById(home).style.display = none;
		}
    }
 
 
</script>
<%
Password = UCASE(TRIM(Request.Form("pwd")))
back = request.querystring("back")
QCID = Request.Querystring("QCID")

%>	
	
</head>
<body>
    <p>&nbsp;</p>
    <div class="toolbar">
    <h1 id="pageTitle"></h1>
        <a id="backButton" class="button"  href="back.html"></a>
                <a class="button leftButton" type="cancel" href="QCInventory" target="_self">QC</a>
<!--        USE THIS TO PUT ON PAGES LIKE XA7.ASP ETC, NOT ON THIS PAGE-->
        <!--<a class="button leftButton" type="cancel" href="demos.html">Demos</a>-->
  
    </div>
<%
if UCASE(Password = "JODY") or UCASE(SPassword) = "JODY" then

%>
	<ul id="home" title="Shipping Admin Tools" selected="true">
		<li class="group">Accessory Master Library </li>
			<li><a href="ShippingLibraryAdd.asp" target="_self">Add New Type of Accessory</a></li>
			<li><a href="ShippingLibraryEdit.asp" target="_self">Manage Existing Accessory List</a></li>
		
		<li class="group">Truck Management</li>
			<li><a href="ShippingTruckEdit.asp" target="_self">Manage Trucks</a></li>
			<li><a href="ShippingTruckClose.asp" target="_self">Close Truck</a></li>	
			<li><a href="ShippingTruckReopen.asp" target="_self">Re-open Closed Truck</a></li>	
			
		<li class="group">Shipping Inventory Reporting</li>
			<li><a href="ShippingTruckView.asp" target="_self">View All Trucks</a></li>
			<li><a href="ShippingItemView.asp" target="_self">View Shipping List by Truck</a></li>
			<li><a href="ShippingItemViewNoTruck.asp" target="_self">View Shipping List Unassigned</a></li>
		<li class="group">Shipping Item Management</li>
			<li><a href="ShippingItemEdit.asp" target="_self">Manage Existing Shipping List</a></li>
		<li class="group">Logout</li>
		<li><a href="ShipAdminLogout.asp" target="_self">Logout</a></li>
</ul>

<% 

else
%>
<form id="adminpass" title="Administrative Tools" class="panel" name="enter" action="ShipAdminHome.asp" method="post" target="_self" selected="True">
<fieldset>
			<div class="row" >
				<label>Password:</label>
				<input type="password" name='pwd' id='pwd' ></input>
			</div>
</fieldset>
<a class="whiteButton" href="javascript:adminpass.submit()">Enter password</a>
	</form>
	
<%
end if
%>


</div>

</body>
</html>
