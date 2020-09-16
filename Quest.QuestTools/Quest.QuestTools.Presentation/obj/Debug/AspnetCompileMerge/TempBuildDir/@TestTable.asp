<html>
	<head>
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script type="text/javascript" src="javascript/jquery.floatThead.min.js"></script>
<style>
	table thead tr th {font-family: arial; font-size: 12px;}
</style>
</head>
<script>
	$( document ).ready(function() {
		//alert('Hi');
		
		var $table = $('table.demo');
		$table.floatThead();
		
	});
</script>
	<body>
		<div style='height: 400px;'>a</div>
		<table class='table demo' style='border-collapse: separate;'>
			<thead><tr style='background-color: #eeeeee;'><th>Field 1</th><th>Field 2</th><th>Field 3</th><th>Field 4</th><th>Field 5</th><th>Field 6</th><th>Field 7</th><th>Field 8</th></tr></thead>
<%
		For i = 0 to 100
%>
			<tr><td><%= i %></td><td>2018-07-11</td><td>8:58:12</td><td>200</td><td>4359</td><td>192.168.10.122</td><td>/jobsmanagerdepts.aspx</td><td><%= Now %></td></tr>
<%
		Next
%>
	</table>
	</body>
</html>