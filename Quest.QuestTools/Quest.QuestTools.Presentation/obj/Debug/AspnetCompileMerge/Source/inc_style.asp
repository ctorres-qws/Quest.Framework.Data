<!--#include file="te_config.asp"-->
<link rel="stylesheet" type="text/css" href="te.css">
<title>:: Quest Tools :: </title>

<% 
if request.querystring("hm") = "yes" then
	response.write ""
else
%>
	
		<table border=0 cellspacing=1 cellpadding=2 width=100%>
		<tr>
			<!--<td class="smallertext"></td>-->
			<td class="smallerheader" width=130 align=right>
			<%
			if bProtected then 
				response.write session("teFullName")
				response.write "<a href='te_logout.asp' target='_self'>(logout)</a>" 
			end if
			%>
			</td>
		</tr>
		</table>

<% end if %>
