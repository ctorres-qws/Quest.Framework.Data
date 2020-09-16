<!--#include file="dbpath.asp"-->

<% 

'Create a Query
    SQL = "Select * FROM X_BARCODE"
'Get a Record Set
    Set RS = DBConnection.Execute(SQL)
	
JOB = REQUEST.QueryString("JOB")
FL = REQUEST.QueryString("FLOOR")

'IF FLOOR = "" THEN
'RS.FILTER = "JOB = '" & JOB & "' AND DEPT = '" & DEPT & "'"
'ELSE
'RS.FILTER = "JOB = '" & JOB & "' AND FLOOR = " & FL & " AND DEPT = '" & DEPT & "'" 
'END IF
%>
<BR>
	<table width="600PX" border="0">
  <tr>
    <td>DEPT</td>
    <td>JOB</td>
    <td>FLOOR</td>
    <td>TAG</td>
    <td>EMPLOYEE</td>
    <td>BARCODE</td>
  </tr>
<%
Do while not rs.eof
%>
  <tr>
    <td><% response.write rs.Fields("DEPT") %></td>
    <td><% response.write rs.Fields("JOB") %></td>
    <td><% response.write rs.Fields("FLOOR") %></td>
    <td><% response.write rs.Fields("TAG") %></td>
    <td><% response.write rs.Fields("EMPLOYEE") %></td>
    <td><% response.write rs.Fields("BARCODE") %></td>
  </tr>
  <% rs.movenext
loop
rs.close
set rs = nothing
DBConnection.close
Set DBConnection = nothing
%>
</table>
</ul>


<html>
<br>

