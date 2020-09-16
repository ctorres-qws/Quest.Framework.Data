<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
  <html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
   
</head>

<% 

'Create a Query
    SQL = "Select * FROM Y_INV"
'Get a Record Set
    Set RS = DBConnection.Execute(SQL)
	
	
'Create a Query
    'SQL3 = "DELETE * FROM X_BARCODETEMP1"
'Get a Record Set
    'Set RS3 = DBConnection.Execute(SQL3)	
	
	
' Set rs2 = Server.CreateObject("adodb.recordset")
' strSQL = "SELECT * From X_BARCODETEMP1"
' rs2.Cursortype = 2
' rs2.Locktype = 3
' rs2.Open strSQL, DBConnection

bcid = request.querystring("bcid")
rs.filter = "ID = '" & bcid & "'"
'Do while not rs.eof

bctarget = "anything"
id = rs("id")



%> 
<table width="300" border="0" selected="true" class="heading1">
  <tr>
    <td>Part: <h1><% response.write rs("part") %></h1></td>
  </tr>
  </table>

<table width="300" border="0" selected="true">
  <tr>

    <td>Finish</td>
    <td>L(in)</td>
    <td>Qty</td>
  </tr>
  <tr>

    <td><% response.write rs("colour") %></td>
    <td><% response.write rs("linch") %></td>
    <td><% response.write rs("qty") %></td>
  </tr>
</table><br>
<!--#include file="bcgenerate.asp"--><br />
<img src="/partpic/<% response.write rs("part") %>.png" />


   
  </body>
</html>

<% 

rs.close
' set rs=nothing
' rs2.close
' set rs2=nothing
' rs3.close
' set rs3=nothing
DBConnection.close
set DBConnection=nothing
%>

