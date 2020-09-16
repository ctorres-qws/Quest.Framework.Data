<!--#include file="dbpath.asp"-->                    
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            
			<!-- Glass Tool to be used on the Optima machine - For Maintaining the Optimization Log Files-->
			<!-- Gives Glass information collected with Optima Cutting by Operators-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Optimization Tool</title>
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
Transfer = False
Shift = REQUEST.QueryString("Shift")
if Shift = "" then
	Shift= "DayShift"
end if


EMPLOYEE = request.querystring("EMPLOYEEID")
Opfile = UCASE(request.querystring("OpFile"))

GlassCutDate = DATE
GlassCutTime = TIME
'The date to add on the list for cutting Glass
IsError = False
' Reset the Variable for locating an Error

if Len(Employee) = 4 AND OpFile <> "" then

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		'Process(false)
		Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

else
	IsError = True
	error = "Invalid Employee ID of Optimization File"
end if 
		

Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

'Set Glass Inventory Update Statement
				StrSQL = FixSQLCheck("UPDATE OptimizeLog  SET [GlassCutDate]=#" & GlassCutDate & "#, [GlassCutTime]=#" & GlassCutTime & "#, [Shift]='" & Shift & "', Employee= '" & Employee & "'  WHERE OpFile = '" & Opfile & "'", isSQLServer)
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
				
		'		On Error Resume Next

		'		Dim objFSO
		'			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		'			objFSO.CopyFile "\\Hightower\D:\Macotec Cutfiles\M00001.z01", "\\Hightower\D:\Macotec CNC files\"
		'			Transfer = TRUE
		'		If Err Then
		'			Transfer = FALSE
		'			WScript.Quit 1
		'		End If

DbCloseAll

End Function

  
 %>

     

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle">Optimize Glass</h1>
		<!--Back Button Removed-->
    </div>
   
   
   
    <form id="igline" title="Optimize Glass" class="panel" name="igline" action="OptimizationLogGlass.asp" method="GET" selected="true">
         <h2>Optimize Glass - <%response.write Shift%></h2>
		 
		 <%

		 
		 
		 DayShiftColour ="whiteButton"
		 NightShiftColour ="whiteButton"

		 
		 Select Case Shift
			Case "DayShift"
				DayShiftColour ="greenButton"
			Case "NightShift"
				NightShiftColour ="greenButton"
		 End Select
		 %>
		 
		 
		<div class="row">
			<table>
			<tr>
				<td><a class="<%response.write DayShiftColour%>" href="OptimizationLogGlass.asp?Shift=DayShift" target = "_Self" >Day Shift</a> </td>
				<td><a class="<%response.write NightShiftColour%>" href="OptimizationLogGlass.asp?Shift=NightShift" target = "_Self" >Night Shift </a> </td>
			</tr>
			
			</table>
		</div>
		
		 <% if Opfile = "" then
		 response.write ""
		 else %>
			<div class="row">
                <label><% 
				if IsError = False then
					response.write OpFile & " - Cut" 
					if Transfer = True then
						response.write " And Sent" 
					else
						'response.write " AND NOT SENT"
					end if
				else
				response.write error
				end if				
				%></label>
              
            </div>
            <% 			
			end if %>
			<BR>
        <fieldset>
       
	   
	        <div class="row">
                <label>Employee #</label>
                <input type="text" name='employeeID' id='inputbce' >
            </div>
	   
            
			<div class="row">
                <label>Optimization File</label>
				 <select name="Opfile">
				<%
				Set rs = Server.CreateObject("adodb.recordset")
				strSQL = FixSQLCheck("Select * FROM OptimizeLog WHERE ( ISNULL(SHIFT) OR SHIFT='') order By Opfile ASC", b_SQL_Server)
				rs.Cursortype = 2
				rs.Locktype = 3
				rs.Open strSQL, DBConnection 
				if rs.eof then
				else
				
				Do while not rs.eof
					response.write "<option value='"
					response.write rs("OpFile")
					response.write "'>"
					response.write rs("OpFile") & " - " & rs("Job") & rs("Floor")
					response.write "</option>"
				rs.movenext
				loop
				end if
				%>


            </div>
					
				<input type="hidden" name='Shift' id='Shift' value='<%response.write Shift%>' />
           
		   </fieldset>
			 <BR>
				<a class="whiteButton" href="javascript:igline.submit()">Submit</a>
				
            </form>
			
			
		
<%
rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
Set objFSO = Nothing
%>	

</body>
</html>