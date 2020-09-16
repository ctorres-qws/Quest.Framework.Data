<!--#include file="@common.asp"-->
<%
	'-------------------------------------------------------------
	'TableEditoR 0.6 Beta
	'http://www.2enetworx.com/dev/projects/tableeditor.asp
	
	'File: index.asp
	'Description: Default page for TableEditoR
	'Written By Hakan Eskici on Nov 01, 2000

	'You may use the code for any purpose
	'But re-publishing is discouraged.
	'See License.txt for additional information	

	'Change Log:
	'-------------------------------------------------------------
	'# Nov 15, 2000 by Hakan Eskici
	'Added permission assignment for Table and Field functions
	'-------------------------------------------------------------
%>

<!--#include file="te_config1.asp"-->
<%
'We may come here as a result of a session timeout
'or direct access without opening a session
'redirection will occur after login
if request("comebackto") <> "" then
	sReferer = request("comebackto")
	sGoBackTo = "?" & request.querystring
end if

sub AskForLogin(sText)
%>
<div class="csBody">
<p class="smallheader">
	<%
		sMsg = sText
		if request("comebackto") <> "" then
			sMsg = sMsg & "<br>Please re-login."
			if sText <> "" then
				%><!--#include file="te_header.asp"--><%
			end if
		else
			if sText = "" then 
				sMsg = sMsg & "<br>Please login."
			else
				%><!--#include file="te_header.asp"--><%
			end if
		end if
		response.write sMsg
	%>
</p>

<form action="index.asp<%=sGoBackTo%>" method="post">
<table border=0>
	<tr>
		<td class="smallerheader">User Name</td>
		<td><input type="text" name="txtUserName" class="tbflat"></td>
	</tr>
	<tr>
		<td class="smallerheader">Password</td>
		<td><input type="password" name="txtPassword" class="tbflat"></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" name="cmdLogin" class="cmdflat" value=" Login "></td>
	</tr>
</table>
</form>


<%
	if sText = "" then
	else
		%><!--#include file="te_footer.asp"--><%
	end if

end sub

	if request("cmdLogin") <> "" then
		'User provided the user name and password
		
		'Open the connections and create the recordset object
		OpenRS GetConnectionStr(true)
		
		sUserName = request("txtUserName")
		sPassword = request("txtPassword")

		sSQL = "SELECT * FROM Users WHERE UserName = '" & sUserName & "'"
		rs.Open sSQL, , , adCmdTable
		if not (rs.bof or rs.eof) then
			if rs("Password") = sPassword then
				'Login succeeded
				'Store info into session
				session("teUserName") = sUserName
				session("teFullName") = rs("FullName")
				session("rAdmin") = rs("rAdmin")
				session("rTracking") = rs("rTracking")
				session("rSupplier") = rs("rSupplier")
				session("rMidLevel") = rs("rMidLevel")
				session("rQueryExec") = rs("rQueryExec")
				session("rSQLExec") = rs("rSQLExec")
				session("rTableAdd") = rs("rTableAdd")
				session("rTableEdit") = rs("rTableEdit")
				session("rTableDel") = rs("rTableDel")
				session("rFldAdd") = rs("rFieldAdd")
				session("rFldEdit") = rs("rFieldEdit")
				session("rFldDel") = rs("rFieldDel")
				session("Location") = rs("Location")

                if sReferer = "" then
					if session("Location") = "USA" then
						response.redirect "indexTexas.html"
					else
						response.redirect "index.html"
					end if
				else
					response.redirect sReferer
				end if
			else
				'Login failed - Wrong password
				AskForLogin "Incorrect password."
			end if
		else
			'User not found
			AskForLogin "Incorrect credentials."
		end if
		
		CloseRS()
		
	else
%>
<!--#include file="te_header.asp"-->
<%
		if bProtected then
		'If protection is ON, ask for login
			AskForLogin ""
		else
		'if protection is OFF, display a warning
%>
	<p class="smallheader">
		You are not using protection!
	</p>
	Any visitor who knows the exact location of the TableEditor files may view or change the information in your databases.<br>
	To enable protection, open <strong>te_config.asp</strong> file and set <strong>bProtected = True</strong>.<br><br>
	You may go to <a href="indexa.asp" style="color: #0000FF;">Admin Page</a> now.

<%
		end if
%><%'<!--#include file="te_footer.asp"-->
%><%
	end if
%>
</div>
<!--#include file="te_footer.asp"-->