<!--#include file="dbpath.asp"-->
<%
	Dim str_MsgErr
	Dim cn_SQL, rs_Data
	Dim str_Job, str_Floors, str_Windows
	Dim str_Colour

	str_Job = Request("Jobs")
	str_Colour = Request("Colours")

	Set cn_SQL = Server.CreateObject("ADODB.Connection")
	cn_SQL.Open GetConnectionStr(true)

	'Response.Write(str_Job & "<br/>")
	'Response.Write(Request("AddMat_Part") & "<br/>")
	'Response.Write(Request("AutoPost") & "<br/>")
	'Response.Write(Request("Action") & "<br/>")
	'Response.End()

	If Request("AutoPost") <> "TRUE" Then
		Select Case(UCase(Request("PageAction")))
			Case "EDIT_JOB"
				Call EditJob(str_Job, Request("Edit_Floors"), Request("Edit_Windows"))
			Case "ADD_JOB"
				str_Job = Request("AddJob_Name")
				Call AddJob(Request("AddJob_Name"), Request("AddJob_Floors"), Request("AddJob_Windows"))
			Case "SELECT_JOB"
				str_Job = Request("Job_Select")
				str_Colour = ""
			Case "SELECT_COLOURS"
			Case "ADD_MATERIAL"
				Call AddMaterial(str_Job, Request("AddMat_Part"), Request("AddMat_Size"), Request("AddMat_Qty"), Request("Colours"))
			Case "MATERIAL_DELETE"
				Call MaterialDelete(Request("MaterialID"))
		End Select
	End If

	Function EditJob(str_Job, str_Floors, str_Windows)
		Set rs_Job = Server.CreateObject("ADODB.Recordset")
		rs_Job.CursorType = GetDBCursorTypeInsert
		rs_Job.LockType = GetDBLockTypeInsert
		rs_Job.Open "SELECT * FROM _qws_Planning_Jobs WHERE JobName='" & str_Job & "'", cn_SQL
		rs_Job.Fields("Floors").Value = str_Floors
		rs_Job.Fields("Windows").Value = str_Windows
		rs_Job.Update
		rs_Job.Close(): Set rs_Job = Nothing
	End Function

	Function AddJob(str_Job, str_Floors, str_Windows)
		Dim rs_Job

		Set rs_Job = Server.CreateObject("ADODB.Recordset")
		rs_Job.CursorType = GetDBCursorTypeInsert
		rs_Job.LockType = GetDBLockTypeInsert
		rs_Job.Open "SELECT * FROM _qws_Planning_Jobs WHERE ID=-1", cn_SQL
		rs_Job.AddNew
		rs_Job.Fields("JobName").Value = str_Job
		rs_Job.Fields("Floors").Value = str_Floors
		rs_Job.Fields("Windows").Value = str_Windows
		rs_Job.Update
		rs_Job.Close(): Set rs_Job = Nothing
		'Response.Redirect("MaterialPlanningJobEntry.asp?Jobs=" & str_Job)
		'Response.End()
	End Function

	Function AddMaterial(str_Job, str_Part, str_Size, str_Qty, str_Color)
		Dim rs_Job
		Dim a_Parts: a_Parts = Split(str_Part & ",", ",")

		str_Colour = str_Color

		Set rs_Job = Server.CreateObject("ADODB.Recordset")
		rs_Job.CursorType = GetDBCursorTypeInsert
		rs_Job.LockType = GetDBLockTypeInsert
		rs_Job.Open "SELECT * FROM _qws_Planning_Materials WHERE ID=-1", cn_SQL
		rs_Job.AddNew
		rs_Job.Fields("JobName").Value = UCase(str_Job)
		rs_Job.Fields("Part").Value = Trim(UCase(a_Parts(0)))
		rs_Job.Fields("Description").Value = UCase(a_Parts(1))
		rs_Job.Fields("Colour").Value = str_Color
		If str_Size <> "" Then rs_Job.Fields("Size").Value = str_Size
		rs_Job.Fields("Qty").Value = Trim(str_Qty)
		rs_Job.Update
		rs_Job.Close(): Set rs_Job = Nothing
		'Response.Redirect("MaterialPlanningJobEntry.asp?Jobs=" & str_Job)
		'Response.End()
	End Function

	Function MaterialDelete(str_ID)
		cn_SQL.Execute("DELETE FROM _qws_Planning_Materials WHERE ID=" + str_ID)
	End Function

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>Quest Dashboard</title>
	<meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
	<link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<link rel="stylesheet" href="/iui/iui.css" type="text/css" />
	<link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css?v=1"  type="text/css"/>
	<script type="application/x-javascript" src="/iui/iui.js"></script>
<link rel="stylesheet" href="javascript/jquery-ui.css">
<script type="text/javascript" src="javascript/jquery-1.11.3.js"></script>
<script type="text/javascript" src="javascript/jquery-ui.js"></script>
	<script type="text/javascript">
		iui.animOn = true;
	</script>
	<style>
		.csForm { background-color: #eaeaea; border: 2px solid #cccccc; border-radius: 5px; width: 600px !important; position: fixed; top: 100px !important; z-index: 9999;}
		.csFormContainer { padding: 50px; }

		.csButton { height: 30px !important; padding: 10px 8px;}

		.csLabel { width: 120px; text-align: right; }
		.csData { width: 120px !important; }

		select { font-size: 22px; }

		input[type='text'], select  {
			margin: 0px 5px 5px 5px;
			padding: 0px 0px 0px 0px !important;
			border-radius: 4px;
			border: 1px solid rgb(200, 200, 200);
			border-image: none;
			text-align: left !important;
			height: 40px;
			xwidth: 120px !important;
		}

		.csLargeInput { width: 300px !important; }

		table, tr { border-bottom: 0px !important; }

	.ui-state-hover, .ui-autocomplete li:hover
	{
		background: rgb(238,238,238);
		margin-left: 1px;
		margin-right: 1px;
	}

.ui-autocomplete {
  font-weight: normal;
  position: absolute;
  top: 100%;
  left: 0;
  z-index: 1000;
  float: left;
  display: none;
  min-width: 280px;
  width: 500px !important;
  padding: 4px 0;
  margin: 2px 0 0 0;
  list-style: none;
  background-color: #ffffff;
  border-color: #ccc;
  border-color: rgba(0, 0, 0, 0.2);
  border-style: solid;
  border-width: 1px;
  -webkit-border-radius: 5px;
  -moz-border-radius: 5px;
  border-radius: 5px;
  -webkit-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2);
  -moz-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2);
  box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2);
  -webkit-background-clip: padding-box;
  -moz-background-clip: padding;
  background-clip: padding-box;
  *border-right-width: 2px;
  *border-bottom-width: 2px;
}

.csTable { border: 1px solid #999999; ; font-size: 13px; }

.csTable tr { height: 28px; }

.csTable tr:nth-child(odd){
  background-color: #eaeaea;
  color: #0;
}

.csTableHdr {
  background-color: #cccccc !important;
}

.csSep {
	border-top: 1px solid #cccccc;
}

.csTblForm { width: 1300px !important; border: 0px solid #cccccc; }

	</style>
</head>
<body>

	<div class="toolbar">
		<h1 id="pageTitle"></h1>
		<a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Home</a>
	</div>

<ul id="screen1" title="Material Planning" selected="true">
	<form method="post" name="fMain">
	<input type="hidden" name="PageAction" value="">
	<input type="hidden" name="Job_Select" value="">
	<input type="hidden" name="AutoPost" value="TRUE">
	<input type="hidden" name="MaterialID" value="">

<div class="csForm jqAddJob" style="display: none;">
	<div class="csFormContainer">
	<table>
		<tr>
			<td>Job:</td>
			<td><input type="text" name="AddJob_Name" maxlength="3" class="csLargeInput"></td>
		</tr>
		<tr>
			<td>Floors:</td>
			<td><input type="text" name="AddJob_Floors"></td>
		</tr>
		<tr>
			<td>Windows:</td>
			<td><input type="text" name="AddJob_Windows"></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td></td>
		</tr>
		<tr>
			<td></td>
			<td valign="middle"><a class="csButton" type="cancel" href="javascript: void()" onclick="AddJobHide();" target="_self" title="Add Job">Cancel</a>&nbsp; <a class="csButton" type="cancel" href="javascript: void()" onclick="AddJobSave();" target="_self" title="Add Job">&nbsp;&nbsp;&nbsp;Save&nbsp;&nbsp;&nbsp;</a></td>
		</tr>
	</table>
</div>
</div>

	<li>&nbsp;
	<div style="padding-left: 10px;">
		<table xclass="csTblForm" border="0">
			<tr>
				<td class="csLabel">&nbsp;Select a Job:</td>
				<td style="width: 330px;">
					<select name="Jobs" class="jqJobs csLargeInput">
						<option>Select Job</option>
<%

Set rs_Data = Server.CreateObject("ADODB.Recordset")
rs_Data.CursorType = GetDBCursorType
rs_Data.LockType = GetDBLockType
rs_Data.Open "SELECT * FROM _qws_Planning_Jobs ORDER BY JobName ASC", cn_SQL

Do While Not rs_Data.EOF
	Dim str_Selected: str_Selected = ""
	If str_Job = rs_Data.Fields("JobName").Value Then
		str_Selected = " selected='selected' "
		str_Floors = rs_Data.Fields("Floors").Value
		str_Windows = rs_Data.Fields("Windows").Value
	End If
	Response.Write("<option value='" & rs_Data.Fields("JobName").Value & "' " & str_Selected & ">" & rs_Data.Fields("JobName").Value & "</option>")
	rs_Data.MoveNext
Loop

rs_Data.Close()
Set rs_Data = Nothing
%>
				</select>
				</td>
<% If str_Job <> "" Then %>
		<td class="csLabel">Select Colour: </td>
		<td align="left">
			<select name="Colours" class="jqColours">
<%

Set rs_Data = Server.CreateObject("ADODB.Recordset")
rs_Data.CursorType = GetDBCursorType
rs_Data.LockType = GetDBLockType
rs_Data.Open "SELECT * FROM Y_Color WHERE Job='" & str_Job & "' ORDER BY Project ASC", cn_SQL

If str_Colour = "" Then
	If Not rs_Data.EOF Then
		str_Colour = rs_Data("Project")
	End If
End If

Do While Not rs_Data.EOF
	str_Selected = ""
	If UCase(rs_Data("Project")) = UCase(str_Colour) Then str_Selected = " selected "
	Response.Write("<option value='" & rs_Data("Project") & "' " & str_Selected & ">" & rs_Data("Project") & "</option>")
	rs_Data.MoveNext
Loop
rs_Data.Close()
Set rs_Data = Nothing
%>
			</select>
		</td>
<% End If %>
				<td style="width: 40%;"></td>
				<td><a class="button rightButton" type="cancel" href="javascript: void()" onclick="AddJobShow();" target="_self" title="Add Job">Add New Job</a></td>
			</tr>
		</table>
	</div>
	</li>
	<li class="jqJobInfo">
<div style="padding-left: 10px;">
	<div>&nbsp;</div>
	Job Information: 
<table class="csTblForm" border="0">
	<tr>
		<td class="csLabel">Job:&nbsp;</td><td style="width: 330px;"><input type="text" name="Edit_Job" value="<%= str_Job %>" class="csLargeInput" /></td>
		<td class="csLabel">Floors:&nbsp;</td><td class="csData"><input type="text" name="Edit_Floors" value="<%= str_Floors %>"></td>
		<td class="csLabel">Windows:&nbsp;</td><td class="csData"><input type="text" name="Edit_Windows" value="<%= str_Windows %>"></td>
		<td valign="middle">&nbsp;&nbsp;&nbsp;&nbsp;<a class="csButton" type="cancel" href="javascript: void()" onclick="EditJobSave();" target="_self" title="Save">Save</a></td>
		<td></td>
		<td></td>
	</tr>
</table>
</div>
</li>
<li class="jqJobMaterialEstimate">
<div>&nbsp;</div>
<div style="padding-left: 10px;">
	Materials Estimate:  Add New
<table  class="csTblForm" border="0">
	<tr>
		<td class="csLabel">Part:&nbsp;</td><td style="width: 330px;"><input type="text" name="AddMat_Part" value="" class="jqAddMat_Part csLargeInput" autocomplete="off" /></td>
		<td class="csLabel">Size:&nbsp;</td><td class="csData"><input type="text" name="AddMat_Size" value=""></td>
		<td class="csLabel">Qty:&nbsp;</td><td class="csData"><input type="text" name="AddMat_Qty" value=""></td>
		<td valign="middle">&nbsp;&nbsp;&nbsp;&nbsp;<a class="csButton" type="cancel" href="javascript: void()" onclick="AddMaterialSave();" target="_self" title="Add Material">Add Material</a></td>
	</tr>
</table>
<div style='xpadding-left: 100px;'>
<div>&nbsp;</div>
<table style="width: 70%;" >
	<tr>
		<td>Parts List:</td>
	</tr>
</table>
<table style="width: 70%;" class="csTable">
	<tr class="csTableHdr"><td>&nbsp;Colour</td><td>Part</td><td>Description</td><td>Size</td><td>Qty</td><td>Action</td></tr>
<%

Set rs_Data = Server.CreateObject("ADODB.Recordset")
rs_Data.CursorType = GetDBCursorType
rs_Data.LockType = GetDBLockType
rs_Data.Open "SELECT * FROM _qws_Planning_Materials WHERE Colour='" & str_Colour & "' ORDER BY PART ASC", cn_SQL

Do While Not rs_Data.EOF
%>
<tr>
	<td>&nbsp;<%= rs_Data("Colour") %></td>
	<td><%= rs_Data("Part") %></td>
	<td><%= rs_Data("Description") %></td>
	<td><%= rs_Data("Size") %></td>
	<td><%= rs_Data("Qty") %></td>
	<td><a href="javascript: void();" onclick="MaterialDelete(<%= rs_Data("ID") %>, '<%= rs_Data("Part") %>');">Delete</a></td>
</tr>
<%

	rs_Data.MoveNext
Loop

rs_Data.Close()
Set rs_Data = Nothing

%>
</table>
</div>
</div>
</li>
	</form>
</ul>

	<script>

	var AutoParts = [
<%
Set rs_Data = Server.CreateObject("ADODB.Recordset")
rs_Data.CursorType = GetDBCursorType
rs_Data.LockType = GetDBLockType
rs_Data.Open "SELECT Part, Description FROM Y_Master ORDER BY Part Asc", cn_SQL
Dim b_First: b_First = True
Do While Not rs_Data.EOF
	str_Sep = ","
	If b_First Then str_Sep = ""
	Response.Write(str_Sep & """" & rs_Data("Part") & ", " & Replace(Replace(rs_Data("Description") & "", ",", ""),"""", "") & """")
	b_First = false
	rs_Data.MoveNext
Loop
rs_Data.Close()
Set rs_Data = Nothing
%>
	];

	$(document).ready(function() {
		$(".jqJobInfo").hide();
		$(".jqJobMaterialEstimate").hide();
		$("form").attr('autocomplete', 'off');

<% If str_Job <> "" Then %>
	$(".jqJobInfo").show();
	$(".jqJobMaterialEstimate").show();
	$(".jqAddMat_Part").focus();
<% End If %>

		$(".jqAddMat_Part").autocomplete({
			minLength: 1,
			delay: 250,
			source: AutoParts
		});

	});

	function showProgressBar() {
		$(".jqProgress").show();
	}

	function AddJobShow() {
		$(".jqAddJob").css({
			"left" : parseInt(($(window).width() / 2) - ($(".jqAddJob").width() / 2), 10)
		});

			$(".jqAddJob").show();
	}

	function AddJobHide() {
		$(".jqAddJob").hide();
	}

	function AddJobSave() {
		fMain.AutoPost.value = "";
		fMain.PageAction.value = "ADD_JOB";
		fMain.submit();
	}

	function EditJobSave() {
		fMain.AutoPost.value = "";
		fMain.PageAction.value = "EDIT_JOB";
		fMain.submit();
	}

	function AddMaterialSave() {
		fMain.AutoPost.value = "";
		fMain.PageAction.value = "ADD_MATERIAL";
		fMain.submit();
	}

	function MaterialDelete(str_ID, str_Part) {
		if (confirm('Delete part: ' + str_Part + '?')) {
			fMain.AutoPost.value = "";
			fMain.MaterialID.value = str_ID;
			fMain.PageAction.value = "MATERIAL_DELETE";
			fMain.submit();
		}
	}

	$(".jqJobs").change(function() {
			fMain.AutoPost.value = "";
			fMain.PageAction.value = "SELECT_JOB";
			fMain.Job_Select.value = $(".jqJobs").val();
			fMain.submit();
	});

	$(".jqColours").change(function() {
			fMain.AutoPost.value = "";
			fMain.PageAction.value = "SELECT_COLOURS";
			fMain.submit();
	});

	</script>
	
<%
DBConnection.close
Set DBConnection = nothing
%>

</body>
</html>