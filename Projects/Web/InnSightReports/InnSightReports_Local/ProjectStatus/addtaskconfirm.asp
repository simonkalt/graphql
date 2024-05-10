<%@ language=vbscript %>

<!-- #include file="../Include/vbfunc.asp" -->
<!-- #include file="../Data/adovbs.asp" -->

<%
dim cnMain, strSql
set cnMain = server.CreateObject("ADODB.connection")
cnMain.Open Application("sqlProjectStatus_ConnectionString"), Application("sqlProjectStatus_RuntimeUsername"), Application("sqlProjectStatus_RuntimePassword")
if IsEmpty(Request.QueryString("TaskID")) then
	cnMain.Execute "INSERT INTO tblTask (Task, AssignedDeveloperID) VALUES (" & CheckString(Request.Form("txtTask"),"") & ", " & Request.Form("cboAssignedDeveloper") & ")"
else
	strSql = "UPDATE tblTask SET Task = " & CheckString(Request.Form("txtTask"),", ")
	strSQL = strSQL & "AssignedDeveloperID = " & Request.Form("cboAssignedDeveloper") & ", "
	strSQL = strSQL & "ResolutionDate = CheckString(Request.Form("txtResolutionDate",", ")
	strSQL = strSQL & "Resolution = " & CheckString(Request.Form("txtResolution"),"")
	'cnMain.Execute strSQL
	Response.Write strSQL
end if

do until not cnMain.State = adStateExecuting
	doevents
loop

cnMain.Close
set cnMain = nothing

%>


<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
</BODY>
</HTML>

<script language="javascript">
	window.opener.location.href = "../ProjectStatus/Tasks.asp";
	/* window.opener.refresh(); 
	window.close();*/
</script>