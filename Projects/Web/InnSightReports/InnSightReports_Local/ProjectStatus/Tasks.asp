<% @language="vbscript" %>

<%
dim cnMain, rs
set cnMain = server.CreateObject("ADODB.connection")
set rs = server.CreateObject("ADODB.recordset")
cnMain.Open Application("sqlProjectStatus_ConnectionString"), Application("sqlProjectStatus_RuntimeUsername"), Application("sqlProjectStatus_RuntimePassword")
rs.Open "SELECT t.*, d.Name FROM tblTask t INNER JOIN tblDevelopers d ON t.AssignedDeveloperID = d.DeveloperID ORDER BY t.ResolutionDate, t.DateEntered DESC", cnMain

%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>InnSight Reports Project Status System</title>
</head>

<body bgcolor="#F9D568">

<p><b><font face="Tahoma"><u>InnSight Reports Project Status System</u></font></b></p>
<p>&nbsp;<input type="button" value="Add a New Task" name="cmdAddNewTask" onclick="javascript:window.open('AddEditTask.asp','wndAddTask','top=100,left=100,scrollbars=no,height=350,width=600,titlebar=Add Task,menubar=no,resizable=no')"></p>
	<table border="1" width="100%" cellpadding=3>
	  <tr>
	    <td width="6%">&nbsp;</td>
	    <td width="16%"><font face="Tahoma" size="2"><b>Date Entered</b></font></td>
	    <td width="58%"><font face="Tahoma" size="2"><b>Task</b></font></td>
	    <td width="12%"><font face="Tahoma" size="2"><b>&nbsp;Assigned</b></font></td>
	    <td width="11%"><font face="Tahoma" size="2"><b>Resolved</b></font></td>
	  </tr>
		<%Do Until rs.EOF%>		
			<tr bgcolor=#cccccc>
			  <td width="6%" align="center"><font face="Tahoma" size="2"><a href="#" onclick="javascript:window.open('AddEditTask.asp?TaskID=<%=rs("TaskID")%>','wndAddTask','top=100,left=100,scrollbars=no,height=350,width=600,titlebar=Add Task,menubar=no,resizable=no')">Edit</a></font></td>
			  <td width="16%"><font face="Tahoma" size="2"><%=FormatDateTime(rs("DateEntered"),vbShortDate)%></font></td>
			  <td width="58%"><font face="Tahoma" size="2"><%=rs("Task")%></font></td>
			  <td width="12%"><font face="Tahoma" size="2"><%=rs("Name")%></font></td>
			  <td width="11%" align="center" bgcolor=fffff><img src="images/<%if IsDate(rs("ResolutionDate")) then Response.Write("checked.gif") else Response.Write("UnChecked.gif")%>"></td>
			</tr>
		<%	rs.MoveNext
		Loop%>
	</table>
</body>

</html>

<%
rs.Close
cnMain.Close
set rs = nothing
set cnMain = nothing
%>