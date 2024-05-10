<%@ language=vbscript%>

<%
dim cnMain, rs, booNewRec
set cnMain = server.CreateObject("ADODB.connection")
set rs = server.CreateObject("ADODB.recordset")
set rsTask = server.CreateObject("ADODB.recordset")
cnMain.Open Application("sqlProjectStatus_ConnectionString"), Application("sqlProjectStatus_RuntimeUsername"), Application("sqlProjectStatus_RuntimePassword")
Set rs = cnMain.Execute("SELECT * FROM tblDevelopers ORDER BY Name")

if not IsEmpty(Request.QueryString("TaskID")) then
	booNewRec = False
	strTitle = "Edit Task"
	set rsTask = cnMain.Execute("SELECT * FROM tblTask WHERE TaskID = " & Request.QueryString("TaskID"))
	dteDateEntered = FormatDateTime(rsTask("DateEntered"),vbLongDate)
	strTask = rsTask("Task")
	intAssignedDeveloperID = rsTask("AssignedDeveloperID")
	dteResolutionDate = rsTask("ResolutionDate")
	strResolution = rsTask("Resolution")
else
	booNewRec = True
	strTitle = "Add Task"
	dteDateEntered = FormatDateTime(now(),vbLongDate)
	strTask = ""
	intAssignedDeveloperID = 0
	dteResolutionDate = ""
	strResolution = ""
end if
%>

<html>

<head>
<title>Add Task</title>
</head>

<body bgcolor="#f9d568">

<b><u><font face="Tahoma"><%=strTitle%><BR></font></u></b>
<BR>
<form method="post" action="AddTaskConfirm.asp?TaskID<%=("="+Request.QueryString("TaskID"))%>">
<TABLE align=center border=1 cellPadding=1 cellSpacing=1 width="75%" background="" bgColor=silver style="WIDTH: 75%">
  
  <TR>
    <TD>
      <TABLE border=0 height=132 width="100%" id=TABLE1 background="" cellPadding=1 cellSpacing=1 style="HEIGHT: 150px; WIDTH: 558px">
        
        <TR>
          <TD height=19 vAlign=top width="28%"></TD>
          <TD height=19 width="72%">&nbsp;</TD></TR>
        
        <TR>
          <TD height=19 vAlign=top width="28%">
            <P align=right><FONT face=Tahoma 
            size=2><STRONG>Date/Time:</STRONG></FONT></P></TD>
          <TD height=19 width="72%"><FONT face=Tahoma size=2><%=dteDateEntered%></FONT></TD></TR>
        <TR>
          <TD height=19 vAlign=top width="28%">
            <P align=right><STRONG><FONT face=Tahoma size=2>Assigned 
            Developer:</FONT></STRONG></P></TD>
          <TD height=19 width="72%">
				<SELECT name=cboAssignedDeveloper size=1>
					<%
					if not rs.EOF then
						rs.MoveFirst
					end if
					do until rs.EOF
						if cint(intAssignedDeveloperID)=cint(rs("DeveloperID")) then
							strSelected = "selected"
						else
							strSelected = ""
						end if%>
						<option <%=strSelected%> value=<%=rs("DeveloperID")%>><%=rs("Name")%></option>
						<%rs.MoveNext
					loop
					%>
				</SELECT>
				
            </TD></TR>
        <TR>
          <TD height=20 vAlign=top width="28%">
            <P align=right><STRONG><FONT face=Tahoma 
            size=2>Task:</FONT></STRONG></P></TD>
          <TD height=20 width="72%">
            <P><TEXTAREA cols=49 name=txtTask style="HEIGHT: 38px; WIDTH: 411px"><%=strTask%></TEXTAREA>&nbsp;</P></TD></TR>
        <TR>
          <TD height=19 vAlign=top width="28%">Resolution Date:</TD>
          <TD height=19 width="72%"><INPUT type="text" id=txtResolutionDate name=txtResolutionDate value="<%=dteResolutionDate%>"></TD></TR>
        <TR>
          <TD height=19 vAlign=top width="28%">Resolution:</TD>
          <TD height=19 width="72%"><INPUT type="text" id=txtResolution name=txtResolution value="<%=strResolution%>" size=60></TD></TR>
        <TR>
          <TD height=19 vAlign=top width="28%">&nbsp;</TD>
          <TD height=19 width="72%">&nbsp;</TD></TR>
        <TR>
          <TD height=25 vAlign=top width="28%"></TD>
          <TD height=25 width="72%">
            <P align=right><INPUT name=B1 type=submit value=Save>&nbsp;<INPUT name=B2 type=reset value=Reset></P></TD></TR>
</TABLE>&nbsp;</TD></TR></TABLE>
</form>
<DIV align=right>
<table>
		<tr>
			<td>
				<P align=right><INPUT id=cmdClose name=cmdClose type=button value="Close this window" onclick="javascript:window.close()"></P>
			</td>
		</tr>
	</table></DIV>
</body>
</p>
</html>
<%
if not booNewRec then
	rsTask.Close
	set rsTask = nothing
end if
rs.Close
cnMain.Close
set rs = nothing
set cnMain = nothing
%>