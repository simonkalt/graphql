<%@ Language=VBScript %>
<%Response.Expires = 0%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Edit User</title>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub cmdCancel_onclick
	' do nothing and go back
	window.parent.location.href = "UserSetup.asp"
End Sub

-->
</SCRIPT>
</HEAD>
<body bgcolor=silver topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" link="black" vlink="black" alink="black">
<!--#include file = "Header.inc" ---> 

<form action="EditUserConfirm.asp?UserID=<%=Request.QueryString("UserID")%>"  method=post id=form2 name=form2>
<%

	Set cnSQL = Server.CreateObject("ADODB.Connection")
	Set rsUsers = Server.CreateObject("ADODB.Recordset")
	Set rsCompanyUsers = Server.CreateObject("ADODB.Recordset")
	Set rsSuperUser = Server.CreateObject("ADODB.Recordset")
  
	cnSQL.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")

'	Set rsUsers = cnSQL.Execute("SELECT tblUser.UserID, tblUser.UserName, tblUser.LoginName, tblUser.Password, tblUser.Admin from tblUser Where tblUser.UserID = " & Request.QueryString("UserID") & " Order by tblUser.UserName")
	Set rsUsers = cnSQL.Execute("SELECT tblUser.UserID, tblUser.UserName, tblUser.LoginName, tblUser.Password, tblUser.EmailAddress from tblUser Where tblUser.UserID = " & Request.QueryString("UserID") & " Order by tblUser.UserName")
  
	Dim strSQL
	Dim bAdmin 
	Dim bSuperUser 
	Dim strEmailAddress

	bAdmin = 0
	bSuperUser = 0
%>

<% If Not rsUsers.EOF Then %>
	<table align="left" border="0" width="758" cellpadding=1 cellspacing=0>
	  <tr>
	    <td align="right"><font face="Tahoma" size="2">User Name:</font></td>
	    <td> 
			<% Response.Write "<INPUT type=""" & "text" &""" id=txtUserName name=txtUserName size=31 value=" & rsUsers.Fields("UserName") & ">" %>
	    </td>
	  </tr>
	  <tr>
	    <td align="right"><font face="Tahoma" size="2">Login Name:</font></td>
	    <td>
			<% Response.Write "<INPUT type=""" & "text" &""" id=txtLoginName name=txtLoginName size=31 value=" & rsUsers.Fields("LoginName") & ">" %>
	    </td>
	  </tr>
	  <tr>
	    <td align="right"><font face="Tahoma" size="2">Password:</font></td>
	    <td>
			<% Response.Write "<INPUT type=""password""" & "text" &""" id=txtPwd name=txtPwd size=31 value=" & rsUsers.Fields("Password") & ">" %>
	    </td>
	  </tr>


	  <tr>
	    <td align="right"><font face="Tahoma" size="2">Administrator:</font></td>
	    <td> 
			<%
				strEmailAddress = rsUsers.Fields("EmailAddress")
			
				' Get Admin privileges from tblCompanyUser
				strSQL = "SELECT Admin from tblCompanyUser Where UserID = " & Request.QueryString("UserID") & " AND CompanyID = " & Session("CompanyID")
				Set rsCompanyUsers = cnSQL.Execute(strSQL)
			
				if (not rsCompanyUsers.BOF) and (not rsCompanyUsers.EOF) then
					bAdmin = rsCompanyUsers.Fields("Admin")
				end if
			
				if (bAdmin = 0) then
					Response.Write "<SELECT id=cboAdmin name=cboAdmin style=""height: 23; width: 210"">  <OPTION selected value=0>No</OPTION>    <OPTION value=1>Yes</OPTION>     </SELECT>"
				else
					Response.Write "<SELECT id=cboAdmin name=cboAdmin style=""height: 23; width: 210"">  <OPTION value=0>No</OPTION>    <OPTION selected value=1>Yes</OPTION>     </SELECT>"
				end if
			%>
		</td>
	  </tr>


	  <tr>
	    <td align="right"><font face="Tahoma" size="2">Super User:</font></td>
	    <td>
			<% 
				If Session("SuperUser") = 1 Then 
					Response.Write "<SELECT id=cboSuperUser name=cboSuperUser style=""HEIGHT: 23px; WIDTH: 210px"">" 
				else
					Response.Write "<SELECT id=cboSuperUser name=cboSuperUser style=""HEIGHT: 23px; WIDTH: 210px"" disabled>" 
				End If 
				
				' Get SuperUser privileges from tblUser
				strSQL = "SELECT SuperUser FROM tblUser WHERE UserID = " & Request.QueryString("UserID")
				Set rsSuperUser = cnSQL.Execute(strSQL)
		
				if (not rsSuperUser.BOF) and (not rsSuperUser.EOF) then
					bSuperUser = rsSuperUser.Fields("SuperUser")
				end if
		
				if (bSuperUser = 0) then
				    Response.Write "<OPTION selected value=0>No</OPTION> <OPTION value=1>Yes</OPTION> </SELECT>"
				else
				    Response.Write "<OPTION value=0>No</OPTION> <OPTION selected value=1>Yes</OPTION> </SELECT>"
				end if
			%>
		</td>
	  </tr>
	  <tr>
	    <td align="right"><font face="Tahoma" size="2">Email Address:</font></td>
	    <td>
			<% Response.Write "<INPUT type=""" & "text" &""" id=txtEmailAddress name=txtEmailAddress size=31 value=" & strEmailAddress & ">" %>
		</td>
	  </tr>
	  <tr>
		<td align="right">Assigned Hotels:</td>
		<td>
			<table cellpadding=0 cellspacing=0>
				<tr>
					<td>
						<SELECT size=10 id=select1 name=select1>
							<OPTION>Hard Coded Hotel 1</OPTION>
							<OPTION>Hard Coded Hotel 2</OPTION>
							<OPTION>Hard Coded Hotel 3</OPTION>
							<OPTION>Hard Coded Hotel 4</OPTION>
							<OPTION>Hard Coded Hotel 5</OPTION>
							<OPTION>Hard Coded Hotel 6</OPTION>
							<OPTION>Hard Coded Hotel 7</OPTION>
							<OPTION>Hard Coded Hotel 8</OPTION>
							<OPTION>Hard Coded Hotel 9</OPTION>
							<OPTION>Hard Coded Hotel 10</OPTION>
						</SELECT>
					</td>
				</tr>
			</table>
		</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2" align="center"> 
			<INPUT id=submit1 name=submit1 type=Submit value=Submit>&nbsp; <INPUT id=cmdCancel name=cmdCancel type=button value=Cancel>
		</td>
	  </tr>
	</table>

<% End If %>

<INPUT type="hidden" id=txtUserID name=txtUserID value=<%=Request.QueryString("UserID")%>>
</form>


</BODY>
</HTML>
