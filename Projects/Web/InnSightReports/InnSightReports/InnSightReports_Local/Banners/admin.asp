<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Product:  Ban Man Pro
'   Author:   Joe Rohrbach of Brookfield Consultants
'   Notes:    None
'                  
'
'                         COPYRIGHT NOTICE
'
'   The contents of this file are protected under the United States
'   copyright laws as an unpublished work, and are confidential and
'   proprietary to Brookfield Consultants.  Its use or disclosure in 
'   whole or in part without the expressed written permission of 
'   Brookfield Consultants is prohibited.
'
'   (c) Copyright 1999 by Brookfield Consultants.  All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


%>
	<!--#include file="loginadmin.asp"-->
<%


	strSQL="SELECT * FROM Administrative"
	Set rs=connBanManPro.Execute(strSQL)      

	strSQL="SELECT * FROM BanManProReports "
	Set rsReports=connBanManPro.Execute(strSQL)      

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Ban Man Pro -- Administration</title>
</head>

<body>
<!--#include file="topbanner.htm"-->
<div align="center"><center>

<table border="0" cellpadding="0" cellspacing="0" width="676" background="images/back2.gif">
  <tr>
    <td><p align="center"><img src="images/administration.gif" WIDTH="676" HEIGHT="50"> <br>
<%
	'check if user has logged in
	If UCase(Session("UserName"))=UCase(Application("AdministratorName")) And UCase(Session("Password"))=UCase(Application("AdministratorPassword")) Then
%><!--#include file="admindata.asp"--><%
	Else
%><!--#include file="login.asp"--><%
	End If
%>    </p>
    <p align="center"><a href="help/preferences.htm" target="_blank"><img src="images/Help.gif" border="0" WIDTH="77" HEIGHT="33"></a>
    <a href="logout.asp"><img border="0" src="images/Logout.gif" WIDTH="85" HEIGHT="33"></a><br>
    <img src="images/bot2.gif" WIDTH="676" HEIGHT="19"></td>
  </tr>
</table>
</center></div><!--#include file="botnavigation.asp"-->

</body>
</html>
<% 	Set rs=Nothing
	Set rss=Nothing
	connBanManPro.Close()
%>