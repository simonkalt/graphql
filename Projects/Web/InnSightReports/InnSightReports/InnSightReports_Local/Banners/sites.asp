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
'   (c) Copyright 2000 by Brookfield Consultants.  All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	strSiteID=CLng(Request.QueryString("SiteID"))
	strTask=Request.QueryString("Task")
        
%>
	<!--#include file="loginadmin.asp"-->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Ban Man Pro Sites</title>
</head>

<body>
<!--#include file="topbanner.htm"-->
<div align="center"><center>

<table border="0" cellpadding="0" cellspacing="0" width="676" background="images/back.gif">
  <tr>
    <td width="100%">
      <p align="center"><img border="0" src="images/banmanprositemanagement.gif" width="678" height="54">
    <p align="center">
<%
	'check if user has logged in
	If UCase(Session("UserName"))=UCase(Application("AdministratorName")) And UCase(Session("Password"))=UCase(Application("AdministratorPassword")) Then
		If Request.QueryString("Task")="Delete" And Request.QueryString("Confirm")="True" Then
			%> <p align="center">Are you sure you want to delete this site?<p align="center">This will delete all Advertiser,Banners,Campaigns and Zones for this site.<p align="center"> <a href="sites.asp?Task=Delete&amp;SiteID=<%=Request.QueryString("SiteID")%>&amp;Confirm=False"><img src="images/delsmall.gif" alt="Delete Site" border="0" WIDTH="38" HEIGHT="18"></a></p><%
		Else
%> <!--#include file="addsite.asp"--> <%
		End If
	Else
%> <!--#include file="login.asp"--> <%
	End If
%> 

      <p align="center"><a href="sites.asp?Task=AddNew"><img border="0" src="images/AddSiteBlue.gif" WIDTH="137" HEIGHT="33"></a>
      <a href="sites.asp"><img border="0" src="images/viewallblue.gif" WIDTH="102" HEIGHT="34"></a>  <a href="help/sites.htm" target="_blank"><img src="images/Help.gif" border="0" WIDTH="77" HEIGHT="33"></a>
      <a href="logout.asp"><img border="0" src="images/Logout.gif" WIDTH="85" HEIGHT="33"></a><br>
    <img src="images/botpiece.gif" WIDTH="677" HEIGHT="21"></td>
  </tr>
</table>
</center></div><!--#include file="botnavigation.asp"-->

</body>
</html>
<% 	
	connBanManPro.Close()
%>