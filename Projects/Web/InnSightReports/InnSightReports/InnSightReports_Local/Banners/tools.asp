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



	If Csng(Request.Form("ZoneID")) >0 Then
		strZoneID=CLng(Request.Form("ZoneID"))
	Else
		strZoneID=CLng(Request.QueryString("ZoneID"))
	End If

	strTask=Request.QueryString("Task")
        
%>
	<!--#include file="loginadmin.asp"-->


<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Ban Man Pro Tools</title>
</head>

<body>
<!--#include file="topbanner.htm"-->
<div align="center"><center>

<table border="0" cellpadding="0" cellspacing="0" width="676" background="images/back.gif">
  <tr>
    <td width="100%">
      <p align="center"><img border="0" src="images/BanManProTools.gif" WIDTH="678" HEIGHT="54">
    <p align="center"><%
	'check if user has logged in
	If UCase(Session("UserName"))=UCase(Application("AdministratorName")) And UCase(Session("Password"))=UCase(Application("AdministratorPassword")) Then

 If Request("Task")="TemporaryStop" Then
		Application("BMP_Emergency_Stop")=True
		Response.Write "<p align=center>Ad Serving Temporarily Stopped."
	ElseIf Request("Task")="Start" Then
		Application("BMP_Emergency_Stop")=False
		Response.Write "<p align=center>Ad Serving Restarted."
	End If 
	
		If Request("Task")="" Then %>
<p align="center"><a href="tools.asp?Task=Purge"><img border="0" src="images/PurgeDatabase.gif" width="434" height="43" alt="This tool will remove any old statistics from the database.  It will not affect any campaigns which have not been deleted from the database."></a><br>
<% If Application("BMP_Emergency_Stop")<>True Then %>
 <a href="tools.asp?Task=TemporaryStop">
<img border="0" src="images/TemporarilyStopServingAds.gif" alt="Use this tool to temporarily stop serving ads on your web site.  When executed all banner requests will be served a blank 1X1 pixel image.  This feature is useful if your database server temporarily goes down." WIDTH="434" HEIGHT="43"></a><br>
<%Else %>
 <a href="tools.asp?Task=Start">
<img border="0" src="images/StartServingAds.gif" WIDTH="434" HEIGHT="43"></a><br>
<% End If%>
&nbsp;
		<%Else
			If Request("Task")="Purge" Then
			%>
			<!--#include file="purge.asp"-->
			<%
			Else
			End If
		End If
	Else
%> <!--#include file="login.asp"--> <%
	End If
%> 
      <p align="center"><a href="logout.asp"><img border="0" src="images/Logout.gif" WIDTH="85" HEIGHT="33"></a><br>
    <img src="images/botpiece.gif" WIDTH="677" HEIGHT="21"></td>
  </tr>
</table>
</center></div><!--#include file="botnavigation.asp"-->

</body>
</html>
