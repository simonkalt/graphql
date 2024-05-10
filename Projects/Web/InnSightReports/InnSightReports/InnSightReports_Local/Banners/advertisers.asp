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



	If Csng(Request.Form("AdvertiserID")) >0 Then
		strAdvertiserID=CLng(Request.Form("AdvertiserID"))
	Else
		strAdvertiserID=CLng(Request.QueryString("AdvertiserID"))
	End If

%>
	<!--#include file="loginadmin.asp"-->
<%

	'fill drop down list box
	strSQL="SELECT * FROM Advertisers Where UserID=" & CLng(Session("BanManProSiteID")) & " ORDER BY Advertisers.[CompanyName] ASC"
	Set rs=connBanManPro.Execute(strSQL)

	strTask=Request.QueryString("Task")
        

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Ban Man Pro -- Advertisers</title>
</head>

<body>
<!--#include file="topbanner.htm"-->
<div align="center"><center>

<table border="0" cellpadding="0" cellspacing="0" width="676" background="images/back.gif">
  <tr>
    <td background="images/back.gif"><map name="FPMap0">
      <area href="zones.asp" shape="rect" coords="543, 10, 640, 30">
      <area href="campaigns.asp" shape="rect" coords="366, 12, 475, 31">
      <area href="banners.asp" shape="rect" coords="202, 13, 310, 29"></map><img rectangle="(202,13) (310, 29)  banners.asp" src="images/advertisers.gif" border="0" usemap="#FPMap0" WIDTH="677" HEIGHT="49"><br>
    <br>   
<!--#include file="datetime.asp"-->
    <p><%
	'check if user has logged in
	If UCase(Session("UserName"))=UCase(Application("AdministratorName")) And UCase(Session("Password"))=UCase(Application("AdministratorPassword")) Then
		If Request.QueryString("Task")="Delete" And Request.QueryString("Confirm")="True" Then
			%> <p align="center">Are you sure you want to delete this advertiser?<br>This will also delete all banners and campaigns associated with this advertiser.</p><p align="center"> <a href="advertisers.asp?Task=Delete&amp;AdvertiserID=<%=Request.QueryString("AdvertiserID")%>&amp;Confirm=False"><img src="images/delsmall.gif" alt="Delete Advertiser" border="0" WIDTH="38" HEIGHT="18"></a></p><%
		Else
%> <!--#include file="addadvertiser.asp"--> <%
		End If
	Else
%> <!--#include file="login.asp"--> <%
	End If
%> 
      <p align="center"><a href="advertisers.asp?Task=AddNew"><img src="images/addadvertiserblue.gif" alt="Add New Advertiser" border="0" WIDTH="153" HEIGHT="33"></a><%If Trim(strAdvertiserID)<>"0" Then%><a href="advertisers.asp?Task=Delete&amp;AdvertiserID=<%=strAdvertiserID%>&amp;Confirm=True"><img src="images/delblue.gif" alt="Delete Advertiser" border="0" WIDTH="91" HEIGHT="34"></a><%End If%><a href="advertisers.asp?Task=ViewAll"><img src="images/viewallblue.gif" alt="View All Advertisers" border="0" WIDTH="102" HEIGHT="34"></a><a href="logout.asp"><img border="0" src="images/Logout.gif" WIDTH="85" HEIGHT="33"></a><br>
    <img src="images/botpiece.gif" WIDTH="677" HEIGHT="21"></td>
  </tr>
</table>
</center></div><!--#include file="botnavigation.asp"-->

</body>
</html>
<% 	connBanManPro.Close()
%>