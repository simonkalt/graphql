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


	If Csng(Request.Form("CampaignID")) >0 Then
		strCampaignID=Clng(Request.Form("CampaignID"))
	Else
		strCampaignID=Clng(Request.QueryString("CampaignID"))
	End If

%>
	<!--#include file="loginadmin.asp"-->
<%


	'fill drop down list box
	'strSQL="SELECT * FROM GetAllCampaigns"
	'strSQL="SELECT Campaigns.*, Advertisers.CompanyName FROM Advertisers INNER JOIN Campaigns ON Advertisers.AdvertiserID = Campaigns.AdvertiserID Where Advertisers.UserID=" & CLng(Session("BanManProSiteID"))
	
	'Set rs=connBanManPro.Execute(strSQL)

	strTask=Request.QueryString("Task")
        
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Ban Man Pro -- Campaigns</title>
</head>

<body>
<!--#include file="topbanner.htm"-->
<div align="center"><center>

<table border="0" cellpadding="0" cellspacing="0" width="676" background="images/back.gif">
  <tr>
    <td width="100%"><!--webbot bot="ImageMap" rectangle="(549,10) (643, 31)  zones.asp" rectangle="(203,8) (308, 30)  banners.asp" rectangle="(25,11) (136, 30)  advertisers.asp" src="images/campaigns.gif" border="0" startspan --><MAP NAME="FrontPageMap"><AREA SHAPE="RECT" COORDS="549, 10, 643, 31" HREF="zones.asp"><AREA SHAPE="RECT" COORDS="203, 8, 308, 30" HREF="banners.asp"><AREA SHAPE="RECT" COORDS="25, 11, 136, 30" HREF="advertisers.asp"></MAP><a href="../_vti_bin/shtml.exe/Banners/campaigns.asp/map"><img ismap usemap="#FrontPageMap" border="0" height="49" src="images/campaigns.gif" width="677"></a><!--webbot bot="ImageMap" endspan i-checksum="57614" --><br>
    </font>
    </p>
<p align="center">
<!--#include file="datetime.asp"-->
</p>
<%
	'check if user has logged in
	If UCase(Session("UserName"))=UCase(Application("AdministratorName")) And UCase(Session("Password"))=UCase(Application("AdministratorPassword")) Then
		If Request.QueryString("Task")="Delete" And Request.QueryString("Confirm")="True" Then
			%> <p align="center">Are you sure you want to delete this campaign?<p align="center"> <a href="campaigns.asp?Task=Delete&amp;CampaignID=<%=Request.QueryString("CampaignID")%>&amp;Confirm=False"><img src="images/delsmall.gif" alt="Delete Campaign" border="0" WIDTH="38" HEIGHT="18"></a></p><%
		Else
%> <!--#include file="addcampaign.asp"--> <%
		End If

	Else
%><!--#include file="login.asp"--><%
	End If
%>    
    <p align="center"><a href="campaigns.asp?Task=AddNew"><img src="images/addcampaignblue.gif" alt="Add Campaign" border="0" WIDTH="154" HEIGHT="34"></a><%If Trim(strCampaignID)<>"0" Then%><a href="campaigns.asp?Task=Delete&amp;CampaignID=<%=strCampaignID%>&amp;Confirm=True"><img src="images/delblue.gif" alt="Delete Campaign" border="0" WIDTH="91" HEIGHT="34"></a><% End If%><a href="campaigns.asp?Task=ViewAll"><img src="images/viewallblue.gif" alt="View All Campaigns" border="0" WIDTH="102" HEIGHT="34"></a><a href="campaigns.asp?Task=Expired"><img src="images/viewexpired.gif" width="137" height="31" alt="View Expired Campaigns" border="0"></a>
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

