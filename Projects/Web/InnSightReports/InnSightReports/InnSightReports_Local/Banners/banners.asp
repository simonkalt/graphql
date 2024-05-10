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


	If Csng(Request.Form("BannerID")) >0 Then
		strBannerID=CLng(Request.Form("BannerID"))
	Else
		strBannerID=CLng(Request.QueryString("BannerID"))
	End If

	strTask=Request.QueryString("Task")
        
%>
	<!--#include file="loginadmin.asp"-->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Ban Man Pro -- Banners</title>
</head>

<body>
<!--#include file="topbanner.htm"-->
<div align="center"><center>

<table border="0" cellpadding="0" cellspacing="0" width="676" background="images/back.gif">
  <tr>
    <td width="100%"><!--webbot bot="ImageMap" rectangle="(542,13) (635, 31)  zones.asp" rectangle="(368,10) (473, 32)  campaigns.asp" rectangle="(28,12) (136, 31)  advertisers.asp" src="images/banners.gif" border="0" startspan --><MAP NAME="FrontPageMap"><AREA SHAPE="RECT" COORDS="542, 13, 635, 31" HREF="zones.asp"><AREA SHAPE="RECT" COORDS="368, 10, 473, 32" HREF="campaigns.asp"><AREA SHAPE="RECT" COORDS="28, 12, 136, 31" HREF="advertisers.asp"></MAP><a href="../_vti_bin/shtml.exe/Banners/banners.asp/map"><img ismap usemap="#FrontPageMap" border="0" height="51" src="images/banners.gif" width="677"></a><!--webbot bot="ImageMap" endspan i-checksum="20908" --><br>
    </font>

<p align="center">
<!--#include file="datetime.asp"-->
</p>
    <%
	'check if user has logged in
	If UCase(Session("UserName"))=UCase(Application("AdministratorName")) And UCase(Session("Password"))=UCase(Application("AdministratorPassword")) Then
		If Request.QueryString("Task")="Delete" And Request.QueryString("Confirm")="True" Then
			%> <p align="center">Are you sure you want to delete this banner?<p align="center"> <a href="banners.asp?Task=Delete&amp;BannerID=<%=Request.QueryString("BannerID")%>&amp;Confirm=False"><img src="images/delsmall.gif" alt="Delete Banner" border="0" WIDTH="38" HEIGHT="18"></a></p><%
		Else
%> <!--#include file="addbanner.asp"--> <%
		End If

	Else
%> <!--#include file="login.asp"--> <%
	End If
%> 
    <p align="center"><a href="banners.asp?Task=AddNew"><img src="images/addbannerblue.gif" alt="Add Banner" border="0" WIDTH="148" HEIGHT="34"></a><a href="banners.asp?Task=Advanced"><img src="images/Advanced.gif" align="absmiddle" border="0" WIDTH="144" HEIGHT="28" alt="Add A New Banner using your own banner code."></a><%If Trim(strBannerID)<>"0" Then%><a href="banners.asp?Task=Delete&amp;BannerID=<%=strBannerID%>&amp;Confirm=True"><img src="images/delblue.gif" alt="Delete Banner" border="0" WIDTH="91" HEIGHT="34"></a><%End If%><a href="banners.asp?Task=ViewAll"><img src="images/viewallblue.gif" alt="View All Banners" border="0" WIDTH="102" HEIGHT="34"></a><a href="logout.asp"><img border="0" src="images/Logout.gif" WIDTH="85" HEIGHT="33"></a><br>
    <img src="images/botpiece.gif" WIDTH="677" HEIGHT="21"></td>
  </tr>
</table>
</center></div><!--#include file="botnavigation.asp"-->

</body>
</html>
<% 	

	connBanManPro.Close()
%>


