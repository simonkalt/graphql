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
<title>Ban Man Pro -- Zones</title>
</head>

<body>
<!--#include file="topbanner.htm"-->
<div align="center"><center>

<table border="0" cellpadding="0" cellspacing="0" width="676" background="images/back.gif">
  <tr>
    <td width="100%"><!--webbot bot="ImageMap" rectangle="(366,14) (476, 32)  campaigns.asp" rectangle="(198,13) (304, 30)  banners.asp" rectangle="(29,11) (134, 33)  advertisers.asp" src="images/zones.gif" border="0" startspan --><MAP NAME="FrontPageMap"><AREA SHAPE="RECT" COORDS="366, 14, 476, 32" HREF="campaigns.asp"><AREA SHAPE="RECT" COORDS="198, 13, 304, 30" HREF="banners.asp"><AREA SHAPE="RECT" COORDS="29, 11, 134, 33" HREF="advertisers.asp"></MAP><a href="../_vti_bin/shtml.exe/Banners/zones.asp/map"><img ismap usemap="#FrontPageMap" border="0" height="51" src="images/zones.gif" width="677"></a><!--webbot bot="ImageMap" endspan i-checksum="59802" -->
      </font>
    </p>
<p align="center">
<!--#include file="datetime.asp"-->
</p>
    <p align="center"><%
	'check if user has logged in
	If UCase(Session("UserName"))=UCase(Application("AdministratorName")) And UCase(Session("Password"))=UCase(Application("AdministratorPassword")) Then
		If Request.QueryString("Task")="Delete" And Request.QueryString("Confirm")="True" Then
			%> <p align="center">Are you sure you want to delete this zone?<p align="center"> <a href="zones.asp?Task=Delete&amp;ZoneID=<%=Request.QueryString("ZoneID")%>&amp;Confirm=False"><img src="images/delsmall.gif" alt="Delete Zone" border="0" WIDTH="38" HEIGHT="18"></a></p><%
		Else
%> <!--#include file="addzone.asp"--> <%
		End If
	Else
%> <!--#include file="login.asp"--> <%
	End If
%> 
      <p align="center"><a href="zones.asp?Task=AddNew"><img src="images/addzoneblue.gif" alt="Add Zone" border="0" WIDTH="114" HEIGHT="34"></a><%If Trim(strZoneID)<>"0" Then%><a href="zones.asp?Task=Delete&amp;ZoneID=<%=strZoneID%>&amp;Confirm=True"><img src="images/delblue.gif" alt="Delete Zone" border="0" WIDTH="91" HEIGHT="34"></a><% End If%><a href="zones.asp?Task=ViewAll"><img src="images/viewallblue.gif" alt="View All Zones" border="0" WIDTH="102" HEIGHT="34"></a><a href="logout.asp"><img border="0" src="images/Logout.gif" WIDTH="85" HEIGHT="33"></a><br>
    <img src="images/botpiece.gif" WIDTH="677" HEIGHT="21"></td>
  </tr>
</table>
</center></div><!--#include file="botnavigation.asp"-->

</body>
</html>
<% 	
	connBanManPro.Close()
%>

