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
	
	<!--#include file="dbconnect.asp"-->
	
<%
	'connect to database
	Set connBanManPro=Server.CreateObject("ADODB.Connection") 
	connBanManPro.Mode = 3      '3 = adModeReadWrite
	connBanManPro.Open Application("BannerManagerConnectString")

	If Request.QueryString("Login")="True" then
		If Trim(Request.Form("lstWebSites"))<>"" Then
			Session("BanManProSiteID")=CLng(Request.Form("lstWebSites"))
		Else
			Session("BanManProSiteID")=0
		End If
		'Determine if this is an Advertiser Logging In
		If Trim(Request.Form("UserName")) <> "" Then
			strSQL="SELECT * FROM Advertisers WHERE LoginName IN ('" & Replace(Request.Form("UserName"),"'","''") & "') AND LoginPassword IN ('" &  Replace(Request.Form("Password"),"'","''") & "')"
			Set rsAdvertiser=connBanManPro.Execute(strSQL)
			If rsAdvertiser.EOF <> True Then
				Session("AdvertiserName")=Request.Form("UserName")
				Session("AdvertiserPassword")=Request.Form("Password")
				Session("AdvertiserID")=rsAdvertiser("AdvertiserID")
				If Request.Form("StoreCookie")="ON" Then
					Response.Cookies ("BanManPro")("AdvertiserName") = Session("AdvertiserName")
					Response.Cookies ("BanManPro")("AdvertiserPassword") = Session("AdvertiserPassword")
				Else
					Response.Cookies ("BanManPro")("AdvertiserName") = ""
					Response.Cookies ("BanManPro")("AdvertiserPassword") = ""
				End If
				Response.Cookies ("BanManPro").Expires=Now + 180
			End If
		End If

		Session("UserName")=Request.Form("UserName")
		Session("Password")=Request.Form("Password")
	End If


%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Ban Man Pro -- Stats/Reports</title>
</head>

<body>
<!--#include file="topbanner.htm"-->
<div align="center"><center>

<table border="0" cellpadding="0" cellspacing="0" width="676" background="images/back2.gif">
  <tr>
    <td><p align="center"><img src="images/statsandreports.gif" WIDTH="676" HEIGHT="52"> <br>
<%
	'check if user has logged in
	If (UCase(Session("UserName"))=Ucase(Application("AdministratorName")) And UCase(Session("Password"))=UCase(Application("AdministratorPassword"))) OR (UCase(Session("UserName"))=UCase(Session("AdvertiserName")) And UCase(Session("Password"))=UCase(Session("AdvertiserPassword"))) AND (Session("UserName")<>"" AND Session("Password")<> "") Then
%>
<p align="center">
<!--#include file="datetime.asp"-->
</p>
<!--#Include File="reportsum.asp"-->
<!--#include file="stats.asp"-->
<% If Not rsReport1.EOF Then %>
<!--#include file="chart.asp"-->
<% End If %>
<!--#include file="detreports.asp"-->
<%
	Else
blnAdvertiserLogin=True
%><!--#include file="login.asp"--><%
	End If
%>    
    <p align="center"><a href="help/stats.htm" target="_blank"><img src="images/Help.gif" border="0" WIDTH="77" HEIGHT="33"></a><br>
    <img src="images/bot2.gif" WIDTH="676" HEIGHT="19"></td>
  </tr>
</table>
</center></div><%If Session("AdvertiserID")=0 Then%><!--#include file="botnavigation.asp"--><%End If%>

</body>
</html>
<% 	Set rs=Nothing
	Set rss=Nothing
	Set rsReport1=Nothing
	connBanManPro.Close()
%>