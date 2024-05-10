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

	'Error Traps
	If Trim(Request.Form("UserName"))="" Then
		Response.Write "User Name is Required"
		Response.End
	End If
	If Trim(Request.Form("Password1"))="" Then
		Response.Write "Password is Required"
		Response.End
	End If
	If Request.Form("Password1") <> Request.Form("Password2") Then
		Response.Write "Passwords must be identical in both fields."
		Response.End
	End If
	If UCase(Request.Form("UserName")) = "ADMIN" Or UCase(Request.Form("Password1")) = "ADMIN" Then
		Response.Write "You cannot use a username or password of ADMIN"
		Response.End
	End If

'If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then

	'Get ID for this user
	strSQL2="SELECT * FROM Administrative WHERE Administrative.[AdministratorName] = '" & Application("AdministratorName") & "' AND Administrative.[AdministratorPassword] = '" & Application("AdministratorPassword") & "'"
	Set rs=connBanManPro.Execute(strSQL2)

	If Not rs.EOF Then
		'Update database
 		strSQL2="UPDATE Administrative SET "
		strSQL2=strSQL2 & "AdministratorName='" &  FixBlank(Request.Form("UserName")) & "',"  
		strSQL2=strSQL2 & "AdministratorPassword='" &  FixBlank(Request.Form("Password1")) & "',"
		strSQL2=strSQL2 & "AdministratorEmail='" &  FixBlank(Request.Form("AdministratorEmail")) & "',"
		strSQL2=strSQL2 & "DomainURL='" &  FixBlank(Request.Form("DomainURL")) & "'," 
		strSQL2=strSQL2 & "MailProgram='" &  FixBlank(Request.Form("MailProgram")) & "'," 
		strSQL2=strSQL2 & "EmailWhenCampaignExpires=" &  SetTrueFalse(Request.Form("EmailWhenCampaignExpires")) & "," 
		strSQL2=strSQL2 & "ServerPath='" &  FixBlank(Request.Form("ServerPath")) & "'," 
		strSQL2=strSQL2 & "MailServer='" &  FixBlank(Request.Form("MailServer")) & "'," 
		strSQL2=strSQL2 & "CacheBustingMode=" &  SetTrueFalse(Request.Form("CacheBustingMode")) & ","
		'version 2.0 parameters
 		strSQL2=strSQL2 & "DateFormat='" &  Request.Form("DateFormat") & "',"
		strSQL2=strSQL2 & "UniqueClickHour=" &  Request.Form("UniqueClickHour") & ","
		'strSQL2=strSQL2 & "DatabaseUpdateFrequency=" &  Request.Form("DatabaseUpdateFrequency") & ", "
		strSQL2=strSQL2 & "DailyReport=" &  SetTrueFalse(Request.Form("DailyReport")) & ","
		strSQL2=strSQL2 & "WeeklyReport=" &  SetTrueFalse(Request.Form("WeeklyReport")) & ","
		strSQL2=strSQL2 & "SmoothingMinutes=" & Clng(Request.Form("SmoothingMinutes")) & ","
		strSQL2=strSQL2 & "ZoneAverageDays=" & Clng(Request.Form("ZoneAverageDays")) & ", "
		strSQL2=strSQL2 & "SlotOption=" & SetTrueFalse(Request.Form("SlotOption")) & ", "
		strSQL2=strSQL2 & "StandardCampaignLength=" & Clng(Request.Form("StandardCampaignLength")) & ", "
		If IsNumeric(Request.Form("GuaranteedImpressionsPerSlot")) Then
			lngGuaranteed=Clng(Request.Form("GuaranteedImpressionsPerSlot"))
		Else
			lngGuaranteed=0
		End If
		strSQL2=strSQL2 & "GuaranteedImpressionsPerSlot=" & lngGuaranteed & " "
		'end version 2.0 parameters
		strSQL2=strSQL2 & "WHERE Administrative.[UserID] =" & rs("UserID")

		connBanManPro.Execute strSQL2

		'Update Reports Available to advertisers
		strSQL2="Update BanManProReports Set "
		strSQL2=strSQL2 & "Reports_SummaryByDay=" & SetTrueFalse(Request("Reports_SummaryByDay")) & ","
		strSQL2=strSQL2 & "Reports_SummaryByBanner=" & SetTrueFalse(Request("Reports_SummaryByBanner")) & ","
		strSQL2=strSQL2 & "Reports_SummaryByBannerByDay=" & SetTrueFalse(Request("Reports_SummaryByBannerByDay")) & ","
		strSQL2=strSQL2 & "Reports_SummaryByZone=" & SetTrueFalse(Request("Reports_SummaryByZone")) & ","
		strSQL2=strSQL2 & "Reports_SummaryByZoneByDay=" & SetTrueFalse(Request("Reports_SummaryByZoneByDay")) & ","
		strSQL2=strSQL2 & "Reports_ClickDetail=" & SetTrueFalse(Request("Reports_ClickDetail"))
		connBanManPro.Execute strSQL2

		'update session information
		Session("UserName")=Request.Form("UserName")
		Session("Password")=Request.Form("Password1")

		'reset Application Variables
		Application("AdministratorName")=""
		%>
		<!--#include file="dbconnect.asp"-->
		<%

		%>
		<!--#include file="banmanfunc.asp"-->
		<%
		'update zone averages
		If IsNumeric(Application("ZoneAverageDays")) Then
			UpdateBanManProZoneAverages Application("ZoneAverageDays")
		End If

	Else
		Response.Write "User not found"
	End If
'End If

	Response.Redirect "Admin.asp"
%>

<% ''''''''''''''''''''Change blank fields to " "   '''''''''''''''''''''''''''''''''''''''''''
Function FixBlank(strParameter)
	If Trim(strParameter)="" Then
		FixBlank=" "
	Else
		FixBlank=Replace(strParameter, "'", "''")
	End If
End Function %>

<% ''''''''''''''''''''Set False Check box to 0 " "   '''''''''''''''''''''''''''''''''''''''''''
Function SetTrueFalse(strParameter)
	If Trim(strParameter)="-1" Then
		SetTrueFalse=-1
	Else
		SetTrueFalse=0
	End If
End Function %>