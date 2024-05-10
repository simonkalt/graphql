<%
	If Request("ReportFormat")="EXCEL" Then
		Response.Buffer=True
		Response.ContentType = "application/vnd.ms-excel"
  	End If
%>
<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Product:  Ban Man Pro
'   Author:   Joe Rohrbach of Brookfield Consultants
'   Notes:    Main Report Creator Module
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

	'get parameters from form
	'DATE INFORMATION
	If Application("DateFormat")="MM/DD/YYYY"  Then
		'US Date
		strStartDate=DateValue(DateSerial(Request("StartYear"),Request("StartMonth"),Request("StartDay")))
		strEndDate=DateValue(DateSerial(Request("EndYear"),Request("EndMonth"),Request("EndDay")))
	ElseIf Application("DateFormat")="DD/MM/YYYY" Then

		strStartDate=(Request("StartDay") & "/" & Request("StartMonth") & "/" & Request("StartYear"))
		strEndDate=(Request("EndDay") & "/" & Request("EndMonth") & "/" & Request("EndYear"))

	End If

	strStartDateSQL="CONVERT(DATETIME,'" & Request("StartMonth") & "/" & Request("StartDay") & "/" & Request("StartYear") & "',101)"
	strEndDateSQL="CONVERT(DATETIME,'" & Request("EndMonth") & "/" & Request("EndDay") & "/" & Request("EndYear") & "',101)"

	'CAMPAIGN
	strCampaignID=Request("Campaign")
	If strCampaignID="All Campaigns" Then
		blnAllCampaigns=True
	Else
		blnAllCampaigns=False
	End If

	'REPORT TYPE
	strReportType=Request("ReportType")

	If CLng(Session("BanManProSiteID"))<>0 Then
		strExtra="AND Impressions.UserID=" & CLng(Session("BanManProSiteID"))
	Else
		strExtra=""
	End If

	Select Case strReportType
		Case "Summary By Day"
			'CampaignStatsByDay
			If blnAllCampaigns=True Then
				strSQL=" SELECT Impressions.CampaignID, Impressions.ImpressionDay, Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Sum(ClickCounts.Clicks) AS SumOfClicks, Impressions.AdvertiserID, Campaigns.CampaignName "
				strSQL=strSQL & " FROM (Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID "
				strSQL=strSQL & " WHERE ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & "))  " & strExtra
				strSQL=strSQL & " GROUP BY Impressions.CampaignID, Impressions.ImpressionDay, Impressions.AdvertiserID, Campaigns.CampaignName"
			Else
				strSQL=" SELECT Impressions.CampaignID, Impressions.ImpressionDay, Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Sum(ClickCounts.Clicks) AS SumOfClicks, Impressions.AdvertiserID, Campaigns.CampaignName "
				strSQL=strSQL & " FROM (Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID "
				strSQL=strSQL & " WHERE (Impressions.CampaignID=" & strCampaignID & ")  AND  ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) "
				strSQL=strSQL & " GROUP BY Impressions.CampaignID, Impressions.ImpressionDay, Impressions.AdvertiserID, Campaigns.CampaignName"
			End If
		Case "Summary By Banner"
			If blnAllCampaigns=True Then
				strSQL=" SELECT Impressions.CampaignID, Impressions.BannerID, Impressions.AdvertiserID, Count(Impressions.ImpressionDay) AS CountOfImpressionDay, Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignName, Banners.AdDescription "
				strSQL=strSQL & " FROM ((Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) INNER JOIN Banners ON Impressions.BannerID = Banners.BannerID) INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID "
				strSQL=strSQL & " WHERE ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) " & strExtra
				strSQL=strSQL & " GROUP BY Impressions.CampaignID, Impressions.BannerID, Impressions.AdvertiserID, Campaigns.CampaignName, Banners.AdDescription "
				strSQL=strSQL & " ORDER BY Campaigns.CampaignName"
			Else
				strSQL=" SELECT Impressions.CampaignID, Impressions.BannerID, Impressions.AdvertiserID, Count(Impressions.ImpressionDay) AS CountOfImpressionDay, Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignName, Banners.AdDescription "
				strSQL=strSQL & " FROM ((Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) INNER JOIN Banners ON Impressions.BannerID = Banners.BannerID) INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID "
				strSQL=strSQL & " WHERE (Impressions.CampaignID=" & strCampaignID & ")  AND ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) "
				strSQL=strSQL & " GROUP BY Impressions.CampaignID, Impressions.BannerID, Impressions.AdvertiserID, Campaigns.CampaignName, Banners.AdDescription "
				strSQL=strSQL & " ORDER BY Campaigns.CampaignName"
			End If
		Case "Summary By Banner By Day"
			If blnAllCampaigns=True Then
				strSQL=" SELECT Impressions.CampaignID, Impressions.BannerID, Impressions.AdvertiserID, Impressions.ImpressionDay, Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignName, Banners.AdDescription "
				strSQL=strSQL & " FROM ((Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) INNER JOIN Banners ON Impressions.BannerID = Banners.BannerID) INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID "
				strSQL=strSQL & " WHERE ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) " & strExtra
				strSQL=strSQL & " GROUP BY Impressions.CampaignID, Impressions.BannerID, Impressions.AdvertiserID, Impressions.ImpressionDay, Campaigns.CampaignName, Banners.AdDescription "
			Else
				strSQL=" SELECT Impressions.CampaignID, Impressions.BannerID, Impressions.AdvertiserID, Impressions.ImpressionDay, Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignName, Banners.AdDescription "
				strSQL=strSQL & " FROM ((Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) INNER JOIN Banners ON Impressions.BannerID = Banners.BannerID) INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID "
				strSQL=strSQL & " WHERE (Impressions.CampaignID=" & strCampaignID & ")  AND  ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) "
				strSQL=strSQL & " GROUP BY Impressions.CampaignID, Impressions.BannerID, Impressions.AdvertiserID, Impressions.ImpressionDay, Campaigns.CampaignName, Banners.AdDescription "
			End If
		Case "Summary By Zone"
			If blnAllCampaigns=True Then
				strSQL=" SELECT Impressions.CampaignID, Impressions.ZoneID, Count(Impressions.ImpressionDay) AS CountOfImpressionDay, Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignName,Zones.ZoneDescription "
				strSQL=strSQL & " FROM ((Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID ) INNER JOIN Zones ON Impressions.ZoneID = Zones.ZoneID "
				strSQL=strSQL & " WHERE ((Impressions.ImpressionDay >= " & strStartDateSQL & ")  AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) " & strExtra
				strSQL=strSQL & " GROUP BY Impressions.CampaignID, Impressions.ZoneID,  Campaigns.CampaignName, Zones.ZoneDescription "
				strSQL=strSQL & " ORDER BY Zones.ZoneDescription"
			Else
				strSQL=" SELECT Impressions.CampaignID, Impressions.ZoneID, Count(Impressions.ImpressionDay) AS CountOfImpressionDay, Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignName,Zones.ZoneDescription "
				strSQL=strSQL & " FROM ((Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID ) INNER JOIN Zones ON Impressions.ZoneID = Zones.ZoneID "
				strSQL=strSQL & " WHERE (Impressions.CampaignID=" & strCampaignID & ") AND ((Impressions.ImpressionDay >= " & strStartDateSQL & ")  AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) "
				strSQL=strSQL & " GROUP BY Impressions.CampaignID, Impressions.ZoneID,  Campaigns.CampaignName, Zones.ZoneDescription "
				strSQL=strSQL & " ORDER BY Zones.ZoneDescription"
			End If
		Case "Summary By Zone By Day"
			If blnAllCampaigns=True Then
				strSQL=" SELECT Impressions.CampaignID,  Impressions.ZoneID, Impressions.ImpressionDay, Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignName, Zones.ZoneDescription "
				strSQL=strSQL & " FROM (((Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) INNER JOIN Banners ON Impressions.BannerID = Banners.BannerID) INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID) INNER JOIN Zones ON Impressions.ZoneID = Zones.ZoneID"
				strSQL=strSQL & " WHERE ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) " & strExtra
				strSQL=strSQL & " GROUP BY Impressions.CampaignID, Impressions.ZoneID, Impressions.ImpressionDay, Campaigns.CampaignName,  Zones.ZoneDescription "
				strSQL=strSQL & " ORDER BY Zones.ZoneDescription,Campaigns.CampaignName,Impressions.ImpressionDay"
			Else
				strSQL=" SELECT Impressions.CampaignID,   Impressions.ZoneID, Impressions.ImpressionDay, Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignName, Zones.ZoneDescription "
				strSQL=strSQL & " FROM (((Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) INNER JOIN Banners ON Impressions.BannerID = Banners.BannerID) INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID) INNER JOIN Zones ON Impressions.ZoneID = Zones.ZoneID"
				strSQL=strSQL & " WHERE (Impressions.CampaignID=" & strCampaignID & ")  AND  ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) "
				strSQL=strSQL & " GROUP BY Impressions.CampaignID, Impressions.ZoneID,  Impressions.ImpressionDay, Campaigns.CampaignName, Zones.ZoneDescription "
				strSQL=strSQL & " ORDER BY Zones.ZoneDescription,Campaigns.CampaignName,Impressions.ImpressionDay"
			End If
		Case "Click Detail"
			'show clicks
			If blnAllCampaigns=True Then
				strSQL="SELECT Clicks.CampaignID, Clicks.ZoneID, Clicks.AdvertiserID, Clicks.BannerID, Clicks.ClickIP, Clicks.ClickHost, Clicks.ClickReferringURL, Clicks.ClickDateTime, Clicks.ClickBrowser, Clicks.ClickScript_Name, Campaigns.CampaignName "
				strSQL=strSQL & "FROM Campaigns LEFT JOIN Clicks ON Campaigns.CampaignID = Clicks.CampaignID "
				strSQL=strSQL & "WHERE ((Clicks.CampaignID>0) AND (( Clicks.ClickDateTime >= " & strStartDateSQL & ") AND   (Clicks.ClickDateTime <= DATEADD(day,1," & strEndDateSQL & ")))) AND Clicks.UserID=" & CLng(Session("BanManProSiteID"))
				strSQL=strSQL & "ORDER BY Clicks.CampaignID, Clicks.ClickDateTime DESC"
			Else
				strSQL="SELECT Clicks.CampaignID, Clicks.ZoneID, Clicks.AdvertiserID, Clicks.BannerID, Clicks.ClickIP, Clicks.ClickHost, Clicks.ClickReferringURL, Clicks.ClickDateTime, Clicks.ClickBrowser, Clicks.ClickScript_Name, Campaigns.CampaignName "
				strSQL=strSQL & "FROM Campaigns LEFT JOIN Clicks ON Campaigns.CampaignID = Clicks.CampaignID "
				strSQL=strSQL & "WHERE ((Clicks.CampaignID=" & strCampaignID & ") AND (( Clicks.ClickDateTime >= " & strStartDateSQL & ") AND   (Clicks.ClickDateTime <= DATEADD(day,1," & strEndDateSQL & ")))) "
				strSQL=strSQL & "ORDER BY Clicks.CampaignID, Clicks.ClickDateTime DESC"
			End If
		Case "Executive"
				strSQL="SELECT Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, "
				strSQL=strSQL & "Sum(ClickCounts.Clicks) AS SumOfClicks "
				strSQL=strSQL & " FROM Impressions INNER JOIN "
				strSQL=strSQL & "ClickCounts ON Impressions.ID = ClickCounts.ClickID "
				strSQL=strSQL & "Where ((Impressions.ImpressionDay >=" & strStartDateSQL & ") AND (Impressions.ImpressionDay <=" & strEndDateSQL & ")) " & strExtra

	
				strSQL2="SELECT Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, "
				strSQL2=strSQL2 & "Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignName, "
				strSQL2=strSQL2 & "Advertisers.CompanyName, Campaigns.CampaignStartDate, "
				strSQL2=strSQL2 & "Campaigns.CampaignEndDate, Campaigns.CampaignType, "
				strSQL2=strSQL2 & "Campaigns.CampaignQuantitySold "
				strSQL2=strSQL2 & "FROM ((Impressions INNER JOIN ClickCounts ON "
				strSQL2=strSQL2 & "Impressions.ID = ClickCounts.ClickID) INNER JOIN Advertisers ON "
				strSQL2=strSQL2 & "Impressions.AdvertiserID = Advertisers.AdvertiserID) INNER JOIN "
				strSQL2=strSQL2 & "Campaigns ON Impressions.CampaignID = Campaigns.CampaignID "
				strSQL2=strSQL2 & "Where (Campaigns.CampaignEndDate >=" & strStartDateSQL & ") " & strExtra
				strSQL2=strSQL2 & "GROUP BY Campaigns.CampaignName, Advertisers.CompanyName, "
				strSQL2=strSQL2 & "Campaigns.CampaignStartDate, Campaigns.CampaignEndDate, "
				strSQL2=strSQL2 & "Campaigns.CampaignType, Campaigns.CampaignQuantitySold "
				strSQL2=strSQL2 & "Order By Campaigns.CampaignEndDate Asc"
				Set rsReportsEx=connBanManPro.Execute(strSQL2)

				strSQL2="SELECT Campaigns.CampaignName, Advertisers.CompanyName, "
				strSQL2=strSQL2 & "Campaigns.CampaignStartDate, Campaigns.CampaignEndDate, Campaigns.CampaignType, "
    				strSQL2=strSQL2 & "Campaigns.CampaignQuantitySold FROM Advertisers INNER JOIN "
    				strSQL2=strSQL2 & "Campaigns ON Advertisers.AdvertiserID = Campaigns.AdvertiserID "
				strSQL2=strSQL2 & "WHERE (Campaigns.CampaignStartDate > GETDATE()) AND Campaigns.UserID=" & CLng(Session("BanManProSiteID"))
				Set rsReportsEx2=connBanManPro.Execute(strSQL2)
		Case "Billing"
			If blnAllCampaigns=True Then
				strSQL="SELECT Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignCost, Campaigns.CampaignType, Campaigns.CampaignName "
				strSQL=strSQL & "FROM (Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID "
				strSQL=strSQL & "WHERE ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) " & strExtra
				strSQL=strSQL & "GROUP BY Campaigns.CampaignCost, Campaigns.CampaignType, Campaigns.CampaignName"
			Else
				strSQL="SELECT Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignCost, Campaigns.CampaignType, Campaigns.CampaignName "
				strSQL=strSQL & "FROM (Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID "
				strSQL=strSQL & " WHERE (Impressions.CampaignID=" & strCampaignID & ")  AND  ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) "
				strSQL=strSQL & "GROUP BY Campaigns.CampaignCost, Campaigns.CampaignType, Campaigns.CampaignName"
			End If
		Case "Expiration"
				strSQL2="SELECT Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, "
				strSQL2=strSQL2 & "Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignName, "
				strSQL2=strSQL2 & "Advertisers.CompanyName, Campaigns.CampaignStartDate, "
				strSQL2=strSQL2 & "Campaigns.CampaignEndDate, Campaigns.CampaignType, "
				strSQL2=strSQL2 & "Campaigns.CampaignQuantitySold,Campaigns.CampaignID "
				strSQL2=strSQL2 & "FROM ((Impressions INNER JOIN ClickCounts ON "
				strSQL2=strSQL2 & "Impressions.ID = ClickCounts.ClickID) INNER JOIN Advertisers ON "
				strSQL2=strSQL2 & "Impressions.AdvertiserID = Advertisers.AdvertiserID) INNER JOIN "
				strSQL2=strSQL2 & "Campaigns ON Impressions.CampaignID = Campaigns.CampaignID "
				strSQL2=strSQL2 & "Where (Campaigns.CampaignEndDate >=" & strStartDateSQL & ") AND (Campaigns.CampaignEndDate <=" & strEndDateSQL & ") " & strExtra
				strSQL2=strSQL2 & "GROUP BY Campaigns.CampaignID,Campaigns.CampaignName, Advertisers.CompanyName, "
				strSQL2=strSQL2 & "Campaigns.CampaignStartDate, Campaigns.CampaignEndDate, "
				strSQL2=strSQL2 & "Campaigns.CampaignType, Campaigns.CampaignQuantitySold "
				strSQL2=strSQL2 & "Order By Campaigns.CampaignEndDate Asc"
				strSQL=strSQL2
		Case "Cross Site Summary By Zone"
			If blnAllCampaigns=True Then
				strSQL="SELECT BanManProWebSites.SiteName, SUM(Impressions.ImpressionCount) "
				strSQL=strSQL & " AS SumOfImpressionCount, SUM(ClickCounts.Clicks) "
				strSQL=strSQL & " AS SumOfClicks, Zones.ZoneDescription "
				strSQL=strSQL & "FROM BanManProWebSites INNER JOIN "
				strSQL=strSQL & "    ((Impressions INNER JOIN "
				strSQL=strSQL & "    ClickCounts ON Impressions.ID = ClickCounts.ClickID) "
				strSQL=strSQL & "    INNER JOIN "
				strSQL=strSQL & "    Zones ON Impressions.ZoneID = Zones.ZoneID) ON "
				strSQL=strSQL & "    BanManProWebSites.SiteID = Impressions.UserID "
				strSQL=strSQL & "   WHERE ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) "
				strSQL=strSQL & "GROUP BY BanManProWebSites.SiteName, "
				strSQL=strSQL & "    Zones.ZoneDescription "
				strSQL=strSQL & "ORDER BY BanManProWebSites.SiteName, Zones.ZoneDescription "
			Else
				strSQL="SELECT BanManProWebSites.SiteName, SUM(Impressions.ImpressionCount) "
				strSQL=strSQL & " AS SumOfImpressionCount, SUM(ClickCounts.Clicks) "
				strSQL=strSQL & " AS SumOfClicks, Zones.ZoneDescription "
				strSQL=strSQL & "FROM BanManProWebSites INNER JOIN "
				strSQL=strSQL & "    ((Impressions INNER JOIN "
				strSQL=strSQL & "    ClickCounts ON Impressions.ID = ClickCounts.ClickID) "
				strSQL=strSQL & "    INNER JOIN "
				strSQL=strSQL & "    Zones ON Impressions.ZoneID = Zones.ZoneID) ON "
				strSQL=strSQL & "    BanManProWebSites.SiteID = Impressions.UserID "
				strSQL=strSQL & "   WHERE (Impressions.CampaignID=" & strCampaignID & ")  AND ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) "
				strSQL=strSQL & "GROUP BY BanManProWebSites.SiteName, "
				strSQL=strSQL & "    Zones.ZoneDescription "
				strSQL=strSQL & "ORDER BY BanManProWebSites.SiteName, Zones.ZoneDescription "
			End If
		Case "Cross Site Summary By Campaign"
			If blnAllCampaigns=True Then
				strSQL="SELECT BanManProWebSites.SiteName, "
				strSQL=strSQL & "     SUM(Impressions.ImpressionCount) "
				strSQL=strSQL & "     AS SumOfImpressionCount, SUM(ClickCounts.Clicks) "
				strSQL=strSQL & "     AS SumOfClicks, Campaigns.CampaignName "
				strSQL=strSQL & " FROM (BanManProWebSites INNER JOIN "
				strSQL=strSQL & "     (Impressions INNER JOIN "
				strSQL=strSQL & "     ClickCounts ON Impressions.ID = ClickCounts.ClickID) ON "
				strSQL=strSQL & "     BanManProWebSites.SiteID = Impressions.UserID) INNER JOIN "
				strSQL=strSQL & "     Campaigns ON "
				strSQL=strSQL & "     Impressions.CampaignID = Campaigns.CampaignID "
				strSQL=strSQL & "   WHERE ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) "
				strSQL=strSQL & " GROUP BY BanManProWebSites.SiteName, "
				strSQL=strSQL & "     Campaigns.CampaignName "
				strSQL=strSQL & " ORDER BY BanManProWebSites.SiteName, "
				strSQL=strSQL & "     Campaigns.CampaignName "
			Else
				strSQL="SELECT BanManProWebSites.SiteName, "
				strSQL=strSQL & "     SUM(Impressions.ImpressionCount) "
				strSQL=strSQL & "     AS SumOfImpressionCount, SUM(ClickCounts.Clicks) "
				strSQL=strSQL & "     AS SumOfClicks, Campaigns.CampaignName "
				strSQL=strSQL & " FROM (BanManProWebSites INNER JOIN "
				strSQL=strSQL & "     (Impressions INNER JOIN "
				strSQL=strSQL & "     ClickCounts ON Impressions.ID = ClickCounts.ClickID) ON "
				strSQL=strSQL & "     BanManProWebSites.SiteID = Impressions.UserID) INNER JOIN "
				strSQL=strSQL & "     Campaigns ON "
				strSQL=strSQL & "     Impressions.CampaignID = Campaigns.CampaignID "
				strSQL=strSQL & "   WHERE (Impressions.CampaignID=" & strCampaignID & ")  AND ((Impressions.ImpressionDay >= " & strStartDateSQL & ") AND   (Impressions.ImpressionDay <= " & strEndDateSQL & ")) "
				strSQL=strSQL & " GROUP BY BanManProWebSites.SiteName, "
				strSQL=strSQL & "     Campaigns.CampaignName "
				strSQL=strSQL & " ORDER BY BanManProWebSites.SiteName, "
				strSQL=strSQL & "     Campaigns.CampaignName "
			End If
	End Select
On Error Resume Next
	'Now grab data from database
	Set rsReports=connBanManPro.Execute(strSQL)

If Request("ReportFormat")<>"EXCEL" Then
%>

<html>
<head>
<title></title>
</head>
<body>
<p align="center"><img src="images/bannermanagerbanner.gif" WIDTH="544" HEIGHT="89"></p>
<% End If %>
<% '**********Now Print Report ************************************************************************
sngImpressions=0
sngClicks=0
Select Case strReportType
	Case "Summary By Day" %>
<div align="center"><center>

<table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000">
  <tr>
    <td bgcolor="#7A74FA" colspan="5">
      <p align="center"><b><font face="Arial" size="3">Summary By Day
      Advertising Report</font></b></td>
  </tr>
  <tr>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Campaign Name</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Date</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Clicks</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Impressions</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Click Rate</font></strong></td>
  </tr>
<% Do While NOT rsReports.EOF 
sngImpressions=sngImpressions+rsReports("SumOfImpressionCount")
sngClicks=sngClicks+rsReports("SumOfClicks") 
If rsReports("SumOfImpressionCount")<>"0" Then
	sngPercent=FormatPercent(rsReports("SumOfClicks")/rsReports("SumOfImpressionCount")) 
Else
	sngPercent="0.00%"
End If
%>
  <tr>
    <td><font face="Arial" size="3"><%If Session("AdvertiserID")<=0 And Request("ReportFormat")<>"EXCEL" Then%><a href="campaigns.asp?Task=Edit&CampaignID=<%=rsReports("CampaignID")%>"><%=rsReports("CampaignName")%></a><%Else%><%=rsReports("CampaignName")%><%End If%></font></td>
    <td align="center"><font face="Arial" size="3"><%=FormatDateTime(rsReports("ImpressionDay"),vbShortDate)%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("SumOfClicks")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("SumOfImpressionCount")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=sngPercent%></font></td>
  </tr>
<% rsReports.MoveNext
Loop 
If sngImpressions=0 Then
	varPercent="0 %"
Else
	varPercent=FormatPercent(sngClicks/sngImpressions)
End If
%>
</table>
<p><font face="Arial" size="3">Total Clicks: <strong><%=sngClicks%></strong><br>
Total Impressions: <strong><%=sngImpressions%></strong><br>
Overall Click Rate: </font> <strong><%=varPercent%></strong></p>
</center></div>

<% Case "Summary By Banner" '***************************************************************************%>

<div align="center"><center>

<table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000">
  <tr>
    <td bgcolor="#7A74FA" colspan="5">
      <p align="center"><b><font face="Arial" size="3">Summary By Banner
      Advertising Report</font></b></td>
  </tr>
  <tr>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Campaign Name</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Banner</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Clicks</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Impressions</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Click Rate</font></strong></td>
  </tr>
<% Do While NOT rsReports.EOF 
If rsReports("SumOfImpressionCount")<>"0" Then
	sngPercent=FormatPercent((rsReports("SumOfClicks")/rsReports("SumOfImpressionCount")))  
Else
	sngPercent="0.00%"
End If
sngImpressions=sngImpressions+rsReports("SumOfImpressionCount")
sngClicks=sngClicks+rsReports("SumOfClicks") 
%>
  <tr>
    <td><font face="Arial" size="3"><%If Session("AdvertiserID")<=0 And Request("ReportFormat")<>"EXCEL" Then%><a href="campaigns.asp?Task=Edit&CampaignID=<%=rsReports("CampaignID")%>"><%=rsReports("CampaignName")%></a><%Else%><%=rsReports("CampaignName")%><%End If%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("AdDescription")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("SumOfClicks")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("SumOfImpressionCount")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=sngPercent%></font></td>
  </tr>
<% rsReports.MoveNext
Loop 
If sngImpressions=0 Then
	varPercent="0 %"
Else
	varPercent=FormatPercent(sngClicks/sngImpressions)
End If %>
</table>
<p><font face="Arial" size="3">Total Clicks: <strong><%=sngClicks%></strong><br>
Total Impressions: <strong><%=sngImpressions%></strong><br>
Overall Click Rate: </font> <strong><%=varPercent%></strong></p>
</center></div>

<% Case "Summary By Banner By Day" '***************************************************************************%>

<div align="center"><center>

<table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000">
  <tr>
    <td bgcolor="#7A74FA" colspan="6">
      <p align="center"><b><font face="Arial" size="3">Summary By Banner By Day
      Advertising Report</font></b></td>
  </tr>
  <tr>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Campaign Name</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Banner</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Date</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Clicks</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Impressions</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Click Rate</font></strong></td>
  </tr>
<% Do While NOT rsReports.EOF 
If rsReports("SumOfImpressionCount")<>"0" Then
	sngPercent=FormatPercent((rsReports("SumOfClicks")/rsReports("SumOfImpressionCount")))  
Else
	sngPercent="0.00%"
End If
sngImpressions=sngImpressions+rsReports("SumOfImpressionCount")
sngClicks=sngClicks+rsReports("SumOfClicks") 
%>
  <tr>
    <td><font face="Arial" size="3"><%If Session("AdvertiserID")<=0 And Request("ReportFormat")<>"EXCEL" Then%><a href="campaigns.asp?Task=Edit&CampaignID=<%=rsReports("CampaignID")%>"><%=rsReports("CampaignName")%></a><%Else%><%=rsReports("CampaignName")%><%End If%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("AdDescription")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=FormatDateTime(rsReports("ImpressionDay"),vbShortDate)%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("SumOfClicks")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("SumOfImpressionCount")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=sngPercent%></font></td>
  </tr>
<% rsReports.MoveNext
Loop 
If sngImpressions=0 Then
	varPercent="0 %"
Else
	varPercent=FormatPercent(sngClicks/sngImpressions)
End If%>
</table>
<p><font face="Arial" size="3">Total Clicks: <strong><%=sngClicks%></strong><br>
Total Impressions: <strong><%=sngImpressions%></strong><br>
Overall Click Rate: <strong><%=varPercent%></strong></font></p>
</center></div>

<% Case "Summary By Zone" '***************************************************************************%>

<div align="center"><center>

<table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000">
  <tr>
    <td bgcolor="#7A74FA" colspan="5">
      <p align="center"><b><font face="Arial" size="3">Summary By Zone
      Advertising Report</font></b></td>
  </tr>
  <tr>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Zone</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Campaign Name</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Clicks</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Impressions</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Click Rate</font></strong></td>
  </tr>
<% Do While NOT rsReports.EOF 
If rsReports("SumOfImpressionCount")<>"0" Then
	sngPercent=FormatPercent((rsReports("SumOfClicks")/rsReports("SumOfImpressionCount")))  
Else
	sngPercent="0.00%"
End If
sngImpressions=sngImpressions+rsReports("SumOfImpressionCount")
sngClicks=sngClicks+rsReports("SumOfClicks") 
%>
  <tr>
    <td><font face="Arial" size="3"><%=rsReports("ZoneDescription")%></font></td>
    <td align="center"><font face="Arial" size="3"><%If Session("AdvertiserID")<=0 And Request("ReportFormat")<>"EXCEL" Then%><a href="campaigns.asp?Task=Edit&CampaignID=<%=rsReports("CampaignID")%>"><%=rsReports("CampaignName")%></a><%Else%><%=rsReports("CampaignName")%><%End If%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("SumOfClicks")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("SumOfImpressionCount")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=sngPercent%></font></td>
  </tr>
<% rsReports.MoveNext
Loop 
If sngImpressions=0 Then
	varPercent="0 %"
Else
	varPercent=FormatPercent(sngClicks/sngImpressions)
End If %>
</table>
<p><font face="Arial" size="3">Total Clicks: <strong><%=sngClicks%></strong><br>
Total Impressions: <strong><%=sngImpressions%></strong><br>
Overall Click Rate: <strong><%=varPercent%></strong></font></p>
</center></div>

<% Case "Summary By Zone By Day" '***************************************************************************%>

<div align="center"><center>

<table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000">
  <tr>
    <td bgcolor="#7A74FA" colspan="6">
      <p align="center"><b><font face="Arial" size="3">Summary By Zone By Day
      Advertising Report</font></b></td>
  </tr>
  <tr>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Zone</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Campaign Name</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Date</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Clicks</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Impressions</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Click Rate</font></strong></td>
  </tr>
<% Do While NOT rsReports.EOF 
If rsReports("SumOfImpressionCount")<>"0" Then
	sngPercent=FormatPercent((rsReports("SumOfClicks")/rsReports("SumOfImpressionCount")))  
Else
	sngPercent="0.00%"
End If
sngImpressions=sngImpressions+rsReports("SumOfImpressionCount")
sngClicks=sngClicks+rsReports("SumOfClicks") 
%>
  <tr>
    <td><font face="Arial" size="3"><%=rsReports("ZoneDescription")%></font></td>
    <td align="center"><font face="Arial" size="3"><%If Session("AdvertiserID")<=0 And Request("ReportFormat")<>"EXCEL" Then%><a href="campaigns.asp?Task=Edit&CampaignID=<%=rsReports("CampaignID")%>"><%=rsReports("CampaignName")%></a><%Else%><%=rsReports("CampaignName")%><%End If%></font></td>
    <td align="center"><font face="Arial" size="3"><%=FormatDateTime(rsReports("ImpressionDay"),vbShortDate)%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("SumOfClicks")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("SumOfImpressionCount")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=sngPercent%></font></td>
  </tr>
<% rsReports.MoveNext
Loop 
If sngImpressions=0 Then
	varPercent="0 %"
Else
	varPercent=FormatPercent(sngClicks/sngImpressions)
End If%>
</table>
<p><font face="Arial" size="3">Total Clicks: <strong><%=sngClicks%></strong><br>
Total Impressions: <strong><%=sngImpressions%></strong><br>
Overall Click Rate: <strong><%=varPercent%></strong></font></p>
</center></div>

<% Case "Click Detail" '***************************************************************************%>

<div align="center"><center>

<table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000">
  <tr>
    <td bgcolor="#7A74FA" colspan="4">
      <p align="center"><b><font face="Arial" size="3">Click Detail Advertising
      Report</font></b></td>
  </tr>
  <tr>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Campaign Name</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Date/Time</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">Browser</font></strong></td>
    <td bgcolor="#7A74FA"><strong><font face="Arial" size="3">IP Address</font></strong></td>
  </tr>
<% Do While NOT rsReports.EOF  %>
  <tr>
    <td><font face="Arial" size="3"><%=rsReports("CampaignName")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("ClickDateTime")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("ClickBrowser")%></font></td>
    <td align="center"><font face="Arial" size="3"><%=rsReports("ClickIP")%></font></td>
  </tr>
<% rsReports.MoveNext
Loop %>
</table>
</center></div>
<% Case "Executive"
'********************************************************************************************************
'****** Executive Report ********************************************************************************
'********************************************************************************************************
%>
<!--#Include File="Executive.asp"-->
<% Case "Billing" %>
<%
'********************************************************************************************************
'****** Billing Report **********************************************************************************
'********************************************************************************************************
%>
<div align="center">
  <center>
  <table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000">
    <tr>
      <td align="center" bgcolor="#7A74FA" colspan="8"><b><font face="Arial" size="3">Billing
        Report</font></b></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#7A74FA"><font face="Arial" size="2"><b>Campaign Name</b></font></td>
      <td align="center" bgcolor="#7A74FA"><font face="Arial" size="2"><b>Period Starting</b></font></td>
      <td align="center" bgcolor="#7A74FA"><font face="Arial" size="2"><b>Period Ending</b></font></td>
      <td align="center" bgcolor="#7A74FA"><font face="Arial" size="2"><b>Clicks</b></font></td>
      <td align="center" bgcolor="#7A74FA"><font face="Arial" size="2"><b>Impression</b></font></td>
      <td align="center" bgcolor="#7A74FA"><font face="Arial" size="2"><b>Ad Type</b></font></td>
      <td align="center" bgcolor="#7A74FA"><font face="Arial" size="2"><b>Cost</b></font></td>
      <td align="center" bgcolor="#7A74FA"><font face="Arial" size="2"><b>Total for Period</b></font></td>
    </tr>
<% Do While NOT rsReports.EOF  
If rsReports("CampaignType")="Flat Rate" Then
	If IsNull(rsReports("CampaignCost"))=False And Trim(rsReports("CampaignCost"))<> "" Then
		strTotal=FormatCurrency(rsReports("CampaignCost"))
	Else
		strTotal=FormatCurrency("0")
	End If
ElseIf rsReports("CampaignType")="CPM" Then
	If IsNull(rsReports("CampaignCost"))=False And Trim(rsReports("CampaignCost"))<> "" Then
		strTotal=FormatCurrency(rsReports("CampaignCost")*(rsReports("SumOfImpressionCount")/1000))
	Else
		strTotal=FormatCurrency("0")
	End If
ElseIf rsReports("CampaignType")="Per Click" Then
	If IsNull(rsReports("CampaignCost"))=False  And Trim(rsReports("CampaignCost"))<> "" Then
		strTotal=FormatCurrency(rsReports("CampaignCost")*(rsReports("SumOfClicks")))
	Else
		strTotal=FormatCurrency("0")
	End If
End If
%>
    <tr>
      <td align="center"><font face="Arial" size="2"><%=rsReports("CampaignName")%></font></td>
      <td align="center"><font face="Arial" size="2"><%=strStartDate%></font></td>
      <td align="center"><font face="Arial" size="2"><%=strEndDate%></font></td>
      <td align="center"><font face="Arial" size="2"><%=rsReports("SumOfClicks")%></font></td>
      <td align="center"><font face="Arial" size="2"><%=rsReports("SumOfImpressionCount")%></font></td>
      <td align="center"><font face="Arial" size="2"><%=rsReports("CampaignType")%></font></td>
      <td align="center"><font face="Arial" size="2"><%=rsReports("CampaignCost")%></font></td>
      <td align="center"><font face="Arial" size="2"><%=strTotal%></font></td>
    </tr>
<% rsReports.MoveNext
Loop %>
  </table>
  </center>
</div>
<% Case "Expiration" %>
<p align="center">&nbsp;</p>
<div align="center">
  <center>
  <table border="1" cellpadding="2" cellspacing="0" width="526">
    <tr>
      <td width="466" align="center" colspan="8"><font face="Arial" size="3"><b>Campaign
        Expiration Summary</b><br>
        </font><font face="Arial" size="2">(All campaigns with expiration dates
        during the selected dates.&nbsp; Note that campaigns may expire sooner
        if based on impressions or clicks.)</font></td>
    </tr>
    <tr>
      <td width="58" align="center"><font size="2" face="Arial"><b>Campaign</b></font></td>
      <td width="58" align="center"><font size="2" face="Arial"><b>Advertiser</b></font></td>
      <td width="58" align="center"><font size="2" face="Arial"><b>Start Date</b></font></td>
      <td width="58" align="center"><font size="2" face="Arial"><b>End Date</b></font></td>
      <td width="58" align="center"><font size="2" face="Arial"><b>Quantity<br>
        Sold</b></font></td>
      <td width="58" align="center"><font face="Arial" size="2"><b>Impressions</b></font></td>
      <td width="59" align="center"><font face="Arial" size="2"><b>Clicks</b></font></td>
      <td width="59" align="center"><font face="Arial" size="2"><b>Click Rate</b></font></td>
    </tr>
    <% Do While Not rsReports.EOF %>
    <tr>
      <td width="58" align="center"><font face="Arial" size="2"><%If Session("AdvertiserID")<=0 And Request("ReportFormat")<>"EXCEL" Then%><a href="campaigns.asp?Task=Edit&CampaignID=<%=rsReports("CampaignID")%>"><%=rsReports("CampaignName")%></a><%Else%><%=rsReports("CampaignName")%><%End If%>
        </font></td>
      <td width="58" align="center"><font face="Arial" size="2"><%=rsReports("CompanyName")%>
        </font></td>
      <td width="58" align="center"><font face="Arial" size="2"><%=rsReports("CampaignStartDate")%>
        </font></td>
      <td width="58" align="center"><font face="Arial" size="2"><%=FormatDateTime(rsReports("CampaignEndDate"),vbShortDate)%>
        </font></td>
      <td width="58" align="center"><font face="Arial" size="2"><%=rsReports("CampaignQuantitySold")%>
        </font></td>
      <td width="58" align="center"><font face="Arial" size="2"><%=rsReports("SumOfImpressionCount")%>
        </font></td>
      <td width="59" align="center"><font face="Arial" size="2"><%=rsReports("SumOfClicks")%>
        </font></td>
      <td width="59" align="center"><font face="Arial" size="2"><%=FormatPercent(rsReports("SumOfClicks")/rsReports("SumOfImpressionCount"))%>
        </font></td>
    </tr>
    <% rsReports.MoveNext
Loop %>
  </table>
  </center>
</div>

<% Case "Cross Site Summary By Zone" %>
<p align="center">&nbsp;</p>
<div align="center">
  <center>
  <table border="1" cellpadding="2" cellspacing="0" width="526">
    <tr>
      <td align="center" colspan="5" bgcolor="#7A74FA"><font face="Arial" size="3"><b>All
        Sites Summary By Zone<br>
        <%=strStartDate%>-<%=strEndDate%></b></font></td>
    </tr>
    <tr>
      <td align="left" bgcolor="#7A74FA"><font size="2" face="Arial"><b>Site</b></font></td>
      <td align="left" bgcolor="#7A74FA"><font size="2" face="Arial"><b>Zone</b></font></td>
      <td align="center" bgcolor="#7A74FA"><font size="2" face="Arial"><b>Impressions</b></font></td>
      <td align="center" bgcolor="#7A74FA"><font size="2" face="Arial"><b>Clicks</b></font></td>
      <td align="center" bgcolor="#7A74FA"><font size="2" face="Arial"><b>Click
        Rate</b></font></td>
    </tr>
    <% Do While Not rsReports.EOF %>
    <tr>
      <td align="left"><font face="Arial" size="2"><%=rsReports("SiteName")%>
        </font></td>
      <td align="left"><font face="Arial" size="2"><%=rsReports("ZoneDescription")%>
        </font></td>
      <td align="center"><font face="Arial" size="2"><%=rsReports("SumOfImpressionCount")%>
        </font></td>
      <td align="center"><font face="Arial" size="2"><%=rsReports("SumOfClicks")%>
        </font></td>
<% If rsReports("SumOfImpressionCount")>0 Then
	strTemp=FormatPercent(rsReports("SumOfClicks")/rsReports("SumOfImpressionCount"))
Else
	strTemp=FormatPercent(0)
End If
%>
      <td align="center"><font face="Arial" size="2"><%=strTemp%>
        </font></td>
    </tr>
    <% rsReports.MoveNext
Loop %>
  </table>
  </center>
</div>

<% Case "Cross Site Summary By Campaign" %>
<p align="center">&nbsp;</p>
<div align="center">
  <center>
  <table border="1" cellpadding="2" cellspacing="0" width="526">
    <tr>
      <td align="center" colspan="5" bgcolor="#7A74FA"><font face="Arial" size="3"><b>All
        Sites Summary By Campaign<br>
        <%=strStartDate%>-<%=strEndDate%></b></font></td>
    </tr>
    <tr>
      <td align="left" bgcolor="#7A74FA"><font size="2" face="Arial"><b>Site</b></font></td>
      <td align="left" bgcolor="#7A74FA"><font size="2" face="Arial"><b>Campaign</b></font></td>
      <td align="center" bgcolor="#7A74FA"><font size="2" face="Arial"><b>Impressions</b></font></td>
      <td align="center" bgcolor="#7A74FA"><font size="2" face="Arial"><b>Clicks</b></font></td>
      <td align="center" bgcolor="#7A74FA"><font size="2" face="Arial"><b>Click
        Rate</b></font></td>
    </tr>
    <% Do While Not rsReports.EOF %>
    <tr>
      <td align="left"><font face="Arial" size="2"><%=rsReports("SiteName")%>
        </font></td>
      <td align="left"><font face="Arial" size="2"><%=rsReports("CampaignName")%>
        </font></td>
      <td align="center"><font face="Arial" size="2"><%=rsReports("SumOfImpressionCount")%>
        </font></td>
      <td align="center"><font face="Arial" size="2"><%=rsReports("SumOfClicks")%>
        </font></td>
<% If rsReports("SumOfImpressionCount")>0 Then
	strTemp=FormatPercent(rsReports("SumOfClicks")/rsReports("SumOfImpressionCount"))
Else
	strTemp=FormatPercent(0)
End If
%>
      <td align="center"><font face="Arial" size="2"><%=strTemp%>
        </font></td>
    </tr>
    <% rsReports.MoveNext
Loop %>
  </table>
  </center>
</div>

<% End Select %>

<% If Request("ReportFormat")<>"EXCEL" Then %>


<div align="left">
<table BORDER="1" BGCOLOR="#FFFFFF" CELLSPACING="0" align="left" width="640">
<font FACE="Arial" COLOR="#000000"><THEAD>
<tr>
</font>
<% End If %>

<% Set rsReports=Nothing %>
</table>
</div>

</body>
