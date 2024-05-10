<%
	If Session("AdvertiserID")=0 Then
		strSQL="SELECT Impressions.CampaignID, Impressions.AdvertiserID, Count(Impressions.ImpressionDay) "
		strSQL=strSQL & " AS CountOfImpressionDay, Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, "
		strSQL=strSQL & " Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignName, Campaigns.CampaignStartDate, "
		strSQL=strSQL & " Campaigns.CampaignEndDate "
		strSQL=strSQL & " FROM (Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) "
		strSQL=strSQL & " INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID "
		strSQL=strSQL & " Where Impressions.UserID=" & CLng(Session("BanManProSiteID")) & " AND Campaigns.CampaignEndDate >= getdate() "
		strSQL=strSQL & " GROUP BY Impressions.CampaignID, Impressions.AdvertiserID, Campaigns.CampaignName, "
		strSQL=strSQL & " Campaigns.CampaignStartDate, Campaigns.CampaignEndDate "		
		Set rs=connBanManPro.Execute(strSQL)   

		strSQL="SELECT Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Impressions.ImpressionDay "
		strSQL=strSQL & " FROM Impressions "
		strSQL=strSQL & " Where Impressions.ImpressionDay  > DATEADD(day, -7, getdate())"
		strSQL=strSQL & " AND Impressions.UserID=" & CLng(Session("BanManProSiteID"))
		strSQL=strSQL & "GROUP BY Impressions.ImpressionDay "
		strSQL=strSQL & "ORDER BY Impressions.ImpressionDay ASC"
  		Set rsReport1=connBanManPro.Execute(strSQL)   
   
	Else
		strSQL="SELECT Impressions.CampaignID, Impressions.AdvertiserID, Count(Impressions.ImpressionDay) "
		strSQL=strSQL & " AS CountOfImpressionDay, Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, "
		strSQL=strSQL & " Sum(ClickCounts.Clicks) AS SumOfClicks, Campaigns.CampaignName, Campaigns.CampaignStartDate, "
		strSQL=strSQL & " Campaigns.CampaignEndDate "
		strSQL=strSQL & " FROM (Impressions INNER JOIN Campaigns ON Impressions.CampaignID = Campaigns.CampaignID) "
		strSQL=strSQL & " INNER JOIN ClickCounts ON Impressions.ID = ClickCounts.ClickID "
		strSQL=strSQL & " WHERE (((Impressions.AdvertiserID)=" & Session("AdvertiserID") & "))"
		strSQL=strSQL & " GROUP BY Impressions.CampaignID, Impressions.AdvertiserID, Campaigns.CampaignName, "
		strSQL=strSQL & " Campaigns.CampaignStartDate, Campaigns.CampaignEndDate "		
		Set rs=connBanManPro.Execute(strSQL)   

		strSQL="SELECT Sum(Impressions.ImpressionCount) AS SumOfImpressionCount, Impressions.ImpressionDay "
		strSQL=strSQL & " FROM Impressions "
		strSQL=strSQL & " Where Impressions.ImpressionDay  > DATEADD(day, -7, getdate())"
		strSQL=strSQL & " AND Impressions.UserID=" & CLng(Session("BanManProSiteID"))
		strSQL=strSQL & " AND (Impressions.AdvertiserID)=" & Session("AdvertiserID")
		strSQL=strSQL & "GROUP BY Impressions.ImpressionDay "
		strSQL=strSQL & "ORDER BY Impressions.ImpressionDay ASC"
  		Set rsReport1=connBanManPro.Execute(strSQL)  


	End If
%>