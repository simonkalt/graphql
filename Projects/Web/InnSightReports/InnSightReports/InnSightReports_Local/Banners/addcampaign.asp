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
	
	strMessage=""
	blnFoundError=False

	Select Case strTask
		Case "AddNew" 
			'obtain list of advertisers
			strSQL="SELECT * FROM Advertisers Where (UserID=" & CLng(Session("BanManProSiteID")) & " Or UserID=0) ORDER BY Advertisers.[CompanyName] ASC"
			Set rsa=connBanManPro.Execute(strSQL)
			If Request.Form("AdvertiserID") <> "" Then 
				If Request.Form("CampaignType")="StaticText" Then%>
					<!--#include file="addanewcampaigns.asp"-->
			<%	Else 
					If Application("SlotOption")=True Then%>
						<!--#include file="addanewcampaignslot.asp"-->
					<% Else %>
						<!--#include file="addanewcampaign.asp"-->
					<% End If %>
				<%End IF
			Else %>
			<!--#include file="selectadvertiser.asp"-->
			<%
			End If
			Set rsa=Nothing
		Case "Edit"
			'edit record
			If Trim(strCampaignID) <> "" Then
				strSQL2="SELECT Campaigns.*, Advertisers.CompanyName FROM Advertisers INNER JOIN Campaigns ON Advertisers.AdvertiserID = Campaigns.AdvertiserID WHERE (((Campaigns.CampaignID)=" & strCampaignID & "))"
				Set rsc=connBanManPro.Execute(strSQL2)
				'obtain list of advertisers
				'strSQL="SELECT * FROM Advertisers Where (UserID=" & CLng(Session("BanManProSiteID")) & " Or UserID=0) ORDER BY Advertisers.[CompanyName] ASC"
				'Set rsa=connBanManPro.Execute(strSQL)
				If Not rsc.EOF Then
					If rsc("CampaignDistribution")="Text" Then
 						%>
						<!--#include file="addanewCampaigns.asp"-->
						<%
					Else
						If Application("SlotOption") Then%>
							<!--#include file="addanewcampaignslot.asp"-->
						<% Else %>
							<!--#include file="addanewcampaign.asp"-->
						<% End If %>
						<%
					End If
				End If
				Set rsc=Nothing
				Set rsa=Nothing
			End If	
		Case "Insert"
			'error checks *************************
			'Check if campaign name is already used
			Set rsc=connBanManPro.Execute("Select CampaignName From Campaigns Where CampaignName='" & FixBlank(Request.Form("CampaignName")) & "' AND (UserID=" & CLng(Session("BanManProSiteID")) & " Or UserID=0)")
			blnFoundSameName=False
			If rsc.EOF=False Then
				blnFoundSame=True
				Response.Write "<p align=" & Chr(34) & "center" & Chr(34) & ">Invalid Campaign Name. "
				Response.Write "<p align=" & Chr(34) & "center" & Chr(34) & ">Name already in use for another campaign."
				blnFoundError=True
			End If
			Set rsc=Nothing

			If Trim(Request.Form("CampaignName"))="" Then
				Response.Write "<p align=" & Chr(34) & "center" & Chr(34) & ">Invalid Campaign Name (This is a required field)"
				blnFoundError=True
			End If
			If Trim(Request.Form("CampaignQuantitySold"))="" And Request.Form("CampaignSiteDefault") <> "-1" And Request.Form("CampaignType")<> "Flat Rate" Then
				Response.Write "<p align=" & Chr(34) & "center" & Chr(34) & ">Invalid Quantity (This is a required field)"
				blnFoundError=True
			End If
			If Request.Form("CampaignDistribution")<>"Text" Then			
				If IsDate(Request.Form("StartMonth") & "/" & Request.Form("StartDay") & "/" & Request.Form("StartYear")) OR IsDate(Request.Form("StartDay") & "/" & Request.Form("StartMonth") & "/" & Request.Form("StartYear")) Then
					'strStartDate=(Request.Form("StartMonth") & "/" & Request.Form("StartDay") & "/" & Request.Form("StartYear") & " 00:00:00")
					strStartDate="CONVERT(DATETIME,'" & Request.Form("StartMonth") & "/" & Request.Form("StartDay") & "/" & Request.Form("StartYear") & " 00:00:00',101)"				
				Else
					Response.Write "<p align=center>Invalid Start Date"
					blnFoundError=True
				End If
				If IsDate(Request.Form("EndMonth") & "/" & Request.Form("EndDay") & "/" & Request.Form("EndYear")) OR IsDate(Request.Form("EndDay") & "/" & Request.Form("EndMonth") & "/" & Request.Form("EndYear")) Then
	 				'strEndDate=(Request.Form("EndMonth") & "/" & Request.Form("EndDay")  & "/" &  Request.Form("EndYear") & " 23:59:59")
					strEndDate="CONVERT(DATETIME,'" & Request.Form("EndMonth") & "/" & Request.Form("EndDay")  & "/" &  Request.Form("EndYear") & " 23:59:59',101)"
				Else
					Response.Write "<p align=center>Invalid End Date"
					blnFoundError=True
				End If
			Else
				strStartDate="''"
				strEndDate="''"
			End If
			If Request("CampaignDistribution")="Even" And Request("CampaignType")<>"CPM" Then
				Response.Write "<p align=center>Error.  Evenly distributed campaigns can only be of type CPM"
				blnFoundError=True
			End If
			'Check Sum of weightings
			If Request.Form("CampaignDistribution")<>"Text" Then
				'now add banners to CampaignBanners
				intCnt=0
				sngSum=0
				Do Until intCnt=1000
					intCnt=intCnt+1
					strSelected="chkBannerSelected" & intCnt
					strWeighting="txtBannerWeighting" & intCnt
					If Trim(Request.Form(strSelected)) <> "" Then
						If IsNumeric(Request.Form(strWeighting)) Then
							sngBanWeight=CInt(Request.Form(strWeighting))
						Else
							sngBanWeight=0
						End If
						sngSum=sngSum+Csng(sngBanWeight)
					End If
				Loop
				If sngSum<=0 Then
					Response.Write "<p align=center>Error. You have not properly weighted the banners in this campaign."
					blnFoundError=True
				End If
			End If
			'end error checks *********************
		    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" And blnFoundError=False  Then


			strSQL2="INSERT INTO Campaigns ("
			strSQL2=strSQL2 & "AdvertiserID,"
			strSQL2=strSQL2 & "CampaignName,"
			strSQL2=strSQL2 & "CampaignStartDate,"
			strSQL2=strSQL2 & "CampaignEndDate,"
			strSQL2=strSQL2 & "CampaignType,"
			strSQL2=strSQL2 & "CampaignQuantitySold,"
			strSQL2=strSQL2 & "CampaignCost,"
			strSQL2=strSQL2 & "CampaignDailyStart,"
			strSQL2=strSQL2 & "CampaignDailyEnd,"
			strSQL2=strSQL2 & "CampaignSunday,"
			strSQL2=strSQL2 & "CampaignMonday,"
			strSQL2=strSQL2 & "CampaignTuesday,"
			strSQL2=strSQL2 & "CampaignWednesday,"
			strSQL2=strSQL2 & "CampaignThursday,"
			strSQL2=strSQL2 & "CampaignFriday,"
			strSQL2=strSQL2 & "CampaignSaturday,"
			strSQL2=strSQL2 & "CampaignSiteDefault,"
			strSQL2=strSQL2 & "CampaignDistribution,"
			strSQL2=strSQL2 & "CampaignImpressionsServed,UserID,CampaignKeywords,CampaignQuantityExpected) VALUES ("
			strSQL2=strSQL2 & FixBlank(Request.Form("AdvertiserID")) & ",'" 
			strSQL2=strSQL2 & FixBlank(Request.Form("CampaignName")) & "'," 
			strSQL2=strSQL2 & strStartDate & "," 
			strSQL2=strSQL2 & strEndDate & ",'" 
			strSQL2=strSQL2 & FixBlank(Request.Form("CampaignType")) & "'," 
			strSQL2=strSQL2 & Replace(FixZero(Request.Form("CampaignQuantitySold")),",","") & ",'" 
			strSQL2=strSQL2 & FixBlank(Request.Form("CampaignCost")) & "','" 
			strSQL2=strSQL2 & FixBlank(Request.Form("CampaignDailyStart")) & "','" 
			strSQL2=strSQL2 & FixBlank(Request.Form("CampaignDailyEnd")) & "',"
 			strSQL2=strSQL2 & SetTrueFalse(Request.Form("CampaignSunday")) & ","
			strSQL2=strSQL2 & SetTrueFalse(Request.Form("CampaignMonday")) & ","
			strSQL2=strSQL2 & SetTrueFalse(Request.Form("CampaignTuesday")) & ","
			strSQL2=strSQL2 & SetTrueFalse(Request.Form("CampaignWednesday")) & ","
			strSQL2=strSQL2 & SetTrueFalse(Request.Form("CampaignThursday")) & ","
			strSQL2=strSQL2 & SetTrueFalse(Request.Form("CampaignFriday")) & ","
			strSQL2=strSQL2 & SetTrueFalse(Request.Form("CampaignSaturday")) & ","
			strSQL2=strSQL2 & SetTrueFalse(Request.Form("CampaignSiteDefault")) & ",'"
			strSQL2=strSQL2 & FixBlank(Request.Form("CampaignDistribution")) & "'," 		
			strSQL2=strSQL2 & "1," 
			If Request.Form("RunOfNetwork")="ON" Then
				strSQL2=strSQL2 & Clng(0) 
			Else
				strSQL2=strSQL2 & CLng(Session("BanManProSiteID"))  
			End If		
			strSQL2=strSQL2 & ",'" & FixBlank(Request.Form("CampaignKeywords")) & "',0)" 
			connBanManPro.Execute strSQL2

			'get newly added campaign ID
			strSQL2="SELECT Campaigns.CampaignID FROM Campaigns WHERE Campaigns.CampaignName In ('" & FixBlank(Request.Form("CampaignName")) & "') And Campaigns.AdvertiserID=" & Clng(Request.Form("AdvertiserID"))
			Set rsCampaignID=connBanManPro.Execute(strSQL2)
			strCampaignID=rsCampaignID("CampaignID")
			Set rsCampaignID=Nothing

			'Create Record in CampaignClicks Table for this Campaign
			strSQL="INSERT INTO CampaignClicks (CampaignClickID,Clicks) VALUES (" & strCampaignID & ",0)"
			connBanManPro.Execute strSQL

			If Request.Form("CampaignDistribution")<>"Text" Then
				'now add banners to CampaignBanners
				intCnt=0
				sngSum=0
				Do Until intCnt=1000
					intCnt=intCnt+1
					strSelected="chkBannerSelected" & intCnt
					strWeighting="txtBannerWeighting" & intCnt
					If Trim(Request.Form(strSelected)) <> "" Then
						If IsNumeric(Request.Form(strWeighting)) Then
							sngBanWeight=CInt(Request.Form(strWeighting))
						Else
							sngBanWeight=0
						End If
						sngSum=sngSum+Csng(sngBanWeight)
						strSQL2="INSERT INTO CampaignBanners ("
						strSQL2=strSQL2 & "CampaignID,"
						strSQL2=strSQL2 & "BannerID,"
						strSQL2=strSQL2 & "CampaignBannerWeighting,UserID) VALUES ("
						strSQL2=strSQL2 & strCampaignID & ","
						strSQL2=strSQL2 & Request.Form(strSelected) & "," 
						strSQL2=strSQL2 & sngBanWeight & "," & CLng(Session("BanManProSiteID")) & ")" 
						connBanManPro.Execute strSQL2
					Else
						'Exit Do
					End If
				Loop
			Else
				strSQL2="INSERT INTO CampaignBanners ("
				strSQL2=strSQL2 & "CampaignID,"
				strSQL2=strSQL2 & "BannerID,"
				strSQL2=strSQL2 & "CampaignBannerWeighting,UserID) VALUES ("
				strSQL2=strSQL2 & strCampaignID & ","
				strSQL2=strSQL2 & Request.Form("Banners") & ",100," 
				strSQL2=strSQL2 & CLng(Session("BanManProSiteID")) & ")" 
				connBanManPro.Execute strSQL2
			End If

			%>
			<!--#Include File="banmanfunc.asp"-->
			<%
			If Application("SlotOption")<>True Then
				CalculateBanManProExpectedQuantity  		
			End If
		End If
		If  blnFoundError=False  Then
			If (sngSum >= 101) Then
				%>
				<p align="center"><font face="Arial" size="5">Successfully added new Campaign.</font></p>
				<p align="center"><font face="Arial" size="5">Alert** Banners within this campaign do not sum to 100.  Please fix this problem.</font></p>
				<%  	
			Else
				%>
				<p align="center"><font face="Arial" size="5">Successfully added new Campaign.</font></p>
				<%  	
			End If 
		End If
		Case "Update"
			'error checks *************************
			'Check if campaign name is already used
			Set rsc=connBanManPro.Execute("Select CampaignName From Campaigns Where CampaignName='" & FixBlank(Request.Form("CampaignName")) & "' AND CampaignID<>" & Clng(strCampaignID) & " AND (UserID=" & CLng(Session("BanManProSiteID")) & " Or UserID=0)")
			blnFoundSameName=False
			If rsc.EOF=False Then
				blnFoundSame=True
				Response.Write "<p align=" & Chr(34) & "center" & Chr(34) & ">Invalid Campaign Name. "
				Response.Write "<p align=" & Chr(34) & "center" & Chr(34) & ">Name already in use for another campaign."
				blnFoundError=True
			End If
			Set rsc=Nothing


			If Trim(Request.Form("CampaignName"))="" Then
				Response.Write "<p align=" & Chr(34) & "center" & Chr(34) & ">Invalid Campaign Name (This is a required field)"
				blnFoundError=True
			End If
			If Trim(Request.Form("CampaignQuantitySold"))="" And Request.Form("CampaignSiteDefault") <> "-1" Then
				Response.Write "<p align=" & Chr(34) & "center" & Chr(34) & ">Invalid Quantity (This is a required field)"
				blnFoundError=True
			End If			
			If Request.Form("CampaignDistribution")<>"Text" Then			
				If IsDate(Request.Form("StartMonth") & "/" & Request.Form("StartDay") & "/" & Request.Form("StartYear")) OR IsDate(Request.Form("StartDay") & "/" & Request.Form("StartMonth") & "/" & Request.Form("StartYear")) Then
					'strStartDate=(Request.Form("StartMonth") & "/" & Request.Form("StartDay") & "/" & Request.Form("StartYear") & " 00:00:00")
					strStartDate="CONVERT(DATETIME,'" & Request.Form("StartMonth") & "/" & Request.Form("StartDay") & "/" & Request.Form("StartYear") & " 00:00:00',101)"				
				Else
					Response.Write "<p align=center>Invalid Start Date"
					blnFoundError=True
				End If
				If IsDate(Request.Form("EndMonth") & "/" & Request.Form("EndDay") & "/" & Request.Form("EndYear")) OR IsDate(Request.Form("EndDay") & "/" & Request.Form("EndMonth") & "/" & Request.Form("EndYear")) Then
	 				'strEndDate=(Request.Form("EndMonth") & "/" & Request.Form("EndDay")  & "/" &  Request.Form("EndYear") & " 23:59:59")
					strEndDate="CONVERT(DATETIME,'" & Request.Form("EndMonth") & "/" & Request.Form("EndDay")  & "/" &  Request.Form("EndYear") & " 23:59:59',101)"
				Else
					Response.Write "<p align=center>Invalid End Date"
					blnFoundError=True
				End If
			Else
				strStartDate="''"
				strEndDate="''"
			End If
			If Request("CampaignDistribution")="Even" And Request("CampaignType")<>"CPM" Then
				Response.Write "<p align=center>Error.  Evenly distributed campaigns can only be of type CPM"
				blnFoundError=True
			End If
			'end error checks *********************
		    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" And blnFoundError=False Then

			'Determine if Campaign is currently an even campaign.
			strSQL="Select CampaignDistribution From Campaigns Where CampaignID=" & Clng(strCampaignID)
			Set rs=connBanManPro.Execute(strSQL)
			If rs("CampaignDistribution")<> Request.Form("CampaignDistribution") Then
				If Request.Form("CampaignDistribution")="Weighted" AND rs("CampaignDistribution")="Normal" Then
					'No need to delete
				Else
					'change from anything to keyword, must delete campaigns in zones
					'change to keywords, or weighted to even, must delete campaigns
					strSQL="Delete From ZoneCampaigns Where CampaignID=" & Clng(strCampaignID)
					connBanManPro.Execute strSQL
					%>
					<p align="center">Because you changed the campaign distribution, the<br>
					campaign has been deleted from all zones.  Please<br>
					update your zones to reflect this change.<br>
					<%
				End If
			End If

			strSQL2="UPDATE Campaigns SET "
			strSQL2=strSQL2 & "AdvertiserID=" &  FixBlank(Request.Form("AdvertiserID")) & ","  
			strSQL2=strSQL2 & "CampaignName='" &  FixBlank(Request.Form("CampaignName")) & "',"
			strSQL2=strSQL2 & "CampaignStartDate=" &  strStartDate & "," 
			strSQL2=strSQL2 & "CampaignEndDate=" &  strEndDate & "," 
			strSQL2=strSQL2 & "CampaignType='" &  FixBlank(Request.Form("CampaignType")) & "'," 
			strSQL2=strSQL2 & "CampaignQuantitySold=" &  Replace(FixZero(Request.Form("CampaignQuantitySold")),",","") & "," 
			strSQL2=strSQL2 & "CampaignCost='" &  FixBlank(Request.Form("CampaignCost")) & "'," 
			strSQL2=strSQL2 & "CampaignDailyStart='" &  TimeValue(FixBlank(Request.Form("CampaignDailyStart"))) & "'," 
			strSQL2=strSQL2 & "CampaignDailyEnd='" &  TimeValue(FixBlank(Request.Form("CampaignDailyEnd"))) & "'," 
			strSQL2=strSQL2 & "CampaignSunday=" &  SetTrueFalse(Request.Form("CampaignSunday")) & "," 
			strSQL2=strSQL2 & "CampaignMonday=" &  SetTrueFalse(Request.Form("CampaignMonday")) & "," 
			strSQL2=strSQL2 & "CampaignTuesday=" &  SetTrueFalse(Request.Form("CampaignTuesday")) & "," 
			strSQL2=strSQL2 & "CampaignWednesday=" &  SetTrueFalse(Request.Form("CampaignWednesday")) & "," 
			strSQL2=strSQL2 & "CampaignThursday=" &  SetTrueFalse(Request.Form("CampaignThursday")) & "," 	
			strSQL2=strSQL2 & "CampaignFriday=" &  SetTrueFalse(Request.Form("CampaignFriday")) & "," 
			strSQL2=strSQL2 & "CampaignSaturday=" &  SetTrueFalse(Request.Form("CampaignSaturday")) & "," 
			strSQL2=strSQL2 & "CampaignSiteDefault=" &  SetTrueFalse(Request.Form("CampaignSiteDefault")) & "," 
			strSQL2=strSQL2 & "CampaignKeywords='" & FixBlank(Request.Form("CampaignKeywords")) & "',"
			If Request.Form("RunOfNetwork")="ON" Then
				strSQL2=strSQL2 & "UserID=0,"
			Else
				strSQL2=strSQL2 & "UserID=" & CLng(Session("BanManProSiteID"))  & ","
			End If
			If strEndDate > Now Then
				strSQL2=strSQL2 & "CampaignNotificationSent=0," 
				strSQL2=strSQL2 & "CampaignWarningSent=0," 
			End If
			strSQL2=strSQL2 & "CampaignDistribution='" &  FixBlank(Request.Form("CampaignDistribution")) & "' " 
			strSQL2=strSQL2 & "WHERE Campaigns.[CampaignID]=" & strCampaignID 
			'Response.Write strSQL2
			connBanManPro.Execute strSQL2

			'Must Delete ZoneCampaigns if this campaign has changed to a default
			If Request.Form("CampaignSiteDefault")="-1" Then
				strSQL="Delete From ZoneCampaigns Where CampaignID=" & strCampaignID
				connBanManPro.Execute strSQL
			End if

			'If slot option, must update zones with this slot
			If Application("SlotOption")=True Then
				strSQL="Update ZoneCampaigns Set ZoneCampaignWeighting=" & FixZero(Request.Form("CampaignQuantitySold")) & " Where CampaignID=" & strCampaignID
				connBanManPro.Execute strSQL
			End If

			'now Update Banners
			intCnt=0
			'first delete all records for this campaign
			strSQL2="DELETE FROM CampaignBanners WHERE CampaignBanners.[CampaignID]=" & strCampaignID
			'Response.Write strSQL2
			connBanManPro.Execute strSQL2
			If Request.Form("CampaignDistribution")<>"Text" Then
				'now add banners to CampaignBanners
				intCnt=0
				sngSum=0
				Do Until intCnt=1000
					intCnt=intCnt+1
					strSelected="chkBannerSelected" & intCnt
					strWeighting="txtBannerWeighting" & intCnt
					If Trim(Request.Form(strSelected)) <> "" Then
						If IsNumeric(Request.Form(strWeighting)) Then
							sngBanWeight=CInt(Request.Form(strWeighting))
						Else
							sngBanWeight=0
						End If
						sngSum=sngSum+Csng(sngBanWeight)
						strSQL2="INSERT INTO CampaignBanners ("
						strSQL2=strSQL2 & "CampaignID,"
						strSQL2=strSQL2 & "BannerID,"
						strSQL2=strSQL2 & "CampaignBannerWeighting,UserID) VALUES ("
						strSQL2=strSQL2 & strCampaignID & ","
						strSQL2=strSQL2 & Request.Form(strSelected) & "," 
						strSQL2=strSQL2 & sngBanWeight & "," & CLng(Session("BanManProSiteID")) & ")" 
						connBanManPro.Execute strSQL2
					Else
						'Exit Do
					End If
				Loop
			Else
				strSQL2="INSERT INTO CampaignBanners ("
				strSQL2=strSQL2 & "CampaignID,"
				strSQL2=strSQL2 & "BannerID,"
				strSQL2=strSQL2 & "CampaignBannerWeighting,UserID) VALUES ("
				strSQL2=strSQL2 & strCampaignID & ","
				strSQL2=strSQL2 & Request.Form("Banners") & ",100," 
				strSQL2=strSQL2 & CLng(Session("BanManProSiteID")) & ")" 
				connBanManPro.Execute strSQL2
			End If   
			%>
			<!--#Include File="banmanfunc.asp"-->
			<%
			If Application("SlotOption")<>True Then
				CalculateBanManProExpectedQuantity  		
			End If
			Set rs=Nothing
		    End If
			If  blnFoundError=False  Then
			If (sngSum >= 101) Then
				%>
				<p align="center"><font face="Arial" size="5">Updated Campaign.</font></p>
				<p align="center"><font face="Arial" size="5">Alert** Banners within this campaign do not sum to 100.  Please fix this problem.</font></p>
				<%  	
			Else
				%>
				<p align="center"><font face="Arial" size="5">Successfully Updated Campaign.</font></p>
				<%  	
			End If
			End If  
		Case "Delete"
			'delete entry
			If Trim(strCampaignID) <> "" Then
			    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then
				strSQL2="DELETE FROM Campaigns WHERE Campaigns.[CampaignID]=" & strCampaignID
				connBanManPro.Execute strSQL2
				'delete zone campaigns
				strSQL2=" Delete From ZoneCampaigns Where CampaignID=" & Clng(strCampaignID) 
				connBanManPro.Execute strSQL2
			    End If
				%>
				<p align="center"><font face="Arial" size="5">Record Deleted.</font></p>
				<%  	
			Else 	%>
				<p align="center"><font face="Arial" size="5">Nothing to Delete.</font></p>
				<%
			End If		
		Case "ViewAll", ""
			If strTask="" Then
				'strSQL2="SELECT Top 10 Campaigns.*, Advertisers.CompanyName FROM Advertisers INNER JOIN Campaigns ON Advertisers.AdvertiserID = Campaigns.AdvertiserID Where (Campaigns.UserID=" & CLng(Session("BanManProSiteID")) & " Or Campaigns.UserID=0)"
				strSQL2="SELECT Top 10 validcampaigns_type.CampaignID, validcampaigns_type.AdvertiserID, Campaigns.UserID, Campaigns.CampaignQuantityExpected, "
    				strSQL2=strSQL2 & " Campaigns.CampaignKeywords, Campaigns.CampaignNotificationSent, Campaigns.CampaignWarningSent, Campaigns.CampaignSiteDefault, "
    				strSQL2=strSQL2 & " Campaigns.CampaignExpired, Campaigns.CampaignImpressionsServed, Campaigns.CampaignSaturday, Campaigns.CampaignFriday, "
    				strSQL2=strSQL2 & " Campaigns.CampaignThursday, Campaigns.CampaignWednesday, Campaigns.CampaignTuesday, Campaigns.CampaignMonday, Campaigns.CampaignSunday, "
    				strSQL2=strSQL2 & " Campaigns.CampaignDistribution, Campaigns.CampaignDailyEnd, Campaigns.CampaignDailyStart, Campaigns.CampaignCost, "
    				strSQL2=strSQL2 & " Campaigns.CampaignQuantitySold, Campaigns.CampaignType, Campaigns.CampaignEndDate, Campaigns.CampaignStartDate, Campaigns.CampaignName, "
    				strSQL2=strSQL2 & " Advertisers.CompanyNamE FROM validcampaigns_type INNER JOIN Advertisers ON validcampaigns_type.AdvertiserID = Advertisers.AdvertiserID INNER "
    				strSQL2=strSQL2 & " JOIN Campaigns ON validcampaigns_type.CampaignID = Campaigns.CampaignID "
				strSQL2=strSQL2 & " Where (Campaigns.UserID=" & CLng(Session("BanManProSiteID")) & " Or Campaigns.UserID=0)"			
			Else
				'strSQL2="SELECT Campaigns.*, Advertisers.CompanyName FROM Advertisers INNER JOIN Campaigns ON Advertisers.AdvertiserID = Campaigns.AdvertiserID Where (Campaigns.UserID=" & CLng(Session("BanManProSiteID")) & " Or Campaigns.UserID=0)"
				strSQL2="SELECT Top 10 validcampaigns_type.CampaignID, validcampaigns_type.AdvertiserID, Campaigns.UserID, Campaigns.CampaignQuantityExpected, "
    				strSQL2=strSQL2 & " Campaigns.CampaignKeywords, Campaigns.CampaignNotificationSent, Campaigns.CampaignWarningSent, Campaigns.CampaignSiteDefault, "
    				strSQL2=strSQL2 & " Campaigns.CampaignExpired, Campaigns.CampaignImpressionsServed, Campaigns.CampaignSaturday, Campaigns.CampaignFriday, "
    				strSQL2=strSQL2 & " Campaigns.CampaignThursday, Campaigns.CampaignWednesday, Campaigns.CampaignTuesday, Campaigns.CampaignMonday, Campaigns.CampaignSunday, "
    				strSQL2=strSQL2 & " Campaigns.CampaignDistribution, Campaigns.CampaignDailyEnd, Campaigns.CampaignDailyStart, Campaigns.CampaignCost, "
    				strSQL2=strSQL2 & " Campaigns.CampaignQuantitySold, Campaigns.CampaignType, Campaigns.CampaignEndDate, Campaigns.CampaignStartDate, Campaigns.CampaignName, "
    				strSQL2=strSQL2 & " Advertisers.CompanyNamE FROM validcampaigns_type INNER JOIN Advertisers ON validcampaigns_type.AdvertiserID = Advertisers.AdvertiserID INNER "
    				strSQL2=strSQL2 & " JOIN Campaigns ON validcampaigns_type.CampaignID = Campaigns.CampaignID "
				strSQL2=strSQL2 & " Where (Campaigns.UserID=" & CLng(Session("BanManProSiteID")) & " Or Campaigns.UserID=0)"			

				If Request("AdvertiserID") <> "" Then
					strSQL2=strSQL2 & " AND Advertisers.AdvertiserID=" & CLng(Request("AdvertiserID"))
				End If
			End If
			strSQL2=strSQL2 & " Order By Campaigns.CampaignName ASC"

			Set rsc=connBanManPro.Execute(strSQL2)
			If Not rsc.EOF Then
				strMessage="Listing of all Campaigns in Database."	
				'call include file and create table of all data
				%>
				<!--#include file="showallcampaigns.asp"-->
				<%
			End If
			Set rsc=Nothing
		Case "Expired"
			strSQL2="SELECT Campaigns.*, Advertisers.CompanyName, CampaignClicks.Clicks "
			strSQL2=strSQL2 & "FROM (Advertisers RIGHT JOIN Campaigns ON Advertisers.AdvertiserID = "
			strSQL2=strSQL2 & "Campaigns.AdvertiserID) INNER JOIN CampaignClicks ON Campaigns.CampaignID = "
			strSQL2=strSQL2 & "CampaignClicks.CampaignClickID "
			strSQL2=strSQL2 & "WHERE (((((Campaigns.CampaignEndDate) < getdate()) OR (Campaigns.CampaignType='CPM' AND Campaigns.CampaignQuantitySold<=Campaigns.CampaignImpressionsServed) OR (Campaigns.CampaignType='Per Click' AND Campaigns.CampaignQuantitySold<=CampaignClicks.Clicks)) AND CampaignSiteDefault<>1) AND (Campaigns.UserID=" & CLng(Session("BanManProSiteID")) & " Or Campaigns.UserID=0))"
			If Request("AdvertiserID") <> "" Then
				strSQL2=strSQL2 & " AND Advertisers.AdvertiserID=" & CLng(Request("AdvertiserID"))
			End If
			strSQL2=strSQL2 & " ORDER BY Campaigns.CampaignName"
			Set rsc=connBanManPro.Execute(strSQL2)
			If Not rsc.EOF Then
				strMessage="Listing of all Expired Campaigns."	
				'call include file and create table of all data
				%>
				<!--#include file="showallcampaigns.asp"-->
				<%
			End If
			Set rsc=Nothing
		Case "Link"
				%>
				<!--#include file="viewadcodes.asp"-->
				<%
		Case Else
	End Select
%>


<% ''''''''''''''''''''Change blank fields to " "   '''''''''''''''''''''''''''''''''''''''''''
Function FixBlank(strParameter)
	If Trim(strParameter)="" Then
		FixBlank=" "
	Else
		FixBlank=Replace(strParameter, "'", "''")
		FixBlank=Trim(FixBlank)
	End If
End Function  %>

<% ''''''''''''''''''''Set False Check box to 0 " "   '''''''''''''''''''''''''''''''''''''''''''
Function SetTrueFalse(strParameter)
	If Trim(strParameter)="-1" Then
		SetTrueFalse=-1
	Else
		SetTrueFalse=0
	End If
End Function %>
<% '''''''''''''''''''''change blank field to 0 '''''''''''''''''''''''''''''''''''''''''''''''''
Function FixZero(strData)
	If Trim(strData)="" Then
		FixZero=0
	Else
		FixZero=strData
	End If
End Function
%>