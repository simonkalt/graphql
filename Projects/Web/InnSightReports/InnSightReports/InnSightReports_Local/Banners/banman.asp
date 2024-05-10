<!--#Include file="emergencystop.asp"-->
<!--#include file="banmanfunc.asp"-->
<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Product:  Ban Man Pro Version 2.01
'   Author:   Joe Rohrbach of Brookfield Consultants
'   Notes:    Main Module for Getting Banners/Counting Click Thrus
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
'   (c) Copyright 1999-2000 by Brookfield Consultants.  All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	'retrieve task
	If Trim(Request.QueryString("Task")) <> "" Then
		strTask=Request.QueryString("Task")
	End If

	'determine mode
	If Request.QueryString("Mode")="HTML" Then
		strBMPMode="HTML"
	ElseIf Request.QueryString("Mode")="TEXT" Then
		strBMPMode="TEXT"
	Else 
		strBMPMode="SSI"
	End If

	'determine Site ID
	If Application("BanManProMultiSite")=True Then
		If Trim(Request.QueryString("SiteID")) <> "" Then
			lngBMPSiteID=Request.QueryString("SiteID")
		ElseIf IsNumeric(lngBMPSiteID) And lngBMPSiteID<>"" Then
			'Already defined
		Else
			lngBMPSiteID=0
		End If
	Else
		lngBMPSiteID=0
	End If

	If strTask="Get" And Request("Browser")="NETSCAPE4" And Request("NoCache")="True" Then
		If Instr(Ucase(Request.ServerVariables("HTTP_USER_AGENT")),"MSIE")>0 Then
			Response.Buffer=True
			Response.ContentType="application/x-javascript"
			Response.Write "document.write(' '); "
			Response.End
		End If
	End If


	'User has clicked on banner ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Select Case strTask
	Case "Click"

		If strBMPMode="HTML" Then
			strZoneID=CLng(Request.QueryString("ZoneID"))
			If Request.QueryString("PageID")<> "" Then
				strBMPPageID="PageID_" & Trim(Request.QueryString("PageID")) & "_"
			Else
				strBMPPageID=""
			End If
			strBMPTemp="BannerID_" & strBMPPageID & strZoneID
			strBMPBannerID=Session(strBMPTemp)
			strBMPTemp="AdvertiserID_" & strBMPPageID & strZoneID
			strBMPAdvertiserID=Session(strBMPTemp)
			strBMPTemp="CampaignID_" & strBMPPageID & strZoneID
			strBMPCampaignID=Session(strBMPTemp)
			If Trim(strBMPBannerID)="" Then
				If Trim(HTTP_REFERER) <> "" Then
					Response.Redirect HTTP_REFERER
				Else	
					Response.Redirect "unavail.htm"
				End If
			End If
		Else
			'gather ID's from query string
			strBMPBannerID=CLng(Request.QueryString("BannerID"))
			strBMPAdvertiserID=CLng(Request.QueryString("AdvertiserID"))
			strBMPCampaignID=CLng(Request.QueryString("CampaignID"))
			strZoneID=CLng(Request.QueryString("ZoneID"))
		End If

		strBMPTargetURL=ClickBanManProAd(strBMPAdvertiserID,strBMPBannerID,strBMPCampaignID,strZoneID,lngBMPSiteID)

		'Redirect user to URL
		'Session.Abandon


		If Trim(Request("BanManProRedirect"))<>"" Then
			strBMPTargetURL=Request("BanManProRedirect")
			If InStr(strBMPTargetURL,"BMPQString") >0 Then
				strBMPTargetURL=Replace(strBMPTargetURL,"BMPQString","?")
			End If
			If InStr(strBMPTargetURL,"BMPAMPSAND") >0 Then
				strBMPTargetURL=Replace(strBMPTargetURL,"BMPAMPSAND","&")
			End If

			Response.Redirect strBMPTargetURL
		Else
			Response.Redirect strBMPTargetURL
		End If

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'User pulling banner from server ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Case "Get"

		If Request.QueryString("ZoneID") <> "" Then
			strZoneID=Request.QueryString("ZoneID")
		End If

		If Request.QueryString("Keywords")<>"" Then
			Keywords=Request.QueryString("Keywords")
		Else
			Keywords=""
		End If

		If Request.QueryString("ZoneName")<>"" Then
			ZoneName=Request.QueryString("ZoneName")
		Else
			ZoneName=""
		End If
	
		If strBMPMode="TEXT" Then
			ServeBanManProAdDirectly Clng(Request.QueryString("AdvertiserID")),Clng(Request.QueryString("BannerID")),Clng(Request.QueryString("CampaignID")),Clng(Request.QueryString("ZoneID")),lngBMPSiteID,"TEXT"
		Else
			'Serve Banner Ad Using GetBanManProAd(ZoneID,ZoneName,Keywords,Mode,SiteID)
			GetBanManProAd strZoneID,ZoneName,Keywords,strBMPMode,lngBMPSiteID
		End If

	Case Else
	End Select

	IncludedBMPAlready=True


%>