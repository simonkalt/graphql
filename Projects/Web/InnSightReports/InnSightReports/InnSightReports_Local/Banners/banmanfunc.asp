<!--#include file="dbconnect.asp"-->

<%
'*****************
'Version 2.04
'*****************

'*************************************************************************************************
'     Function for Clicking a Ban Man Pro Ad
'*************************************************************************************************
Function ClickBanManProAd(AdvertiserID,BannerID,CampaignID,ZoneID,SiteID)

		Dim rsTemp

		strBMPAdvertiserID=CLng(AdvertiserID)
		strBMPBannerID=CLng(BannerID)
		strBMPCampaignID=CLng(CampaignID)
		strZoneID=CLng(ZoneID)

		'Establish Database Connection
		Set connBanManPro = gConnADODB(Application("BannerManagerConnectString"))

		If IsNumeric(SiteID) Then
			lngBMPSiteID=Clng(SiteID)
		Else
			Set rsTemp=connBanManPro.Execute("sp_BMP_RetrieveSiteIDBySiteName '" & SiteID & "'")
			If Not rsTemp.EOF Then
				lngBMPSiteID=rsTemp("SiteID")
			Else
				lngBMPSIteID=0
			End If 
			Set rsTemp=Nothing
		End If
			
		'get banner information
		strBMPSQL="sp_BMP_ObtainTargetURL " & strBMPBannerID 
		Set rsBanManPro=connBanManPro.Execute(strBMPSQL)

		'target URL
		strBMPTargetURL=rsBanManPro("AdTargetURL")

		'Destroy Recordset
		Set rsBanBanPro=Nothing

		'Replace Random Number
		If InStr(strBMPTargetURL, "[RandomNumber]") >0 And Request.QueryString("Mode")<>"HTML" Then
			strBMPTargetURL = Replace(strBMPTargetURL,"[RandomNumber]", Request.QueryString("RandomNumber"))
		End If

		'Determine if user is clicking same banner multiple times, only insert record once
		'****Number of hours between clicks from each user*********
		'setting to zero will record each and every click, 12 makes them unique in 12 hours
		'valid values are integers 0,1,2....12...etc.
		intUniqueClickHour=Clng(Application("UniqueClickHour"))
		'**********************************************************
		strBMPSQL2="sp_BMP_ObtainPreviousClick " & strBMPCampaignID & "," & strBMPBannerID & "," & strBMPAdvertiserID & "," & strZoneID & "," & lngBMPSiteID & "," & intUniqueClickHour & ",'" & Request.ServerVariables("REMOTE_ADDR") & "'"
		Set rsBanManPro=connBanManPro.Execute(strBMPSQL2)
		If rsBanManPro.EOF= True Or intUniqueClickHour=0 Then

			'Update Campaign database to account for click through
			strBMPSQL2="sp_BMP_UpdateCampaignClicks " & strBMPCampaignID
			connBanManPro.Execute strBMPSQL2,,AdExecuteNoRecords

			'Update Clicks Table based on Impression ID
			strBMPSQL="sp_BMP_UpdateClickCounts " & strBMPCampaignID & "," & strBMPBannerID & "," & strBMPAdvertiserID & "," & strZoneID & "," & lngBMPSiteID & "," & Year(Date) & "," & Month(Date) & "," & Day(Date)
			connBanManPro.Execute strBMPSQL,,AdExecuteNoRecords

			'Add new record with Click information in Clicks database
			strBMPSQL2="sp_insert_Clicks_1 "
			strBMPSQL2=strBMPSQL2 & strBMPCampaignID & "," 
			strBMPSQL2=strBMPSQL2 & strZoneID & "," 
			strBMPSQL2=strBMPSQL2 & strBMPAdvertiserID & "," 
			strBMPSQL2=strBMPSQL2 & strBMPBannerID & "," 
			strBMPSQL2=strBMPSQL2 & lngBMPSiteID & ",'" 
			strBMPSQL2=strBMPSQL2 & FixBlank(Request.ServerVariables("REMOTE_ADDR")) & "','" 
			strBMPSQL2=strBMPSQL2 & FixBlank(Request.ServerVariables("HTTP_HOST")) & "','" 
			strBMPSQL2=strBMPSQL2 & " ','" 
			If Application("DateFormat")="MM/DD/YYYY"  Then
				strBMPSQL2=strBMPSQL2 & month(date()) & "/" & day(date()) & "/" & year(date()) & " " & hour(time()) & ":" & minute(time()) & ":" & second(time()) & "','"
			Else 'DD/MM/YY
				strBMPSQL2=strBMPSQL2 & day(date()) & "/" & month(date()) & "/" & year(date()) & " " & hour(time()) & ":" & minute(time()) & ":" & second(time()) & "','"
			End If
			strBMPSQL2=strBMPSQL2 & " ','" 
			strBMPSQL2=strBMPSQL2 & " ','" 
			strBMPSQL2=strBMPSQL2 & FixBlank(Request.ServerVariables("HTTP_USER_AGENT")) & "',' '"
			connBanManPro.Execute strBMPSQL2,,AdExecuteNoRecords

		End If

		'close database connection
		closeConnection connBanManPro
		Set rsBanBanPro=Nothing


		If Trim(Request("BanManProRedirect"))<>"" Then
			strBMPTargetURL=Request("BanManProRedirect")
			If InStr(strBMPTargetURL,"BMPQString") >0 Then
				strBMPTargetURL=Replace(strBMPTargetURL,"BMPQString","?")
			End If
			If InStr(strBMPTargetURL,"BMPAMPSAND") >0 Then
				strBMPTargetURL=Replace(strBMPTargetURL,"BMPAMPSAND","&")
			End If
		End If

		ClickBanManProAd=strBMPTargetURL

End Function



'*************************************************************************************************
'     Function for Retrieving a Ban Man Pro Ad
'*************************************************************************************************
Sub GetBanManProAd(ZoneID,ZoneName,Keywords,Mode,SiteID)

		Dim strPreviousDistribution,lngSumEven,connBanManPro,rsTemp

		strBMPMode=Mode
		strKeywords=Keywords		
		blnTryDefaultCampaign=False
		blnFoundBanner=False
		blnFoundCampaign=False
		strPreviousDistribution="Even"
		If IsNumeric(ZoneID) Then
			strZoneID=ZoneID
		Else
			strZoneID=0
		End If

		'Establish Database Connection
		Set connBanManPro = gConnADODB(Application("BannerManagerConnectString"))

		If IsNumeric(SiteID) Then
			lngBMPSiteID=Clng(SiteID)
		Else
			Set rsTemp=connBanManPro.Execute("sp_BMP_RetrieveSiteIDBySiteName '" & SiteID & "'")
			If Not rsTemp.EOF Then
				lngBMPSiteID=rsTemp("SiteID")
			Else
				lngBMPSIteID=0
			End If 
			Set rsTemp=Nothing
		End If

		'First Check if user is calling by keyword.  Keywords always take priority
		If Trim(strKeywords)<>"" Then
			'Look for Keywords
			strSQL="sp_BMP_ObtainKeywordCampaigns"
			Set rsDistinct=connBanManPro.Execute(strSQL)
			'See if a campaign exists with this keyword
			Do While Not rsDistinct.EOF
				If Instr(Ucase(rsDistinct("CampaignKeywords")),UCase(strKeywords))>0 Then
					'use this Campaign
					blnFoundCampaign=True
					lngRandom=25
					intSum=50
					strBMPCampaignID=rsDistinct("CampaignID")
					Exit Do
				End If
				rsDistinct.MoveNext
			Loop
		End If

		If blnFoundCampaign=False Then
			'retrieve valid campaigns for this zone
			If IsNumeric(ZoneID) Then
				strZoneID=CLng(ZoneID)
				strBMPSQL="sp_BMP_RetrieveValidCampaigns " & strZoneID & "," & lngBMPSiteID
			Else

				'call by name
				strBMPSQL="sp_BMP_RetrieveValidCampaignsByDesc '" & (ZoneName) & "'," & lngBMPSiteID
			End If
			Set rsDistinct=connBanManPro.Execute(strBMPSQL)
			
			'Find sum of even campaigns
			lngSumEven=0
			If Not rsDistinct.EOF Then
				strZoneID=Clng(rsDistinct("ZoneID"))
				strBMPSQL="sp_BMP_RetrieveSumValidEvenCampaigns " & rsDistinct("ZoneID") & "," & lngBMPSiteID
				Set rsBanManPro=connBanManPro.Execute(strBMPSQL)
				If Not rsBanManPro.EOF Then
					lngSumEven=rsBanManPro("SumOfEven")
				End If
				Set rsBanManPro=Nothing
			End If
		End If

		'Set Randomizer Just once
		Rnd (-1)  'some bizarre systems also require this
		Randomize

		If Not rsDistinct.EOF Or blnFoundCampaign=True Then

			If blnFoundCampaign=False Then
				rsDistinct.MoveFirst

				'get random number between 1 to 1000
				If lngSumEven>1000 Or Application("SlotOption")=True Then
					lngRandom=Int((lngSumEven - 1 + 1) * Rnd + 1)
				Else
					lngRandom=Int((1000 - 1 + 1) * Rnd + 1)
				End If
				intSum=0

				Do While Not rsDistinct.EOF
					If strPreviousDistribution="Even" And rsDistinct("Even")=False Then
						strPreviousDistribution="Weighted"
						'reset sum of weightings
						intSum=0
						lngRandom=Int((100 - 1 + 1) * Rnd + 1)
					End If
					strBMPCampaignID=rsDistinct("CampaignID")
					strBMPZoneCampaignWeighting=rsDistinct("ZoneCampaignWeighting")
					strZoneID=rsDistinct("ZoneID")
					intSum=intSum + Cint(strBMPZoneCampaignWeighting)
					If lngRandom <= intSum Then
						'use this Campaign
						Exit Do
					Else
						rsDistinct.MoveNext
					End If
				Loop
			End If

			If lngRandom > intSum Then
				'no campaigns to display
				strBMPCampaignID=0
				strBMPBannerID=0
				strBMPAdvertiserID=0
				blnTryDefaultCampaign=True
			Else
				'retrieve all banners under this campaign
				strBMPSQL="sp_BMP_RetrieveValidBanners " & Clng(strBMPCampaignID)

				Set rsBanners=connBanManPro.Execute(strBMPSQL)

				'now determine which banner to show
				'get random number between 1 to 100
				'Randomize
				lngRandom=Int((100 - 1 + 1) * Rnd + 1)
				
				intSum=0
				Do While Not rsBanners.EOF
					strBMPBannerID=rsBanners("BannerID")
					strBMPBannerWeighting=rsBanners("CampaignBannerWeighting")
					strBMPAdvertiserID=rsBanners("AdvertiserID")
					intSum=intSum + CInt(strBMPBannerWeighting)
					If lngRandom <= intSum Then
						'use this Campaign
						Exit Do
					Else
						rsBanners.MoveNext
					End If
				Loop
				If lngRandom > intSum Then
					'no banners to display
					'set ID's to zero
					strBMPCampaignID=0
					strBMPBannerID=0
					strBMPAdvertiserID=0
					blnTryDefaultCampaign=True
				Else
					blnFoundBanner=True
					If strBMPMode="HTML" Then
						'redirect to banner after updating stats
						strBMPTargetBannerURL=rsBanners("AdImageURL")
						If Request.QueryString("PageID")<> "" Then
							strBMPPageID="PageID_" & Trim(Request.QueryString("PageID")) & "_"
						Else
							strBMPPageID=""
						End If
						strBMPTemp="BannerID_" & strBMPPageID & strZoneID
						Session(strBMPTemp)=strBMPBannerID
						strBMPTemp="AdvertiserID_" & strBMPPageID & strZoneID
						Session(strBMPTemp)=strBMPAdvertiserID
						strBMPTemp="CampaignID_" & strBMPPageID & strZoneID
						Session(strBMPTemp)=strBMPCampaignID
						strBMPTemp="ZoneID_" & strBMPPageID & strZoneID
						Session(strBMPTemp)=strZoneID
					Else
						blnAdFragment=rsBanners("AdFragment")
						'now create code for this banner
						strBMPCode=GetCode(strZoneID, strBMPCampaignID, strBMPAdvertiserID, strBMPBannerID, rsBanners("AdTargetURL"), rsBanners("AdAltText"), rsBanners("AdImageURL"), rsBanners("AdBorder"), rsBanners("AdWidth"), rsBanners("AdHeight"), rsBanners("AdAlign"), rsBanners("AdNewWindow"), rsBanners("AdTextUnderneath"), Application("DomainURL"), strBMPMode,blnAdFragment,"",lngBMPSiteID)
						If Request.QueryString("Browser")="NETSCAPE4" Then
							Response.Buffer=True
							Response.ContentType="application/x-javascript"
							strBMPCode=Replace(strBMPCode,"'","\'")
							strBMPCode2=Replace(strBMPCode,vbCRLF," ")
							Response.Write "document.write('" & strBMPCode2 & "'); "
							If Request.QueryString("NoCache")="" Then
								strBMPCode2="adcode=' '"
								Response.Write Replace(strBMPCode2,vbCRLF," ")
							End If
						Else
							Response.Write strBMPCode
						End If
					End If
				End If
			End If
		Else
			'no campaigns to display
			strBMPCampaignID=0
			strBMPBannerID=0
			strBMPAdvertiserID=0
			blnTryDefaultCampaign=True
		End If
		'*****************************************************************************************
		If blnTryDefaultCampaign=True Then

			'Determine if user has included default campaigns
			If IsNumeric(ZoneID) Then
				strBMPSQL="sp_BMP_RetrieveDefaultZones " & strZoneID & "," & lngBMPSiteID 
			Else
				'call by name
				strBMPSQL="sp_BMP_RetrieveDefaultZonesByZoneName '" & ZoneName & "'," & lngBMPSiteID
			End If		

			Set rsBanners=connBanManPro.Execute(strBMPSQL)
			If Not rsBanners.EOF Then
				strZoneID=rsBanners("ZoneID")
				'retrieve default banners of this size
				If NOT rsBanners.EOF Then
					'determine number of defaults
					intCntDefaults=0
					If rsBanners.EOF <> True Then
						'find sum of all banners, just in case it is less than 100
						Do While NOT rsBanners.EOF
							intCntDefaults=intCntDefaults+rsBanners("CampaignBannerWeighting")
							rsBanners.MoveNext
						Loop
						rsBanners.MoveFirst
						lngRandom=Int((intCntDefaults - 1 + 1) * Rnd + 1)
						intSum=0
						Do While Not rsBanners.EOF
							strBMPBannerID=rsBanners("BannerID")
							strBMPBannerWeighting=rsBanners("CampaignBannerWeighting")
							strBMPAdvertiserID=rsBanners("AdvertiserID")
							strBMPCampaignID=rsBanners("CampaignID")
							intSum=intSum + CInt(strBMPBannerWeighting)
							If lngRandom <= intSum Then
								'use this Campaign
								Exit Do
							Else
								rsBanners.MoveNext
							End If
						Loop
						If intCntDefaults=0 Then
							rsBanners.MoveFirst
						End If
						blnFoundBanner=True
						If strBMPMode="HTML" Then
							'redirect to banner after updating stats
							strBMPTargetBannerURL=rsBanners("AdImageURL")
							If Request.QueryString("PageID")<> "" Then
								strBMPPageID="PageID_" & Trim(Request.QueryString("PageID")) & "_"
							Else
								strBMPPageID=""
							End If
							strBMPTemp="BannerID_" & strBMPPageID & strZoneID
							Session(strBMPTemp)=strBMPBannerID
							strBMPTemp="AdvertiserID_" & strBMPPageID & strZoneID
							Session(strBMPTemp)=strBMPAdvertiserID
							strBMPTemp="CampaignID_" & strBMPPageID & strZoneID
							Session(strBMPTemp)=strBMPCampaignID
							strBMPTemp="ZoneID_" & strBMPPageID & strZoneID
							Session(strBMPTemp)=strZoneID
						Else
							'now create code for this banner
							'IF (rsBanners("AdNewWindow")=False AND rsBanners("AdFragment")<>False AND Request("Browser")="NETSCAPE4") Then
							'	blnAdFragment=False
							'Else
								blnAdFragment=rsBanners("AdFragment")
							'End If

							strBMPCode=GetCode(strZoneID, strBMPCampaignID, strBMPAdvertiserID, strBMPBannerID, rsBanners("AdTargetURL"), rsBanners("AdAltText"), rsBanners("AdImageURL"), rsBanners("AdBorder"), rsBanners("AdWidth"), rsBanners("AdHeight"), rsBanners("AdAlign"), rsBanners("AdNewWindow"), rsBanners("AdTextUnderneath"), Application("DomainURL"), strBMPMode,blnAdFragment,"",lngBMPSiteID)
							If Request.QueryString("Browser")="NETSCAPE4" Then
								Response.Buffer=True
								Response.ContentType="application/x-javascript"
								strBMPCode=Replace(strBMPCode,"'","\'")
								strBMPCode2=Replace(strBMPCode,vbCRLF," ")
								Response.Write "document.write('" & strBMPCode2 & "'); "
								If Request.QueryString("NoCache")="" Then
									strBMPCode2="adcode=' '"
									Response.Write Replace(strBMPCode2,vbCRLF," ")
								End If
							Else
								Response.Write strBMPCode
							End If
						End If
					Else
						strBMPCampaignID=0
						strBMPBannerID=0
						strBMPAdvertiserID=0
					End If
				Else
					'set ID's to zero, nothing to display, track page hit
					strBMPCampaignID=0
					strBMPBannerID=0
					strBMPAdvertiserID=0
				End If
			Else
				strBMPCampaignID=0
				strBMPBannerID=0
				strBMPAdvertiserID=0
			End If
		End If

		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'close database connection
		closeConnection connBanManPro
		Set rsBanManPro=Nothing
		Set rsDistinct=Nothing
		Set rsBanners=Nothing


		'Count Impression
		CountBanManProImpression strBMPAdvertiserID,strBMPBannerID,strBMPCampaignID,strZoneID,lngBMPSiteID


		'SEND Email notification if necessary **********************************************
		If Application("EmailWhenCampaignExpires")<>0 Then
			CheckBanManProCampaignExpiration strBMPCampaignID,lngBMPSiteID
		End If

	
		If strBMPMode="HTML" Then
			'redirect to banner
			If Trim(strBMPTargetBannerURL) <> "" Then
				Response.Redirect strBMPTargetBannerURL
			Else
				strBMParPath = Split(Request.ServerVariables("SCRIPT_NAME"), "/")
				strBMParPath(UBound(strBMParPath,1)) = ""
				Response.Redirect Join(strBMParPath, "/") & "blank.gif"
			End If
		Else
			If blnFoundBanner=False Then
				If Request.QueryString("Browser")="NETSCAPE4" Then
					Response.Buffer=True
					Response.ContentType="application/x-javascript"
					If Request.QueryString("NoCache")="True" Then
						Response.Write "document.write(' '); "
					Else
						strBMPCode2="adcode=' '"
						Response.Write strBMPCode2
					End If
				Else
					Response.Write " "
				End If
			End If
		End If

End Sub



'***************************************************************************
' Subroutine for Counting Impressions
'***************************************************************************
Sub CountBanManProImpression(AdvertiserID,BannerID,CampaignID,ZoneID,SiteID)


		strBMPAdvertiserID=AdvertiserID
		strBMPBannerID=BannerID
		strBMPCampaignID=CampaignID
		strZoneID=ZoneID
		lngBMPSiteID=SiteID


		'Establish Database Connection
		Set connBanManPro = gConnADODB(Application("BannerManagerConnectString"))
		
		'Update impressions table if it does not exist for zone/day, create record
		'All this is now done in one stored procedure
		'NOTE**** The Date must absolutely be Month/Day/Year in ALL COUNTRIES
		'NOTE**** The internal stored proceedure will convert to your countries date format
		strBMPSQL="sp_BMP_UpdateCampaignAndImpressions "
		strBMPSQL=strBMPSQL & Clng(strBMPAdvertiserID) & ","
		strBMPSQL=strBMPSQL & Clng(strBMPBannerID) & ","
		strBMPSQL=strBMPSQL & Clng(strBMPCampaignID) & ","
		strBMPSQL=strBMPSQL & Clng(strZoneID) & ","
		strBMPSQL=strBMPSQL & Clng(lngBMPSiteID) & ","
		strBMPSQL=strBMPSQL & "'" & month(date()) & "/" & day(date()) & "/" & year(date()) & "',0,0"
		connBanManPro.Execute strBMPSQL,,AdExecuteNoRecords

		'close database connection
		closeConnection connBanManPro

		'Update Zone Impressions Used for Smoothing Algorithm
		strZoneName="BMPZoneImpressions_" & strZoneID
		If IsNumeric(Application(strZoneName))=False Then
			Application.Lock
			Application(strZoneName)=1
			Application.Unlock
		Else
			If IsDate(Application("BanManProStatsStartTime"))=False Then
				Application.Lock
				Application("BanManProStatsStartTime")=Now()
				Application("BanManProStatsHour")=Hour(Now())
				Application("BanManProStatsActualTime")=Now()
				Application("BanManProStatsDay")=Day(Date() & " 00:30:00")
				Application("BanManProStatsInProgress")=False
				Application.Unlock
				If Application("SlotOption")<>True Then
					CalculateBanManProExpectedQuantity
				End If
				
				'Establish Database Connection
				Set connBanManPro = gConnADODB(Application("BannerManagerConnectString"))
				'Update DateSent
				strSQL="Update Administrative Set ReportsSentDate=CONVERT(DATETIME,'" & month(date()) & "/" & day(date()) & "/" & year(date()) & "',101)"
				connBanManPro.Execute strSQL,,AdExecuteNoRecords
				closeConnection connBanManPro
			End If
			Application.Lock
			Application(strZoneName)=CLng(Application(strZoneName)) + 1
			Application.Unlock
			lngMinutesDifference=Clng(Abs(DateDiff("n",Application("BanManProStatsStartTime"),Now())))
			If lngMinutesDifference> Clng(Application("BanManProSmoothingMinutes"))*2 Then
				'An error occurred and Application("BanManProStatsInProgress") must be reset
				Application.Lock
				Application("BanManProStatsInProgress")=False
				Application.Unlock
			End If
			If ((lngMinutesDifference>= Clng(Application("BanManProSmoothingMinutes"))) OR (Application("BanManProStatsHour")<>Hour(Now()))) And (Application("BanManProStatsInProgress")=False OR Trim(Application("BanManProStatsInProgress"))="") Then
				Application.Lock
				Application("BanManProStatsInProgress")=True
				Application.Unlock

				'reset stats
				If Application("BanManProStatsHour")<>Hour(Now()) Then

					'********************************************
					'Perform Hourly Routines
					'********************************************
					If Application("SlotOption")<>True Then
						Application.Lock
						Application("BanManProStatsHour")=Hour(Now())
						Application("BanManProStatsActualTime")=Now()
						Application.Unlock
						CalculateBanManProExpectedQuantity
					Else
						Application.Lock
						Application("BanManProStatsHour")=Hour(Now())
						Application("BanManProStatsActualTime")=Now()
						Application.Unlock
					End If


					If Application("BanManProStatsDay")<> Day(Now()) Then

						'Establish Database Connection
						Set connBanManPro = gConnADODB(Application("BannerManagerConnectString"))
						strSQL="Select ReportsSentDate From Administrative"
						Set rsTemp=connBanManPro.Execute(strSQL)
						If Not rsTemp.EOF Then
							If IsDate(rsTemp("ReportsSentDate")) Then
								Application.Lock
								Application("BanManProStatsDay")=Day(Now())
								Application.Unlock

								If (Year(rsTemp("ReportsSentDate"))=Year(Date())) And (Month(rsTemp("ReportsSentDate"))=Month(Date())) AND (Day(rsTemp("ReportsSentDate"))=Day(Date())) Then
									'Do Not Send Report
								Else

									'********************************************
									'Perform Daily Routines 
									'********************************************
									'Send Daily Email Reports if Necessary
									SendBanManProEmailReports 1
									'Check if Day is Sunday, if so, send weekly's
									If DatePart("w",Now(),1)=vbSunday Then
										SendBanManProEmailReports 7
									End If
									'UpdateBanManProZoneAverages 1
									'update zone averages
									If IsNumeric(Application("ZoneAverageDays")) Then
										UpdateBanManProZoneAverages Application("ZoneAverageDays")
									End If

									'For Slot Option, check if extensions are required
									If Application("SlotOption")=True Then
										ExtendSlotCampaigns
									End If

								End If
							End If
						End If

						'Update DateSent
						strSQL="Update Administrative Set ReportsSentDate=CONVERT(DATETIME,'" & month(date()) & "/" & day(date()) & "/" & year(date()) & "',101)"
						connBanManPro.Execute strSQL,,AdExecuteNoRecords

						Set rsTemp=Nothing
						closeConnection connBanManPro
					End If
				End If
				If Application("SlotOption")<>True Then
					CalculateBanManProEvenWeightings(SiteID)
				End If
				Application.Lock
				Application("BanManProStatsStartTime")=Now()
				Application.Unlock

				Application.Lock
				Application("BanManProStatsInProgress")=False
				Application.Unlock
			End If
		End If



End Sub



'***************************************************************************
' Subroutine for Updating Zone Averages for X Period IN dAYS
'***************************************************************************
Sub UpdateBanManProZoneAverages(Period)

	Dim rsBanManProLocal,strSQL,arrData

	'Establish Database Connection
	Set connBanManProLocal = gConnADODB(Application("BannerManagerConnectString"))

	' Execute our query and get back a RS
	Set rsBanManProLocal = Server.CreateObject("ADODB.Recordset")
	strSQL="sp_BMP_RetrieveZoneIDs "
	rsBanManProLocal.Open strSQL, connBanManProLocal
	If not rsBanManProLocal.EOF Then
		arrData = rsBanManProLocal.GetRows	
	End If
	rsBanManProLocal.Close
	Set rsBanManProLocal=Nothing
	closeConnection  connBanManProLocal

	If IsArray(arrData) Then
		'Establish Database Connection
		Set connBanManProLocal = gConnADODB(Application("BannerManagerConnectString"))

		'Now update daily zone averages in database based on yesterday's stats
		For intCounter = LBound(arrData, 2) to UBound(arrData, 2)
			strSQL="sp_BMP_CalculateZoneAverages " & arrData(0,intCounter) & "," & Clng(Period)
			connBanManProLocal.Execute strSQL,,AdExecuteNoRecords
		Next

		'close database connection
		closeConnection connBanManProLocal
		Set rsBanManProLocal=Nothing
	End If

End Sub

'***************************************************************************
' Subroutine for Extending Slot Campaigns
'***************************************************************************
Sub ExtendSlotCampaigns()

	Dim rsBanManProLocal,strSQL

	If IsNumeric(Application("GuaranteedImpressionsPerSlot")) Then
		If Clng(Application("GuaranteedImpressionsPerSlot"))>0 Then

			'Establish Database Connection
			Set connBanManProLocal = gConnADODB(Application("BannerManagerConnectString"))

			'Check for campaigns that should be extended
			strSQL="sp_BMP_RetrieveCampaignsToExtend " & Clng(Application("GuaranteedImpressionsPerSlot"))
			Set rsBanManProLocal=connBanManProLocal.Execute(strSQL)

			Do While Not rsBanManProLocal.EOF
				strSQL="sp_BMP_UpdateCampaignsToExtend " & rsBanManProLocal("CampaignID")
				connBanManProLocal.Execute strSQL,,AdExecuteNoRecords
				rsBanManProLocal.MoveNext
			Loop

			'close database connection
			closeConnection connBanManProLocal
			Set rsBanManProLocal=Nothing

		End If
	End If

End Sub

'***************************************************************************
' Subroutine for Checking if any expired campaign emails should be sent
'***************************************************************************
Sub CheckBanManProCampaignExpiration(CampaignID,SiteID)

		Dim lngDaysExpire

		strBMPCampaignID=CampaignID
		lngBMPSiteID=SiteID

		'Establish Database Connection
		Set connBanManPro = gConnADODB(Application("BannerManagerConnectString"))

		strSQLBMPEmail="sp_BMP_ObtainExpiredCampaigns " & strBMPCampaignID & "," & lngBMPSiteID
		Set rsBMPEmail=connBanManPro.Execute(strSQLBMPEmail)
		If NOT rsBMPEmail.EOF Then

			If rsBMPEmail("CampaignSiteDefault") <> True And Clng(rsBMPEmail("CampaignSiteDefault")) <> 1 Then

				If (rsBMPEmail("CampaignImpressionsServed") >= rsBMPEmail("CampaignQuantitySold")) OR (rsBMPEmail("Clicks") >= rsBMPEmail("CampaignQuantitySold")) OR (rsBMPEmail("CampaignEndDate") < Date() ) Then
					'update database to reflect sent notification
					strSQLBMPEmail="UPDATE Campaigns Set CampaignNotificationSent=1 Where CampaignID=" & rsBMPEmail("CampaignID")
					connBanManPro.Execute strSQLBMPEmail,,AdExecuteNoRecords

					'send email notification to manager
					strBMPMailServer=Application("MailServer")
					strSMPRecipient=Application("AdministratorEmail")
					strBMPSubject="Alert** Ban Man Pro Expired Campaign"
					strBMPMessage="The campaign titled " & rsBMPEmail("CampaignName") & " has expired."
					strBMPFrom=Application("AdministratorEmail")	
					%>
					<!--#include file="mail.asp"-->
					<%
					rsBMPEmail.MoveNext
				Else
					If rsBMPEmail("CampaignWarningSent")<>True Then
						'update database to reflect sent notification
						strSQLBMPEmail="UPDATE Campaigns Set CampaignWarningSent=1 Where CampaignID=" & rsBMPEmail("CampaignID")
						connBanManPro.Execute strSQLBMPEmail,,AdExecuteNoRecords
						'send email notification to manager warning of expiration
						strBMPMailServer=Application("MailServer")
						strSMPRecipient=Application("AdministratorEmail")
						strBMPSubject="Alert** Ban Man Pro Campaign Approaching Expiration"
						strBMPFrom=Application("AdministratorEmail")	
						If rsBMPEmail("CampaignType")="CPM" Then
							strBMPMessage="The campaign titled " & rsBMPEmail("CampaignName") & " expires in " & Clng(rsBMPEmail("CampaignQuantitySold"))-Clng(rsBMPEmail("CampaignImpressionsServed")) & " impressions."
							%>
							<!--#include file="mail.asp"-->
							<%
							rsBMPEmail.MoveNext
						ElseIf rsBMPEmail("CampaignType")="Per Click" Then
							strBMPMessage="The campaign titled " & rsBMPEmail("CampaignName") & " expires in " & Clng(rsBMPEmail("CampaignQuantitySold"))-Clng(rsBMPEmail("Clicks")) & " clicks."
							%>
							<!--#include file="mail.asp"-->
							<%
							rsBMPEmail.MoveNext
						End If
					End If
				End If
			End If
		End If
		Set rsBMPEmail=Nothing

		'Send about to expire message if using Slot Option or Flat Rate Campaign
		If IsNull(Application("BanManProDaysBeforeExpiration"))=False Then
			If Clng(Application("BanManProDaysBeforeExpiration"))>0 Then
				lngDaysExpire=Clng(Application("BanManProDaysBeforeExpiration"))
			Else
				lngDaysExpire=5
			End If
		Else
			lngDaysExpire=5
		End If
		strSQL="SELECT CampaignID,CampaignName, CampaignStartDate, CampaignEndDate, CampaignType FROM Campaigns "
		strSQL=strSQL & " WHERE (CampaignType = N'Flat Rate') AND (GETDATE()  >= DATEADD(day, -" & lngDaysExpire & ", CampaignEndDate)) AND  CampaignWarningSent = 0 And CampaignEndDate>getDate()"
		Set rsBMPEmail=connBanManPro.Execute(strSQL)		
		If Not rsBMPEmail.EOF Then
			'update database to reflect sent notification
			strSQLBMPEmail="UPDATE Campaigns Set CampaignWarningSent=1 Where CampaignID=" & rsBMPEmail("CampaignID")
			connBanManPro.Execute strSQLBMPEmail,,AdExecuteNoRecords
			'send email notification to manager warning of expiration
			strBMPMailServer=Application("MailServer")
			strSMPRecipient=Application("AdministratorEmail")
			strBMPSubject="Alert** Ban Man Pro Campaign Approaching Expiration"
			strBMPFrom=Application("AdministratorEmail")	
			strBMPMessage="The campaign titled " & rsBMPEmail("CampaignName") & " expires on " & rsBMPEmail("CampaignEndDate") 
			%>
			<!--#include file="mail.asp"-->
			<%
		End If
		Set rsBMPEmail=Nothing

		'close database connection
		closeConnection connBanManPro

End Sub

'*************************************************************************************************
'     Sub for Computing/Updating Expected Impressions for Even Campaigns
'*************************************************************************************************

Sub CalculateBanManProExpectedQuantity()
	
	Dim FirstDate,LastDate,blnFoundNow,lngTotalHours,DayOfWeek,connBanManProLocal
	Dim lngHoursInCampaign,lngExpectedImpressions,strSQLUpdate,rsC,lngHoursToDate
	Dim arrData,intCounter

	'Establish Database Connection
	Set connBanManProLocal = gConnADODB(Application("BannerManagerConnectString"))

	'Grab All Even Campaigns
	strSQL="sp_BMP_RetrieveValidCampaignsForCalculatingExpectedQuantity"

	' Execute our query and get back a RS
	Set rsC = Server.CreateObject("ADODB.Recordset")
	rsC.CursorLocation = 3 
	rsC.Open strSQL, connBanManProLocal
	Set rsC.ActiveConnection = Nothing

	Do While Not rsC.EOF

		'Set start/end date
		FirstDate=rsC("CampaignStartDate")
		LastDate=rsC("CampaignEndDate")
		blnFoundNow=False
		lngTotalHours=0
		Do Until FirstDate>LastDate
			'Determine Day of Week
			DayOfWeek=DatePart("w",FirstDate,1,1)
			If (DayOfWeek=vbSunday And rsC("CampaignSunday")=True) OR (DayOfWeek=vbMonday And rsC("CampaignMonday")=True) OR (DayOfWeek=vbTuesday And rsC("CampaignTuesday")=True) OR (DayOfWeek=vbWednesday And rsC("CampaignWednesday")=True) OR (DayOfWeek=vbThursday And rsC("CampaignThursday")=True) OR (DayOfWeek=vbFriday And rsC("CampaignFriday")=True) OR (DayOfWeek=vbSaturday And rsC("CampaignSaturday")=True) Then
				If (Hour(rsC("CampaignDailyStart"))=Hour(rsC("CampaignDailyEnd"))) OR  ((DatePart("h",Firstdate,1,1)<Hour(rsC("CampaignDailyEnd"))) AND (DatePart("h",Firstdate,1,1)>=Hour(rsC("CampaignDailyStart")))) Then 
					'Include This Hour
					lngTotalHours=lngTotalHours+1
					'Response.Write "Date=" & FirstDate & "     DayOfWeek=" & DayOfWeek & "   " & "Hour=" & DatePart("h",Firstdate,1,1) & "<br>"
				End If
			End If
			Firstdate=DateAdd("h",1,FirstDate)
			If FirstDate>=Now And blnFoundNow=False Then
				blnFoundNow=True
				lngHoursToDate=lngTotalHours
			End If
		Loop

		lngHoursInCampaign=lngTotalHours
		lngExpectedImpressions=(rsC("CampaignQuantitySold")/lngHoursInCampaign)*lngHoursToDate
		strSQLUpdate="sp_BMP_UpdateCampaignExpectedQuantity " & lngExpectedImpressions & "," & rsC("CampaignID")
		connBanManProLocal.Execute strSQLUpdate,,adExecuteNoRecords		
		rsC.MoveNext
	Loop

	Set rsC=Nothing	
	closeConnection  connBanManProLocal

End Sub


'*************************************************************************************************
'     Sub for Computing/Updating Weightings for Even Campaigns
'*************************************************************************************************
Sub CalculateBanManProEvenWeightings(SiteID)

	Dim strZoneName,lngZoneImpressions,lngAvailable,lngCampaignID,lngCampaignIDNew
	Dim datBanManProPreviousHour,lngPercentAvailable,lngMinutesDifference,lngMinutesDifferenceSinceHourStarted
	Dim connBanManProLocal,strSQL2,arrData,intCounter
	
	If Trim(SiteID)="" Then
		SiteID=0
	End If

	'Determine Scaling Factor
	datBanManProPreviousHour=Application("BanManProStatsStartTime")

	If IsDate(datBanManProPreviousHour) Then
		lngMinutesDifference=DateDiff("n",datBanManProPreviousHour,Now())
	Else
		lngMinutesDifference=0
	End If

	If IsDate(Application("BanManProStatsActualTime")) Then
		lngMinutesDifferenceSinceHourStarted=60-DateDiff("n",Application("BanManProStatsActualTime"),Now())
		If lngMinutesDifferenceSinceHourStarted<5 Then
			lngMinutesDifferenceSinceHourStarted=60
		End If
	Else
		lngMinutesDifferenceSinceHourStarted=60
	End If

	If lngMinutesDifference>4 Then

		'Establish Database Connection
		Set connBanManProLocal = gConnADODB(Application("BannerManagerConnectString"))

		'Find all zones including this campaign
		strSQL="sp_BMP_RetrieveValidCampaignsForCalculatingEvenWeightings"

		' Execute our query and get back a RS
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strSQL, connBanManProLocal
		If not rs.EOF Then
			arrData = rs.GetRows
		End If
		rs.Close
		Set rs=Nothing
		closeConnection  connBanManProLocal


		If IsArray(arrData) Then

			'Each X Minutes, count total impressions in each zone and make hourly value
			lngZoneImpressions=0
			For intCounter = LBound(arrData, 2) to UBound(arrData, 2)
				'ZoneID=0
				'CampaignID=1
				'UserID=2
				'ZoneCampaignWeighting=3
				'Even=4
				'CampaignQuantitySold=5
				'impressions served=6
				'quantity expected=7
				'start date=8
				'end date=9
		
			   ' If Clng(arrData(2,intCounter))=Clng(SiteID) OR Clng(arrData(2,intCounter))=0 Then
				strZoneName="BMPZoneImpressions_" & Trim(Cstr(arrData(0,intCounter)))
				lngZoneImpressions=lngZoneImpressions + Application(strZoneName)

				lngAvailable=arrData(7,intCounter)-arrData(6,intCounter)
				If IsNumeric(lngAvailable) Then
					'If negative, set available to zero
					If Clng(lngAvailable)<=0 Then
						lngAvailable=0
					Else
						If (arrData(5,intCounter)=arrData(7,intCounter)) And (lngMinutesDifferenceSinceHourStarted<=Application("BanManProSmoothingMinutes")) Then
							lngAvailable=lngAvailable+(0.45*lngAvailable)
						End If
					End If
				Else
					lngAvailable=0
				End If
				lngCampaignID=CLng(arrData(1,intCounter))


				If intCounter<UBound(arrData, 2) Then
					lngCampaignIDNew=arrData(1,intCounter+1)
				End If
				If lngCampaignID<>lngCampaignIDNew Or (intCounter=UBound(arrData, 2)) Then
					'count zone
					'In Each zone committ percentage of impressions to this even campaign
					If lngAvailable<0 Then
						'Not available
						lngPercentAvailable=0
					Else
						'Available
						'First Scale lngZoneImpressions
						If lngMinutesDifference<=0 Then
							lngPercentAvailable=1
						Else
							If lngZoneImpressions>0 Then
								lngZoneImpressions=(lngZoneImpressions/lngMinutesDifference)*lngMinutesDifferenceSinceHourStarted
								lngPercentAvailable=(lngAvailable/lngZoneImpressions)*1000
								If lngPercentAvailable>1000 Then
									lngPercentAvailable=1000
								End If
							Else
								lngPercentAvailable=1
							End If
						End If
	
					End If	
					lngZoneImpressions=0
					If IsNumeric(lngPercentAvailable)=False Then
						lngPercentAvailable=0
					End If
					Set connBanManProLocal = gConnADODB(Application("BannerManagerConnectString"))
					strSQL2="sp_BMP_UpdateZoneCampaignWeightings " & Clng(lngPercentAvailable) & "," & CLng(lngCampaignID)
					connBanManProLocal.Execute strSQL2,,AdExecuteNoRecords
					closeConnection connBanManProLocal
				End If
   			   ' End If
			Next
		End if

		'reset all zone impressions
		Set connBanManProLocal = gConnADODB(Application("BannerManagerConnectString"))
		strSQL="sp_BMP_RetrieveZoneIDs"
		Set rs=connBanManProLocal.Execute(strSQL)
		Do While Not rs.EOF
			strZoneName="BMPZoneImpressions_" & rs("ZoneID")
			Application.Lock
			Application(strZoneName)=0
			Application.Unlock
			rs.MoveNext
		Loop
		'close database connection
		Set rs=Nothing
		closeConnection connBanManProLocal



	End If

End Sub



'*************************************************************************************************
'     Sub for Sending Email Reports To Advertisers
'*************************************************************************************************
Sub SendBanManProEmailReports(Period)


	Dim lngAdvertiserID,lngImpressions,lngClicks,strHeader,strEmail,strTemp,lngAdvertiserIDNew
	Dim strAdminMessage,lngAdminImpression,lngAdminClicks,datYesterday

	'Establish Database Connection
	Set connBanManPro = gConnADODB(Application("BannerManagerConnectString"))

	'grab yesterday's date
	datYesterday=FormatDateTime(DateAdd("d",-1,Date()),vbShortDate)

	'Execute Stored procedure based on yesterday or weekly stats
	If Period=1 Then
		strSQL="sp_BMP_ObtainYesterdaysStats"
		Set rs=connBanManPro.Execute(strSQL)
		strBMPSubject="Daily Advertising Report for " & datYesterday
	ElseIf Period=7 Then
		strSQL="sp_BMP_ObtainWeeklyStats"
		Set rs=connBanManPro.Execute(strSQL)
		strBMPSubject="Weekly Advertising Report Ending " & datYesterday
	End If

	If Not rs.EOF Then
		lngAdvertiserID=rs("AdvertiserID")
		lngAdvertiserIDNew=rs("AdvertiserID")
	End If


	'Loop Through And Email Reports
	strHeader=            "Campaign Name             Impressions            Clicks            Click Rate" & Chr(10)
	strHeader=strHeader & "-----------------------------------------------------------------------------" & Chr(10)
	strBMPMessage=strHeader
	Do While Not rs.EOF

		lngImpressions=lngImpressions+rs("SumOfImpressionCount")
		lngClicks=lngClicks+rs("SumOfClicks")
		lngAdminImpression=lngAdminImpression+rs("SumOfImpressionCount")
		lngAdminClicks=lngAdminClicks+rs("SumOfClicks")

		strEmail=rs("Email")
		If Len(rs("CampaignName")) >= 24 Then
			strTemp=strTemp & Left(rs("CampaignName"),24) & "...  "
		Else
			strTemp=strTemp & rs("CampaignName") & String(29-Len(rs("CampaignName"))," ")
		End If
		strTemp=strTemp & rs("SumOfImpressionCount") & String(23-Len(rs("SumOfImpressionCount"))," ")
		strTemp=strTemp & rs("SumOfClicks") & String(18-Len(rs("SumOfClicks"))," ")
		If rs("SumOfImpressionCount")>0 Then
			strTemp=strTemp & FormatPercent((rs("SumOfClicks")/rs("SumOfImpressionCount"))) & Chr(10)
		Else
			strTemp=strTemp & FormatPercent(0) & Chr(10)
		End If
		strAdminMessage=strAdminMessage & strTemp
		rs.MoveNext
		If Not rs.EOF Then
			lngAdvertiserIDNew=rs("AdvertiserID")
		End If
		If (lngAdvertiserIDNew <> lngAdvertiserID) OR rs.EOF=True Then

			'Send Message To Advertiser			
			strBMPMailServer=Application("MailServer")
			strSMPRecipient=strEmail			
			strBMPMessage=strHeader & strTemp & Chr(10) 
			strBMPMessage=strBMPMessage & "-----------------------------------------------------------------------------" & Chr(10)
			strBMPMessage=strBMPMessage & "Totals                       " & lngImpressions & String(23-Len(lngImpressions)," ") & lngClicks & String(18-Len(lngClicks)," ")
			If lngImpressions>0 Then
				strBMPMessage=strBMPMessage & FormatPercent((lngClicks/lngImpressions)) & Chr(10)
			Else
				strBMPMessage=strBMPMessage & FormatPercent(0) & Chr(10)
			End If
			strBMPFrom=Application("AdministratorEmail")	
			%>
			<!--#include file="mail.asp"-->
			<%
			strTemp=""
			lngImpressions=0
			lngClicks=0
			lngAdvertiserID=lngAdvertiserIDNew
		End If
	Loop

	If (Application("BanManProDailyReport") <> 0 And Period=1) Then
		strSQL="sp_BMP_ObtainYesterdaysStatsAdmin"
		Set rs=connBanManPro.Execute(strSQL)
		strBMPSubject="Daily Advertising Report for " & datYesterday
	ElseIf (Application("BanManProWeeklyReport") <> 0 And Period=7) Then
		strSQL="sp_BMP_ObtainWeeklyStatsAdmin"
		Set rs=connBanManPro.Execute(strSQL)
		strBMPSubject="Weekly Advertising Report Ending " & datYesterday
	End If

	strAdminMessage=""
	strTemp=""
	'Loop Through And Email Reports
	lngAdminImpression=0
	Do While Not rs.EOF
		lngAdminImpression=lngAdminImpression+rs("SumOfImpressionCount")
		lngAdminClicks=lngAdminClicks+rs("SumOfClicks")
		If Len(rs("CampaignName")) >= 24 Then
			strTemp=strTemp & Left(rs("CampaignName"),24) & "...  "
		Else
			strTemp=strTemp & rs("CampaignName") & String(29-Len(rs("CampaignName"))," ")
		End If
		strTemp=strTemp & rs("SumOfImpressionCount") & String(23-Len(rs("SumOfImpressionCount"))," ")
		strTemp=strTemp & rs("SumOfClicks") & String(18-Len(rs("SumOfClicks"))," ")
		If rs("SumOfImpressionCount")>0 Then
			strTemp=strTemp & FormatPercent((rs("SumOfClicks")/rs("SumOfImpressionCount"))) & Chr(10)
		Else
			strTemp=strTemp & FormatPercent(0) & Chr(10)
		End If
		strAdminMessage=strAdminMessage & strTemp
		strTemp=""
		rs.MoveNext
	Loop


	If strAdminMessage<>"" Then
		'Send Administrator Reports
		strBMPMailServer=Application("MailServer")
		strSMPRecipient=Application("AdministratorEmail")
		strBMPFrom=Application("AdministratorEmail")	
		strBMPMessage=strHeader & strAdminMessage & Chr(10) 
		strBMPMessage=strBMPMessage & "-----------------------------------------------------------------------------" & Chr(10)
		strBMPMessage=strBMPMessage & "Totals                       " & lngAdminImpression & String(23-Len(lngAdminImpression)," ") & lngAdminClicks & String(18-Len(lngAdminClicks)," ")
		If lngAdminImpression>0 Then
			strBMPMessage=strBMPMessage & FormatPercent((lngAdminClicks/lngAdminImpression)) & Chr(10)
		Else
			strBMPMessage=strBMPMessage & FormatPercent(0) & Chr(10)
		End If
		%>
		<!--#include file="mail.asp"-->
		<%
	End If


	Set rs=Nothing
	'close database connection
	closeConnection connBanManPro

End Sub

'*************************************************************************************************
'     Function for Directly Serving Image called by ID
'*************************************************************************************************
Sub ServeBanManProAdDirectly(AdvertiserID,BannerID,CampaignID,ZoneID,SiteID,Mode)

	Dim strImageURL,blnAdFragment,strBMPCode,strBMPCode2,rsTemp


	If IsNumeric(SiteID)=False Then
		'Establish Database Connection
		Set connBanManPro = gConnADODB(Application("BannerManagerConnectString"))
		Set rsTemp=connBanManPro.Execute("sp_BMP_RetrieveSiteIDBySiteName '" & SiteID & "'")
		If Not rsTemp.EOF Then
			SiteID=rsTemp("SiteID")
		Else
			SIteID=0
		End If 
		Set rsTemp=Nothing
	End If

	CountBanManProImpression AdvertiserID,BannerID,CampaignID,ZoneID,SiteID


	strImageURL=""
	'Serve Ad Based on Mode
	If Mode="HTML" Then
		'Establish Database Connection
		Set connBanManPro = gConnADODB(Application("BannerManagerConnectString"))

		'Retrieve ImageURL
		strSQL="sp_BMP_RetrieveImageURLByBannerID " & Clng(BannerID)
		Set rs=connBanManPro.Execute(strSQL)
		If Not rs.EOF Then
			strImageURL=rs("AdImageURL")
		Else
			strImageURL="blank.gif"
		End If
		Set rs=Nothing
		closeConnection connBanManPro
	ElseIF Mode="TEXT" Then
		strImageURL="blank.gif"
	Else
		'Establish Database Connection
		Set connBanManPro = gConnADODB(Application("BannerManagerConnectString"))

		'Grab Banner information for this BannerID
		strSQL="sp_BMP_RetrieveBannerByBannerID " & Clng(BannerID)
		Set rs=connBanManPro.Execute(strSQL)
		If Not rs.EOF Then
			IF (rs("AdNewWindow")=False AND rs("AdFragment")<>False AND Request("Browser")="NETSCAPE4") Then
				blnAdFragment=False
			Else
				blnAdFragment=rs("AdFragment")
			End If
			'now create code for this banner
			strBMPCode=GetCode(ZoneID, CampaignID, AdvertiserID, BannerID, rs("AdTargetURL"), rs("AdAltText"), rs("AdImageURL"), rs("AdBorder"), rs("AdWidth"), rs("AdHeight"), rs("AdAlign"), rs("AdNewWindow"), rs("AdTextUnderneath"), Application("DomainURL"), Mode,blnAdFragment,"",SiteID)
		Else
			strBMPCode=" "
		End IF

		'destroy recordset, close connection
		Set rs=Nothing
		closeConnection connBanManPro

		If Request.QueryString("Browser")="NETSCAPE4" Then
			If Request.QueryString("NoCache")="True" Then
				Response.Buffer=True
				Response.ContentType="application/x-javascript"
				strBMPCode=Replace(strBMPCode,"'","\'")
				strBMPCode2=Replace(strBMPCode,vbCRLF," ")
				Response.Write "document.write('" & strBMPCode2 & "'); "
			Else
				Response.Buffer=True
				Response.ContentType="application/x-javascript"
				strBMPCode=Replace(strBMPCode,"'","\'")
				strBMPCode2="adcode='" & strBMPCode & "'"
				Response.Write Replace(strBMPCode2,vbCRLF," ")
			End If
		Else
			Response.Write strBMPCode
		End If
	End If


	'Send Image to browser
	If strImageURL <> "" Then
		Response.Redirect strImageURL
	End If
	

End Sub


'*************************************************************************************************
'     Function for Checking if a Zone Exists
'		Usage: Exists=CheckIfBanManProZoneExists(ZoneValue,ZoneIDOrName)
'		Example: Exists=CheckIfBanManProZoneExists(4,"ID")  'call by ID
'		Example2: Exists=CheckIfBanManProZoneExists("MyZoneName","ZoneName")  'call by name
'		Return Value: Function returns true or false
'*************************************************************************************************
Function CheckIfBanManProZoneExists(ZoneValue,ZoneIDOrName)

	Dim connBanManProLocal

	If ZoneIDOrName="ID" Then
		'Establish Database Connection
		Set connBanManPro = gConnADODB(Application("BannerManagerConnectString"))

		'Execute Stored Proceedure
		Set rsBanManProLocal=connBanManPro.Execute("sp_BMP_RetrieveZoneID " & Clng(ZoneValue))
	 
		If Not rsBanManProLocal.EOF Then
			CheckIfBanManProZoneExists=True
		Else
			CheckIfBanManProZoneExists=False
		End If

		Set rsBanManProLocal=Nothing
		'close database connection
		closeConnection connBanManPro

	ElseIf ZoneIDOrName="ZoneName" Then

		'Establish Database Connection
		Set connBanManPro = gConnADODB(Application("BannerManagerConnectString"))
		
		'Execute Stored Proceedure
		Set rsBanManProLocal=connBanManPro.Execute("sp_BMP_RetrieveZoneIDByZoneName '" & ZoneValue & "'")

		If Not rsBanManProLocal.EOF Then
			CheckIfBanManProZoneExists=True
		Else
			CheckIfBanManProZoneExists=False
		End If

		Set rsBanManProLocal=Nothing
		'close database connection
		closeConnection connBanManPro
	Else
		CheckIfBanManProZoneExists=False
	End If

	

End Function


'***************************************************************************
' Function to establish database connection
'***************************************************************************
Function gConnADODB (ConnectString)
   Dim Conn
   Set Conn = Server.Createobject("ADODB.Connection")
   Conn.ConnectionTimeout = 30
   Conn.Open ConnectString
   Set gConnADODB = Conn
End Function

'***************************************************************************
' Function to create command object
'***************************************************************************
Function getCommand (Conn)
   Dim Cmd
   Set Cmd                  = Server.CreateObject("ADODB.Command")
   Set Cmd.ActiveConnection = Conn
   Set getCommand           = Cmd
End Function

'***************************************************************************
' Function to close database connection
'***************************************************************************
Sub closeConnection (Conn)
   If (IsObject(Conn) ) Then
      Conn.Close
      Set Conn= Nothing
   End If
End Sub



''''''''''''''''''''Change blank fields to " "   '''''''''''''''''''''''''''''''''''''''''''
Function FixBlank(strBMPParameter)
	If Trim(strBMPParameter)="" Then
		FixBlank=" "
	Else
		FixBlank=Replace(strBMPParameter,"'","''")
	End If
End Function 


'********** This Function Creates the HTML code returned to a browser*********************
''''''''''''create ad code '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetCode(ZoneID, CampaignID, AdvertiserID, BannerID, AdTargetURL, AdAltText, AdImageURL, AdBorder, AdWidth, AdHeight, AdAlign, AdNewWindow, AdTextUnderneath, DomainURL, ZoneMode,AdFragment,AdCode,SiteID)

Dim AddNewWindow

'Retrieve Random Number for cache defeating
lngRandom2=Int(Rnd*100000)


If AdFragment=True Then
	'Establish Database Connection
	Set connBanManPro = gConnADODB(Application("BannerManagerConnectString"))
	strSQL="sp_BMP_ObtainAdCode " & BannerID
	Set rsBanManPro=connBanManPro.Execute(strSQL)
	strBMPAdCode=rsBanManPro("AdCodeNetscape")
	If AdNewWindow=False And Trim(strBMPAdCode) <> "" And Request.QueryString("Browser")="NETSCAPE4" Then
		'do nothing
	ElseIF AddNewWindow=False And Trim(strBMPAdCode) ="" And Request.QueryString("Browser")="NETSCAPE4"  Then
		AdFragment=False
	Else
		strBMPAdCode=rsBanManPro("AdCode")
	End If

	'Close Connection
	Set rsBanManPro=Nothing
	closeConnection connBanManPro
End If
If AdFragment <> True Then
	'create URL string    
	strBMPDomString=DomainURL & "?Task=Click&ZoneID=" & ZoneID & "&CampaignID=" & CampaignID & "&AdvertiserID=" & AdvertiserID & "&BannerID=" & BannerID & "&SiteID=" & SiteID
	
	'cache busting code
	If Application("CacheBustingMode")=-1 Or Application("CacheBustingMode")="-1" Then
		strBMPDomString=strBMPDomString & "&RandomNumber=" & lngRandom2
		If InStr(AdImageURL,"[RandomNumber]") >0 Then
			AdImageURL=Replace(AdImageURL,"[RandomNumber]",lngRandom2)
		Else
			If InStr(AdImageURL,"?") >0 Then
				AdImageURL=AdImageURL & "&RandomNumber=" & lngRandom2
			ElseIf Right(AdImageURL,1)="/" Then
				AdImageURL=AdImageURL & lngRandom2
			Else
				AdImageURL=AdImageURL & "?RandomNumber=" & lngRandom2
			End If
		End If
	End If

	'Launch in New Window
	strBMPURLString =  strBMPDomString

	'strBMPAdCode="" 
	If Request("Browser")<> "NETSCAPE4" Then
		strBMPAdCode="<p align=" & AdAlign & ">"
	End If
	If AdNewWindow=True Then
	     strBMPAdCode = strBMPAdCode & "<a href=" & Chr(34) & strBMPURLString & Chr(34) & " target=" & Chr(34) & "_new" & Chr(34) & "><img src=" & Chr(34) & AdImageURL & Chr(34)
        Else
	     strBMPAdCode = strBMPAdCode & "<a href=" & Chr(34) & strBMPURLString & Chr(34) & " target=" & Chr(34) & "_top" & Chr(34) & "><img src=" & Chr(34) & AdImageURL & Chr(34)
	End If

    	strBMPAdCode = strBMPAdCode & "  width=" & Chr(34) & AdWidth & Chr(34) & " height=" & Chr(34) & AdHeight & Chr(34) & " alt=" & Chr(34) & Trim(AdAltText) & Chr(34) & " align=" & Chr(34) & AdAlign & Chr(34) & " border=" & Chr(34) & AdBorder & Chr(34) & "></a><br>"
    	If Trim(AdTextUnderneath) <> "" Then
		If AdNewWindow=True Then
			strBMPAdCode = strBMPAdCode & "<a href=" & Chr(34) & strBMPURLString & Chr(34) & " target=" & Chr(34) & "_new" & Chr(34) & ">" & AdTextUnderneath & "</a>"
    		Else
			strBMPAdCode = strBMPAdCode & "<a href=" & Chr(34) & strBMPURLString & Chr(34) & " target=" & Chr(34) & "_top" & Chr(34) & ">" & AdTextUnderneath & "</a>"
		End If
	End If
End If

	'Replace [RandomNumber] with a random number
	If InStr(strBMPAdCode,"[RandomNumber]") >0 Then
		strBMPAdCode=Replace(strBMPAdCode,"[RandomNumber]",lngRandom2)
	End If

	'Replace [BanManProURL] with target URL
	If InStr(strBMPAdCode,"[BanManProURL]") >0 Then
		'create URL string
		If Trim(strBMPDomString)="" Then    
			strBMPDomString=DomainURL & "?Task=Click&ZoneID=" & ZoneID & "&CampaignID=" & CampaignID & "&AdvertiserID=" & AdvertiserID & "&BannerID=" & BannerID & "&SiteID=" & SiteID
		End If
		strBMPAdCode=Replace(strBMPAdCode,"[BanManProURL]",strBMPDomString)
	End If



	GetCode = strBMPAdCode

End Function

%>