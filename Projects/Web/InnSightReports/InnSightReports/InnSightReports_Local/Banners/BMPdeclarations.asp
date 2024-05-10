	
<%
	'******************************************************************************
	'Variable Declarations for Ban Man Pro Ad Server
	Dim strBMPMode,strBMPBannerWeighting,strBMPSQL2
	Dim strBMPBannerID,strBMPAdvertiserID,strBMPCampaignID,strBMPSQL,strBMPTargetURL
	Dim blnTryDefaultCampaign,lngRandom,intSum,strBMPZoneCampaignWeighting,rsDistinct
	Dim strBMPCode,strBMPCode2,intCntDefaults,strBMPTemp,strBMParPath
	Dim rsBanManPro,rsBanners,rsZoneDefault,rsDefaultCampaigns,rsZoneMode
	Dim strBMPAdCode,strBMPDomString,lngRandom2,strBMPURLString
	Dim intUniqueClickHour
	Dim strBMPMailServer,strSMPRecipient,strSQLBMPEmail
	Dim strBMPSubject,strBMPMessage,strBMPFrom,rsBMPEmail,Mailer
	Dim strZoneID,strTask,blnFoundBanner,blnAdFragment
	Dim connBanManPro,strYourDatabasePath,BannerManagerConnectString,strSQL,rs
	Dim lngBMPSiteID,strStatsCampaignCount,strStatsImpressionCount,IncludedBMPAlready
	Dim strKeywords,blnFoundCampaign,strZoneName,ZoneName,Keywords
	'End Variable Declarations

	'Begin Constants
	Const adCmdStoredProc = &H0004
	Const adExecuteNoRecords = &H00000080
	Const adLockReadOnly = 1
	Const adLockPessimistic = 2
	Const adLockOptimistic = 3
	Const adLockBatchOptimistic = 4

	'---- ConnectModeEnum Values ----
	Const adModeUnknown = 0
	Const adModeRead = 1
	Const adModeWrite = 2
	Const adModeReadWrite = 3
	Const adModeShareDenyRead = 4
	Const adModeShareDenyWrite = 8
	Const adModeShareExclusive = &Hc
	Const adModeShareDenyNone = &H10
	Const adModeRecursive = &H400000

	'---- CursorLocationEnum Values ----
	Const adUseServer = 2
	Const adUseClient = 3
	
	Const adOpenForwardOnly = 0

	'******************************************************************************
%>