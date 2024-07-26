<%


	'******************************************************************************
	'Modify this section for System DSN to connect to database
	'BannerManagerConnectString="DSN=banmanproSQL;UID=banmanproADMIN;pwd=admin1"
	'******************************************************************************

	'******************************************************************************
	'Modify this section for connecting directly to the database using OLEDB
	BannerManagerConnectString="PROVIDER=SQLOLEDB;Data Source=208.37.41.168;UID=sa;PWD=tapatio;DATABASE=BanManPro"
	'******************************************************************************
	

	If Trim(Application("CacheBustingMode"))="" Or Application("AdministratorName")="" Or Application("BannerManagerConnectString")="" Then

		'set database connection
		Set connBanManPro=Server.CreateObject("ADODB.Connection") 
		connBanManPro.Mode = 3      '3 = adModeReadWrite
		connBanManPro.Open BannerManagerConnectString

		'grab administrative data
		strSQL="SELECT * FROM Administrative"
		Set rs=connBanManPro.Execute(strSQL)

		Application.Lock

		'Store Database Connection
		Application("BannerManagerConnectString")=BannerManagerConnectString

		'username
		Application("AdministratorName")=rs("AdministratorName")

		'password
		Application("AdministratorPassword")=rs("AdministratorPassword")

		'email Address
		Application("AdministratorEmail")=rs("AdministratorEmail")

		'server path
		Application("ServerPath")=rs("ServerPath")

		'domain url
		Application("DomainURL")=rs("DomainURL")

		'Email Program
		Application("MailProgram")=rs("MailProgram")

		'EmailWhenCampaignExpires
		Application("EmailWhenCampaignExpires")=rs("EmailWhenCampaignExpires")
	
		'MailServer
		Application("MailServer")=rs("MailServer")
	
		'CacheBustingMode
		Application("CacheBustingMode")=rs("CacheBustingMode")

		'version 2.0 parameters
		'Date Format
		Application("DateFormat")=rs("DateFormat")

		'Unique Click Hour
		Application("UniqueClickHour")=rs("UniqueClickHour")

		'Database Update Frequency
		Application("DatabaseUpdateFrequency")=rs("DatabaseUpdateFrequency")

		'MS Option
		Application("BanManProMultiSite")=False

		'Reports
		Application("BanManProDailyReport")=rs("DailyReport")
		Application("BanManProWeeklyReport")=rs("WeeklyReport")

		'Smoothing Minutes
		Application("BanManProSmoothingMinutes")=Clng(rs("SmoothingMinutes"))

		'Show ZoneStats
		Application("ZoneAverageDays")=rs("ZoneAverageDays")

		'Slot Option, Guranteed Impressions
		Application("SlotOption")=rs("SlotOption")
		Application("GuaranteedImpressionsPerSlot")=Clng(rs("GuaranteedImpressionsPerSlot"))

		'Average Campaign Length
		Application("StandardCampaignLength")=rs("StandardCampaignLength")		

		'Send email notification X days before campaign expires
		Application("BanManProDaysBeforeExpiration")=5	

		Application.Unlock

		'destroy recordset
		Set rs=Nothing
	End If

	'record referring web site
	'If Trim(Session("HTTP_REFERER"))="" Then
	'	Session("HTTP_REFERER") = Request.ServerVariables("HTTP_REFERER")
	'End If
%>

