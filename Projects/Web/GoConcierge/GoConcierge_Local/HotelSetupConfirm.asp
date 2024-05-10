<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))


if len(trim(Request.QueryString("CompanyID"))) > 0 then
	strBackPage = "HotelSetupNewMain.asp"
	strCompanyID = trim(Request.QueryString("CompanyID"))
else
	strBackPage = "Switchboard3.asp"
	strCompanyID = remote.Session("CompanyID")
end if

%>

<!-- #include file="header.inc" -->
<!--#include file = "Map.inc" -->

<SCRIPT LANGUAGE=vbscript RUNAT=Server>
'C:\InnSight\InnSight_Local\ClientUploads\hilton_logo.jpg
Public Function c2FileName(pstrASPUpFileName, pstrType)
Dim strNewFileName
  'Response.Write "pstrASPUpFileName: " & pstrASPUpFileName & "<BR>"


  'If IsNull(pstrASPUpFileName) Then
   'Do this anyway
   ' Response.Write "pstrASPUpFileName was null<BR>"
    'Response.Write "Instr: " & InStrRev(pstrASPUpFileName,"\")
    'pstrDBFileName = Right(pstrASPUpFileName, Len(pstrASPUpFileName) - InStrRev(pstrASPUpFileName,"\"))
    'Response.Write "pstrDBFileName is now: " & pstrDBFileName & "<BR>"
  'End If
	strSuffix = Right(pstrASPUpFileName,4)

	'Rename the file in the db
	'Response.Write "Left: " & Left(pstrDBFileName,Len(Request.QueryString("ID"))+1) & "<BR>"
	'Response.Write "Right: " & pstrPreFix & Request.QueryString("ID") & "<BR>"
	'If Left(pstrDBFileName,Len(Request.QueryString("ID"))+1) = pstrPreFix & Request.QueryString("ID") Then
	  'Response.Write "Didn't add the prefix_<BR>"
	  'strNewFileName = Left(pstrDBFileName,Len(pstrDBFileName)-3) & strSuffix
	'Else
'	  Response.Write "Did add the prefix_<BR>"
	
	  strNewFileName = strCompanyID & "_" & pstrType & strSuffix
'	End If
		
	'Response.Write "This will be the new db filename: " & strNewFileName & "<BR>"

	c2FileName = strNewFileName
End Function


</SCRIPT>


<%

  Dim cnGN
  Dim rsGN
  Dim cmdGN

  Set cnGN = Server.CreateObject("ADODB.Connection")
  Set rsGN = Server.CreateObject("ADODB.Recordset")
  Set rsHL = Server.CreateObject("ADODB.Recordset") ' Hotel's Location
  Set cmdGN = Server.CreateObject("ADODB.Command")
  
  strConnect =  Application("sqlInnSight_ConnectionString")
  strUsername = Application("sqlInnSight_RuntimeUsername")
  strPassword = Application("sqlInnSight_RuntimePassword")
  
  cnGN.Open  strConnect, strUserName, strPassword
  'Response.Write "CompanyID: " & strCompanyID & "<BR>"
  
  cmdGN.CommandText = "Select * from tblCompany Where CompanyID= " & strCompanyID
  cmdGN.CommandType = 1 'adCmdText

  Set cmdGN.ActiveConnection = cnGN

  rsGN.Open cmdGN, ,adOpenKeyset,adLockPessimistic
  rsHL.LockType = adLockPessimistic
  rsHL.CursorType = adOpenKeyset
  rsHL.Open "Select * from tblLocation where LocationID=(select LocationID from tblCompany where CompanyID=" & strCompanyID & ")", cnGN
  
  
  'Response.Write "This is the current LogoLocation (i.e. LETTERHEAD) gif file: " & rsGN("LogoLocation") & "<BR>"
  'Response.Write "This is the current ScreenLogoLocation gif file: " & rsGN("ScreenLogoLocation") & "<BR>"

  'Need to interrogate these values
  
  'Response.Write "To Email: " & strToEMail & "<BR>"
  'Response.Write "From EMail: " & strFromEMail & "<BR>"
  'Response.Write "From Name: " & strFromName & "<BR>"
  
  dim myMailObject, strBodyClient
  Dim strFileName, fso
  
  set fso = CreateObject("Scripting.FileSystemObject")
  Set Upload = Server.CreateObject("Persits.Upload.1")

  Upload.OverwriteFiles = True ' Generate unique names
  
  Upload.SetMaxSize 1048576 

  'Response.Write "About to upload.<BR>"
  
  Count = Upload.Save(Application("ENV_PATH") & "CLIENTUPLOADS\")
  
  'Response.Write "About to enter file collection.<BR>"
  



  ' adding/editting a new hotel bombs here ***********************************************
	For Each File in Upload.Files 
	  'Response.Write "About to rename: " & File.Name & ", " & File.Path & "<BR>"
		Select Case File.Name
			Case "FILE1"
			    strDBFileName = c2FileName(File.Path, "Letterhead")
		        'Response.Write "This is the new strDBFileName: " & strDBFileName & "<BR>"
				rsGN.Fields("LogoLocation") = strDBFileName
				if strCompanyID = remote.Session("CompanyID") then
					remote.Session("LogoLocation") = strDBFileName
				end if
				rsGN.Update
			Case "FILE2"
			    strDBFileName = c2FileName(File.Path, "Logo")
		        'Response.Write "This is the new strDBFileName: " & strDBFileName & "<BR>"
				rsGN.Fields("ScreenLogoLocation") = strDBFileName
				if strCompanyID = remote.Session("CompanyID") then
					remote.Session("ScreenLogoLocation") = strDBFileName
				end if
				rsGN.Update
			Case Else
				'Response.Write "Error with case statement -- Renaming object.<BR>"
				Response.End
	   End Select
	' *******************************************************************************
	'Response.Write "ok"
	'Response.End

	'Response.Write "To: " & strDBFileName & "<BR>"
	'Response.Write ENVIRONMENTPATH & "CLIENTUPLOADS\" & strDBFileName
	'fso.DeleteFile Application("ENV_PATH") & "CLIENTUPLOADS\" & strDBFileName, true
	File.SaveAs(APPLICATION("ENV_PATH") & "CLIENTUPLOADS\" & strDBFileName)
  Next
  set fso = nothing  
  'Response.Write "Out of file collection.<BR>"
 
  'Response.Write "The new filename is: " & strDBFileName & "<BR>"
  'Response.Write "Your photo or movie trailer has been successfully uploaded.<BR><BR>"

	' Display description field
	'Response.Write Upload.Form("Description") & "<BR>"

	' Display all selected categories
	'For Each Item in Upload.Form
	  'If Item.Name = "Category" Then
	    'Response.Write Item.Value & "<BR>"
	  'End If
	'Next
  
    rsGN.Close

	Set cnSQL = Server.CreateObject("ADODB.Connection")
	Set rsCat = Server.CreateObject("ADODB.Recordset")
  
	cnSQL.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")

	' Debugging only
	'Response.Write "Update tblCompany Set tblCompany.CompanyName='" & Upload.Form("txtHotelName") & "', tblCompany.Address1='" & Upload.Form("txtHotelAddress1") & "', tblCompany.Address2='" & Upload.Form("txtHotelAddress2") & "', tblCompany.City='" & Upload.Form("txtHotelCity") & "', tblCompany.State='" & Upload.Form("cboHotelState") & "', tblCompany.PostalCode='" & Upload.Form("txtHotelPostalCode") & "', tblCompany.DirectionsAddress1='" & Upload.Form("txtDirectionsAddress1") & "', tblCompany.DirectionsAddress2='" & Upload.Form("txtDirectionsAddress2") & "', tblCompany.DirectionsCity='" & Upload.Form("txtDirectionsCity") & "', tblCompany.DirectionsState='" & Upload.Form("cboDirectionsState") & "', tblCompany.DirectionsPostalCode='" & Upload.Form("txtDirectionsPostalCode") & "', tblCompany.Phone=" & Upload.Form("txtPhone") & ", tblCompany.Fax=" & Upload.Form("txtFax") & ", tblCompany.Email='" & Upload.Form("txtEmail") & "', tblCompany.WebPage='" & Upload.Form("txtWebPage") & "', tblCompany.LogoLocation='" & Upload.Form("txtLetterheadBitmap") & "', tblCompany.ScreenLogoLocation='" & Upload.Form("txtScreenLogoBitmap") & "', tblCompany.UseCompanyLetterhead = " &  Upload.Form("cboLetterheadDefault") & ", tblCompany.LateTaskTime = " &  Upload.Form("txtLateMinutes") & ", tblCompany.LetterheadMargin = " &  Upload.Form("txtTopMargin") & ", tblCompany.LetterfootMargin = " &  Upload.Form("txtBottomMargin") & "<BR>"

    ' ATTN IMPT: the current CompanyID needs to be taken into account when you update this table!!
	'Set rsCat = cnSQL.Execute("Update tblCompany Set tblCompany.CompanyName='" & Upload.Form("txtHotelName") & "', tblCompany.Address1='" & Upload.Form("txtHotelAddress1") & "', tblCompany.Address2='" & Upload.Form("txtHotelAddress2") & "', tblCompany.City='" & Upload.Form("txtHotelCity") & "', tblCompany.State='" & Upload.Form("cboHotelState") & "', tblCompany.PostalCode='" & Upload.Form("txtHotelPostalCode") & "', tblCompany.DirectionsAddress1='" & Upload.Form("txtDirectionsAddress1") & "', tblCompany.DirectionsAddress2='" & Upload.Form("txtDirectionsAddress2") & "', tblCompany.DirectionsCity='" & Upload.Form("txtDirectionsCity") & "', tblCompany.DirectionsState='" & Upload.Form("cboDirectionsState") & "', tblCompany.DirectionsPostalCode='" & Upload.Form("txtDirectionsPostalCode") & "', tblCompany.Phone=" & Upload.Form("txtPhone") & ", tblCompany.Fax=" & Upload.Form("txtFax") & ", tblCompany.Email='" & Upload.Form("txtEmail") & "', tblCompany.WebPage='" & Upload.Form("txtWebPage") & "', tblCompany.LogoLocation='" & Upload.Form("txtLetterheadBitmap") & "', tblCompany.ScreenLogoLocation='" & Upload.Form("txtScreenLogoBitmap") & "', tblCompany.UseCompanyLetterhead = " &  Upload.Form("cboLetterheadDefault") & ", tblCompany.LateTaskTime = " &  Upload.Form("txtLateMinutes") & ", tblCompany.LetterheadMargin = " &  Upload.Form("txtTopMargin") & ", tblCompany.LetterfootMargin = " &  Upload.Form("txtBottomMargin"))

	'Rewritten by Joseph to make simpler and bug-less. 4/6/01
	' Excellent Joseph.  This IS Neat.  Thank you for correcting my old-style code.
	' Ideally this would be a stored procedure with transactions
	'   that could rollback just in case there's some catastrophe
	'   as this script is processing
	Response.Write strCompanyID
	rsCat.Open "Select * from tblCompany Where CompanyID=" & strCompanyID,cnSQL,adOpenKeyset,adLockPessimistic
	
    rsCat.Fields("Disclaimer").Value = Upload.Form("txtDisclaimer")
    rsCat.Fields("WeatherURL").Value = Upload.Form("txtWeatherURL")
    rsCat.Fields("TicketsURL").Value = Upload.Form("txtTicketsURL")
    rsCat.Fields("MoviesURL").Value = Upload.Form("txtMoviesURL")
    rsCat.Fields("ZagatURL").Value = Upload.Form("txtZagatURL")
    rsCat.Fields("FlightsURL").Value = Upload.Form("txtFlightsURL")
    
	rsCat.Fields("CompanyName").Value = Upload.Form("txtHotelName")
	rsHL.Fields ("CompanyName").Value = Upload.Form("txtHotelName")
	
	rsCat.Fields("Address1").Value = Upload.Form("txtHotelAddress1")
	rsHL.Fields ("Street").Value = Upload.Form("txtHotelAddress1")
	
	rsCat.Fields("Address2").Value = Upload.Form("txtHotelAddress2")
	rsCat.Fields("City").Value = Upload.Form("txtHotelCity") 
	rsHL.Fields ("City").Value = Upload.Form("txtHotelCity") 
	
	rsCat.Fields("State").Value = Upload.Form("cboHotelState") 
	rsHL.Fields ("State").Value = Upload.Form("cboHotelState") 
	
	
	rsCat.Fields("PostalCode").Value = Upload.Form("txtHotelPostalCode")
	rsHL.Fields ("ZIP").Value = Upload.Form("txtHotelPostalCode")
	
	
	rsCat.Fields("Phone").Value = Upload.Form("txtPhone")
	rsHL.Fields ("Phone").Value = Upload.Form("txtPhone")
	
	rsCat.Fields("Fax").Value = Upload.Form("txtFax")
	rsCat.Fields("EMail").Value = Upload.Form("txtEmail") 
	rsCat.Fields("WebPage").Value = Upload.Form("txtWebPage")
	rsCat.Fields("InternetAccess").Value = 1
	rsCat.Fields("UseCompanyLetterhead").Value = Upload.Form("cboLetterheadDefault")
	
	
	
	
	rsCat.Fields("MapScalingFactor").Value = 0 
	rsCat.Fields("SuppressBlanks").Value = 1
	rsCat.Fields("DirectionsDefault").Value = 1 
	rsCat.Fields("LetterheadMargin").Value = Upload.Form("txtTopMargin")
	
	rsCat.Fields("DirectionsAddress1").Value = Upload.Form("txtDirectionsAddress1")
	rsHL.Fields ("altStreet").Value = Upload.Form("txtDirectionsAddress1")
	
	
	rsCat.Fields("DirectionsAddress2").Value = Upload.Form("txtDirectionsAddress2")
	rsCat.Fields("DirectionsCity").Value = Upload.Form("txtDirectionsCity")
	rsHL.Fields ("altCity").Value = Upload.Form("txtDirectionsCity")
	
	rsCat.Fields("DirectionsState").Value = Upload.Form("cboDirectionsState")
	rsHL.Fields ("altState").Value = Upload.Form("cboDirectionsState")
	
	rsCat.Fields("DirectionsPostalCode").Value = Upload.Form("txtDirectionsPostalCode")
	rsHL.Fields ("altZip").Value = Upload.Form("txtDirectionsPostalCode")
	
	rsCat.Fields("ContactFirstName").Value = Upload.Form("txtContactFirst")
	rsCat.Fields("ContactLastName").Value = Upload.Form("txtContactLast")
	
	rsCat.Fields("OTHotelID").Value = e2n(Upload.Form("txtOTHotelID"))
	
	If Cint("0" & Upload.Form("txtActionType")) = 0 Then
		ACtionType = "0"
	Else
		ACtionType = Upload.Form("txtActionType")
	End If
	
	rsCat.Fields("ActionType").Value = ACtionType
	
	if Upload.Form("txtSSID") <> "" then
		rsCat.Fields("SSID").Value = Upload.Form("txtSSID")
	end if
	
	'Response.Write Upload.Form("txtActionType")
	'Response.End 
	
	
	rsCat.Fields("OTMessage").Value = Upload.Form("txtOTMessage")
	rsCat.Fields("ItinIntroTemplate") = Upload.Form("txtItinIntro")
	rsCat.Fields("LogoBGColor") = left(trim(Upload.Form("txtLogoBGColor")),24)
	
	'rsCat.Fields("HighWayPref").Value = Upload.Form("txtHighWayPref")

	If Upload.Form("chkSameAsHotelAddress") = "on" Then
		rsCat.Fields("SameAsHotel").Value = 1
		rsHL.Fields("useAlternate").Value = 0
		sah = 0
	Else
		rsCat.Fields("SameAsHotel").Value = 0
		rsHL.Fields("useAlternate").Value = 1
		sah = 1
	End If
	
	if Upload.Form("selSearchType") <> "" then
		rsCat.Fields("GPSearchID").Value = Upload.Form("selSearchType")
		remote.Session("GPSearchID") = Upload.Form("selSearchType")
	end if
	
	rsCat.Fields("InfoUSAImport").Value = Upload.Form("cboInfoUSA")
	rsCat.Fields("FaxName").Value = Upload.Form("txtFaxName")
	rsCat.Fields("FaxSetup").Value = Upload.Form("txtFaxSetup")
	
	'Response.write upload.form("chkUseGuestProfile")
	'Response.End
	
	if Upload.Form("chkUseGuestProfile") = "on" then
		rsCat.Fields("UseGuestProfile").Value = 1
	else
		rsCat.Fields("UseGuestProfile").Value = 0
	end if
	remote.Session("UseGuestProfile") = rsCat.Fields("UseGuestProfile").Value
	
	If Upload.Form("chkRoutePref") = "on" Then
		rsCat.Fields("RoutePref").Value = 1
	Else
		rsCat.Fields("RoutePref").Value = 0
	End If
	
	If Upload.Form("chkRollover") = "on" Then
		rsCat.Fields("Rollover").Value = 1
	Else
		rsCat.Fields("Rollover").Value = 0
	End If
	
	If Upload.Form("chkConciergePhone") = "on" Then
		rsCat.Fields("showConciergePhone").Value = 1
	Else
		rsCat.Fields("showConciergePhone").Value = 0
	End If
	
	if len(trim(Upload.Form("txtConciergePhone"))) > 0 then
		rsCat.Fields("ConciergePhone").Value = Upload.Form("txtConciergePhone")
	end if	
	
	'If Upload.Form("chkLatLong") = "on" Then
	'	rsCat.Fields("useLatLong").Value = 1
	'Else
	'	rsCat.Fields("useLatLong").Value = 0
	'End If

	rsCat.Fields("LetterfootMargin").Value = Upload.Form("txtBottomMargin")
''	rsCat.Fields("ScreenLogoLocation").Value = "obsolete"
	rsCat.Fields("LateTaskTime").Value = Upload.Form("txtLateMinutes")
	'rsCat.Fields("GuestTaskReportTemplate").Value = Upload.Form("txtGuestTaskReportTemplate")
	if len(trim(Upload.Form("txtGuestServicesExtension"))) > 0 then
		rsCat.Fields("GuestServicesExtension").Value = Upload.Form("txtGuestServicesExtension")
	end if
	if len(trim(Upload.Form("txtBackupInterval"))) > 0 then
		rsCat.Fields("BackupInterval").Value = Upload.Form("txtBackupInterval")
	end if	
	if len(trim(Upload.Form("txtBackupStart"))) > 0 then
		rsCat.Fields("BackupStart").Value = Upload.Form("txtBackupStart")
	end if
	if len(trim(Upload.Form("txtBackupEnd"))) > 0 then
		rsCat.Fields("BackupEnd").Value = Upload.Form("txtBackupEnd")
	end if
	rsCat.Fields("EmailCopy").Value = Upload.Form("txtUsersCopied")
	rsCat.Fields ("TimeZone").Value = Upload.Form ("cboTimeZone")
	'rsCat.Fields ("MapUNC").Value = Upload.Form ("txtMapUNC")
	
	booBustPage1Cache = false
	if rsCat.Fields("DefaultCategory").Value <> Upload.Form("DefaultCategory") then
		booBustPage1Cache = true
	end if
	rsCat.Fields("DefaultCategory").Value = Upload.Form("DefaultCategory")
	rsCat.Fields("DefaultBCK").Value = Upload.Form("grpBCK")
	rsCat.Fields("DefaultSortBy").Value = Upload.Form("DefaultSortBy")
	rsCat.Fields("DefaultState").Value = Upload.Form("DefaultState")

	if Upload.Form("cmbCalView") = "" then
		dv = null
	else
		dv = Upload.Form("cmbCalView")
	end if
	rsCat.Fields("DefaultCalView").Value = dv
	if strCompanyID = remote.Session("CompanyID") then
		remote.session("DefaultCalView") = dv
	end if
	
	rsCat.Fields("ItinFontName").Value = Upload.Form("selItinFontName")
	rsCat.Fields("ItinFontSize").Value = Upload.Form("selItinFontSize")
	'Response.End
	
	rsCat.Fields("QuickLinkDefault").Value = Upload.Form("grpLinkType")
	if Upload.Form("chkForceQL") = "" then
		intForceQL = 0
	else
		if Upload.Form("chkForceQL") = "on" then
			intForceQL = 1
		else
			intForceQL = 0
		end if
	end if
	rsCat.Fields("QuickLinkForce").Value = intForceQL

	'''''''''''''''''
	if isnull(e2n(Upload.Form("txtAssignDistance"))) then
		n = 100
	else
		n = cint(Upload.Form("txtAssignDistance"))
	end if
	
	rsCat.Fields("AssignDistance").Value = n
	
	rsCat.Fields("CompanyType").Value = Upload.Form("txtCompanyType")
	
	if Upload.Form("chkUseCustomLatLong") = "on" then
		rsCat.Fields("UseCustomLatLong").Value = 1
		rsHL.Fields("UseCustomLatLong").Value = 1
	else
		rsCat.Fields("UseCustomLatLong").Value = 0
		rsHL.Fields("UseCustomLatLong").Value = 0
	end if	
	
	rsCat.Fields("Country").Value = Upload.Form("cboHotelCountry") 
	rsCat.Fields("DirectionsCountry").Value = Upload.Form("cboDirectionsCountry") 
	
	
	' Saving Company Preferences

	Function rsPref (fName)	
			If Upload.Form ("boo" & fName) = "1" Then
				rsPrefs(fName) = 1
			Else
				rsPrefs(fName) = 0
			End If
	End Function
	
	set rsPrefs = Server.CreateObject ("Adodb.Recordset")
	

	rsPrefs.LockType = adLockPessimistic
	rsPrefs.CursorType = adOpenKeyset
	rsPrefs.Open "Select * from tblCompanyPrefs where CompanyID=" & strCompanyID, cnSQL
	

	res = rsPref("CompanyName")
	res = rsPref("Contact")
	res = rsPref("Street")
	res = rsPref("City")
	res = rsPref("Phone")
	res = rsPref("PhoneAlternate")
	res = rsPref("FaxNumber")
	res = rsPref("PagerNumber")
	res = rsPref("Email")
	res = rsPref("HotelRating")
	res = rsPref("CostRating")
	res = rsPref("Recommended")
	res = rsPref("live_music")
	res = rsPref("CrossStreets")
	res = rsPref("lWebsite")
	res = rsPref("PrivateNotes")
	res = rsPref("Notes")
	res = rsPref("Hours")
	res = rsPref("Price")
	res = rsPref("Parking")
	res = rsPref("Teaser")
	res = rsPref("Synopsis")
	res = rsPref("Meal")
	res = rsPref("Amenity")
	res = rsPref("Atmosphere")
	res = rsPref("Payment")
	res = rsPref("Neighborhood")
	res = rsPref("Directions")
	res = rsPref("DirectionsToLocation")
	res = rsPref("DirectionsToHotel")
	res = rsPref("HotelNotes")
	res = rsPref("Transportation")
	res = rsPref("MainMap")
	
	rsPrefs.Update 
	set rsPrefs = Nothing
	

	lid = Upload.Form("txtLocationID")
	if len(trim(lid)) = 0 then
		lid = null
	else
		lid = clng(lid)
	end if
	
	strLon = trim(Upload.Form("txtLon"))
	if len(trim(Upload.Form("txtLat"))) > 0 and len(trim(Upload.Form("txtLon"))) > 0 then ' if supplied
		if left(strLon,1) <> "-" then
			strLon = "-" & strLon
		end if
		rsHL.Fields("CGLatitude").Value = CDBL(Upload.Form("txtLat"))
		rsHL.Fields("CGLongitude").Value = CDBL(strLon)
	end if
	
	rsHL.Fields("Latitude").Value = null
	rsHL.Fields("Longitude").Value = null
	rsHL.Update 
	rsHL.Close 

	Set MapPoint = CreateObject("GCN.Locationinitializer")
	
	Response.Write "<br>" & lid
	ll = MapPoint.FindLatLongByID(lid)
	Response.Write "<br>" & ll
	a = split(ll,"|")
	set MapPoint = nothing
	
	
	rsCat.Update
	rsCat.Close()


	if strCompanyID = remote.Session("CompanyID") then
		remote.Session("LogoBGColor") = trim(Upload.Form("txtLogoBGColor"))
		remote.Session("ZagatURL") = Upload.Form("txtZagatURL")
		remote.Session("FlightsURL") = Upload.Form("txtFlightsURL")
	end if
	
	if booBustPage1Cache then
		url = Application("HomePage") & "/CheckRebuildPage1Cache.asp?cid=" & strCompanyID
		set xmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
		xmlhttp.open "GET", url, true
		xmlhttp.send ""
		set xmlhttp = nothing
	end if
		
	Response.Redirect strBackPage
	
	function e2n( v )
		if len(trim(v)) = 0 then
			e2n = null
		else
			e2n = v
		end if
	end function
%>
<a href=Switchboard3.asp>Back to Home Page</a>
