<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))

if isnumeric(Request.Form("txtGID")) then
	intGID = Request.Form("txtGID")
else
	intGID = 0
end if

a = split(Request.Form,"&")

strTTIDs  = "" 'Guest Phone ID's
strATIDs  = "" 'Address Type ID's
strRPIDs  = "" 'Rewards Program ID's
strCCTIDs = "" 'Credit Card Type ID's
strFMIDs  = "" 'Family Member ID's
strIDIDs  = "" 'Important Dates ID's
strPIDs   = "" 'Preferences ID's
strHIDs   = "" 'HotelIDs

for i = 0 to ubound(a)
	if left(a(i),3) <> "pm_" then ' skip phone mask fields
		'Response.write a(i) & "<br>"

		' get id's
		if instr(1,a(i),"selGPPhoneType") > 0 then
			pos1 = instrrev(a(i),"_")+1
			strTTIDs = strTTIDs & mid(a(i),pos1,instrrev(a(i),"=")-pos1) & ","
		end if
		if instr(1,a(i),"selATAddressType") > 0 then
			pos1 = instrrev(a(i),"_")+1
			strATIDs = strATIDs & mid(a(i),pos1,instrrev(a(i),"=")-pos1) & ","
		end if
		if instr(1,a(i),"selGRRewardsType") > 0 then
			pos1 = instrrev(a(i),"_")+1
			strRPIDs = strRPIDs & mid(a(i),pos1,instrrev(a(i),"=")-pos1) & ","
		end if
		if instr(1,a(i),"selCCCCType") > 0 then
			pos1 = instrrev(a(i),"_")+1
			strCCTIDs = strCCTIDs & mid(a(i),pos1,instrrev(a(i),"=")-pos1) & ","
		end if
		if instr(1,a(i),"selFMType") > 0 then
			pos1 = instrrev(a(i),"_")+1
			strFMIDs = strFMIDs & mid(a(i),pos1,instrrev(a(i),"=")-pos1) & ","
		end if
		if instr(1,a(i),"selIDType") > 0 then
			pos1 = instrrev(a(i),"_")+1
			strIDIDs = strIDIDs & mid(a(i),pos1,instrrev(a(i),"=")-pos1) & ","
		end if
		if instr(1,a(i),"selPrefType") > 0 then
			pos1 = instrrev(a(i),"_")+1
			strPIDs = strPIDs & mid(a(i),pos1,instrrev(a(i),"=")-pos1) & ","
		end if

	end if
next

function formatIDs( s )
	dim retVal
	retVal = ""
	if trim(s) <> "" then
		'Response.Write "[" & s & "]"
		retVal = left(s,len(s)-1)
	end if
	formatIDs = retVal
end function

strTTIDs  = formatIDs(strTTIDs)
strATIDs  = formatIDs(strATIDs)
strRPIDs  = formatIDs(strRPIDs)
strCCTIDs = formatIDs(strCCTIDs)
strFMIDs  = formatIDs(strFMIDs)
strIDIDs  = formatIDs(strIDIDs)
strPIDs   = formatIDs(strPIDs)

'Response.Write "<hr>"
'Response.Write "Telephone Type ID's: " & strTTIDs & "<br>"
'Response.Write "Address Type ID's: " & strATIDs & "<br>"
'Response.Write "Rewards Program ID's: " & strRPIDs & "<br>"
'Response.Write "CC ID's: " & strCCTIDs & "<br>"
'Response.Write "Family Member ID's: " & strFMIDs & "<br>"
'Response.Write "Important Date ID's: " & strIDIDs & "<br>"
'Response.Write "Preference ID's: " & strPIDs & "<br>"

'Response.Write "<hr>"

if Request.Form("chkSmoking") = "on" then
	booSmoking = 1
else
	booSmoking = 0
end if

dim cn, cm
set cn = server.CreateObject("adodb.connection")
set cm = server.CreateObject("adodb.command")

cn.Open Application("sqlInnSight_ConnectionString")
cm.ActiveConnection = cn
cm.CommandType = 4
cm.CommandText = "sp_GuestProfileSave"

cm.Parameters.append(cm.CreateParameter("@GID",adInteger,adParamInput,,intGID))
cm.Parameters.append(cm.CreateParameter("@Salutation",adVarChar,adParamInput,50,Request.Form("ddSalutation")))
cm.Parameters.append(cm.CreateParameter("@LastName",adVarChar,adParamInput,50,Request.Form("txtLastName")))
cm.Parameters.append(cm.CreateParameter("@MiddleName",adVarChar,adParamInput,50,Request.Form("txtMiddleName")))
cm.Parameters.append(cm.CreateParameter("@FirstName",adVarChar,adParamInput,50,Request.Form("txtFirstName")))
cm.Parameters.append(cm.CreateParameter("@Company",adVarChar,adParamInput,120,Request.Form("txtCompany")))
cm.Parameters.append(cm.CreateParameter("@Title",adVarChar,adParamInput,120,Request.Form("txtTitle")))
cm.Parameters.append(cm.CreateParameter("@PrimaryPhone",adChar,adParamInput,24,Request.Form("txtPrimaryPhone")))
cm.Parameters.append(cm.CreateParameter("@PhoneExt",adVarChar,adParamInput,12,Request.Form("txtPhoneExt")))
cm.Parameters.append(cm.CreateParameter("@EMail1",adVarChar,adParamInput,80,Request.Form("txtEmail")))
cm.Parameters.append(cm.CreateParameter("@EMail2",adVarChar,adParamInput,80,Request.Form("txtEmail2")))
cm.Parameters.append(cm.CreateParameter("@Smoking",adBoolean,adParamInput,,booSmoking))
cm.Parameters.append(cm.CreateParameter("@PMSGuestID",adChar,adParamInput,24,Request.Form("txtPMSGuestID")))
cm.Parameters.append(cm.CreateParameter("@HotelGuestID",adChar,adParamInput,24,Request.Form("txtHotelGuestID")))
cm.Parameters.append(cm.CreateParameter("@GuestNotes",adVarChar,adParamInput,1024,Request.Form("txtGuestNote")))
cm.Parameters.append(cm.CreateParameter("@NewGID",adInteger,adParamOutput))

cm.Execute

if intGID > 0 then
	GID = intGID
	'Response.Write intGID & " was used."
else
	GID = cm.Parameters("@NewGID").Value
	'Response.Write cm.Parameters("@NewGID").Value & " was created."
end if

aTT  = split(strTTIDs,",")
aAT  = split(strATIDs,",")
aRP  = split(strRPIDs,",")
aCCT = split(strCCTIDs,",")
aFM  = split(strFMIDs,",")
aID  = split(strIDIDs,",")
aP   = split(strPIDs,",")

strSQL = "begin transaction;"
strSQL = "delete tlnkGuestPhone where GuestID = " & gid & ";"
for i = lbound(aTT) to ubound(aTT)
	if Request.Form("txtGPPhoneNumber" & aTT(i)) <> "" then
		strSQL = strSQL & "insert tlnkGuestPhone (GuestID, PhoneTypeID, PhoneNumber, PhoneExt, PhonePrimary, PhoneNote) values (" & gid & ", " & Request.Form("selGPPhoneType_" & aTT(i)) & ", '" & Request.Form("txtGPPhoneNumber" & aTT(i)) & "', '" & Request.Form("txtGPExt_" & aTT(i)) & "', " & cchk(Request.Form("chkGPPrimary_" & aTT(i))) & ", '" & replace(Request.Form("txtGuestNote_" & aTT(i)),"'","''") & "');"
	end if
next
strSQL = strSQL & "delete tlnkGuestAddress where GuestID = " & gid & ";"
for i = lbound(aAT) to ubound(aAT)
	if Request.Form("txtGAStreet_" & aAT(i)) <> "" then
		strSQL = strSQL & "insert tlnkGuestAddress (GuestID, AddressTypeID, Address, Suite, City, State, Zip, AddressPrimary, Note) values (" & gid & ", " & Request.Form("selATAddressType_" & aAT(i)) & ", '" & sq(Request.Form("txtGAStreet_" & aAT(i))) & "', '" & sq(Request.Form("txtGASuite_" & aAT(i))) & "', '" & sq(Request.Form("txtGACity_" & aAT(i))) & "', '" & sq(Request.Form("selGAState_" & aAT(i))) & "', '" & sq(Request.Form("txtGAZip_" & aAT(i))) & "', " & cchk(Request.Form("chkAddressPrimary_" & aAT(i))) & ", '" & sq(Request.Form("txtAddrNote_" & aAT(i))) & "');"
	end if
next
strSQL = strSQL & "delete tlnkGuestRewards where GuestID = " & gid & ";"
for i = lbound(aRP) to ubound(aRP)
	if Request.Form("txtGRProgNum_" & aRP(i)) <> "" then
		strSQL = strSQL & "insert tlnkGuestRewards (GuestID, RewardsTypeID, ProgramID, ProgramNumber, ProgramLevel, Note) values (" & gid & ", " & Request.Form("selGRRewardsType_" & aRP(i)) & ", " & Request.Form("selGRProgramName_" & aRP(i)) & ", '" & sq(Request.Form("txtGRProgNum_" & aRP(i))) & "', '" & sq(Request.Form("txtGRProgLevel_" & aRP(i))) & "', '" & sq(Request.Form("txtGRNote_" & aRP(i))) & "');"
	end if
next
strSQL = strSQL & "delete tlnkGuestCharge where GuestID = " & gid & ";"
for i = lbound(aCCT) to ubound(aCCT)
	if Request.Form("txtGCNumber_" & aCCT(i)) <> "" or Request.Form("selCCCCType_" & aCCT(i)) = remote.session("CashID") then
		strSQL = strSQL & "insert tlnkGuestCharge (GuestID, ChargeTypeID, ChargeNumber, Expiration, ZipCode, ChargePrimary, Note) values (" & gid & ", " & Request.Form("selCCCCType_" & aCCT(i)) & ", '" & sq(Request.Form("txtGCNumber_" & aCCT(i))) & "', '" & sq(Request.Form("txtGCExp_" & aCCT(i))) & "', '" & sq(Request.Form("txtGCZip_" & aCCT(i))) & "', '" & cchk(Request.Form("chkGCPrimary_" & aCCT(i))) & "', '" & sq(Request.Form("txtGCNote_" & aCCT(i))) & "');"
	end if
next
strSQL = strSQL & "delete tlnkGuestFamily where GuestID = " & gid & ";"
for i = lbound(aFM) to ubound(aFM)
	if Request.Form("txtFMFirstName_" & aFM(i)) <> "" then
		strSQL = strSQL & "insert tlnkGuestFamily (GuestID, RelationshipID, Salutation, FirstName, LastName, Birthdate, Age, Note) values (" & gid & ", " & Request.Form("selFMType_" & aFM(i)) & ", '" & sq(Request.Form("selFMSal_" & aFM(i))) & "', '" & sq(Request.Form("txtFMFirstName_" & aFM(i))) & "', '" & sq(Request.Form("txtFMLastName_" & aFM(i))) & "', '" & sq(Request.Form("txtFMDOB_" & aFM(i))) & "', '" & sq(Request.Form("txtFMAge_" & aFM(i))) & "', '" & sq(Request.Form("txtFMNote_" & aFM(i))) & "');"
	end if
next
strSQL = strSQL & "delete tlnkGuestImportantDate where GuestID = " & gid & ";"
for i = lbound(aID) to ubound(aID)
	if Request.Form("txtIDDate_" & aID(i)) <> "" then
		strSQL = strSQL & "insert tlnkGuestImportantDate (GuestID, DateTypeID, DateTypeDate, FirstName, LastName, RelationshipID, Note) values (" & gid & ", " & Request.Form("selIDType_" & aID(i)) & ", '" & sq(Request.Form("txtIDDate_" & aID(i))) & "', '" & sq(Request.Form("txtIDFirstName_" & aID(i))) & "', '" & sq(Request.Form("txtIDLastName_" & aID(i))) & "', " & Request.Form("selIDRelationship_" & aID(i)) & ", '" & sq(Request.Form("txtIDNote_" & aID(i))) & "');"
	end if
next
strSQL = strSQL & "delete tlnkGuestPreference where GuestID = " & gid & ";"
for i = lbound(aP) to ubound(aP)
	if Request.Form("txtGPrefNote_" & aP(i)) <> "" then
		strSQL = strSQL & "insert tlnkGuestPreference (GuestID, PreferenceID, Note) values (" & gid & ", " & Request.Form("selPrefType_" & aP(i)) & ", '" & sq(Request.Form("txtGPrefNote_" & aP(i))) & "');"
	end if
next


if Request.Form("txtAssignedHotels") = "" then
	strSQL = strSQL & "insert tlnkGuestHotel (GuestID, HotelID) values (" & gid & ", " & Remote.Session("CompanyID") & ");"
else
	aHotels = split(unescape(Request.Form("txtAssignedHotels")),",")
	strSQL = strSQL & "delete tlnkGuestHotel where GuestID = " & gid & ";"
	for i = lbound(aHotels) to ubound(aHotels)
		strSQL = strSQL & "insert tlnkGuestHotel (GuestID, HotelID) values (" & gid & ", " & aHotels(i) & ");"
	next
end if

strSQL = strSQL & "commit transaction;"
cn.Execute strSQL




function cchk( s )
	if s = "on" then
		cchk = 1
	else
		cchk = 0
	end if
end function

function sq( s )
	sq = replace(s,"'","''")
end function

cn.Close
set cn = nothing
%>

