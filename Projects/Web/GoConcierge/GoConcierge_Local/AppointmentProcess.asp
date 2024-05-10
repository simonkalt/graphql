<%
Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))
cid = remote.Session("CompanyID")
ddid = Request.Form("txtDDID")

' if for some reason (client error) departments were not passed, 
' default to Concierge department...
if ddid = "" then
	Set cnSQL = Server.CreateObject("ADODB.Connection")
	cnSQL.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")
	set rs = server.CreateObject("adodb.recordset")
	set rs = cnSQL.execute("select DepartmentID from tblDepartment where DepartmentName = 'Concierge'")
	ddid = cstr(rs(0).Value)
	rs.Close
	set rs = nothing
	cnSQL.close
	set cnSQL = nothing
end if

	'if Request.QueryString("CalledFrom") = "IDE" then
	'	Response.Write "<script language=javascript>" & vbcrlf
	'	Response.Write "parent.window.opener.location.replace(parent.window.opener.location.href);" & vbcrlf
	'	Response.Write "parent.window.opener.status = parent.window.opener.location" & vbcrlf
	'	'Response.Write "parent.window.opener.winTemp.location.replace(parent.window.opener.winTemp.location.href);" & vbcrlf
	'	Response.Write "</script>"
	'end if

	dim lngLocationID ', intPeople
	dim strStartDate, strEndDate, strStartTime, strEndTime
	
	Response.Buffer = True
	
	strLocationText = Request.Form("txtLocation")
	
	d1 = trim(Request.Form("d1"))
	if instr(1,d1," ") > 0 then
		strStartDate = right(d1,len(trim(d1))-4) 'Remove day string
	else
		strStartDate = d1
	end if
	strEndDate = strStartDate 'Request.Form("pvDateEnd")
	strStartTime = Request.Form("txtStartTime")
	strEndTime  = Request.Form("txtEndTime")

	'Response.Write "txtLocationID: " & Request.Form("txtLocationID")
	'Response.End
	
	if len(trim(Request.Form("txtLocationID"))) = 0 then
		lngLocationID = 0
	else
		lngLocationID = Request.Form("txtLocationID")
	end if

	if Request.Form("chkEMail") = "on" then
		bitEMail = 1
	else
		bitEMail = 0
	end if
	
	' Change stop concat if spanning multi-days in future
	dtStart = cDate(strStartDate & " " & strStartTime)
	dtStop = cDate(strStartDate & " " & strEndTime)
	if trim(Request.Form("txtDateAdded")) = "" then
		dtDateAdded = date()
	else
		dtDateAdded = cDate(Request.Form("txtDateAdded"))
	end if
	'dtDateAdded = cDate(Request.Form("txtDateAdded"))
	
	If IsNumeric(Request.Form("txtAppointmentID")) Then
	
		Set cnSQL = Server.CreateObject("ADODB.Connection")
		Set rsLoc = Server.CreateObject("ADODB.Recordset")
		Set rsAppt = Server.CreateObject("ADODB.Recordset")
  
		cnSQL.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")

		'First Update the Appointment Table
		booAddItAnyway = false
		if Request.Form("txtAppointmentID") <> 0 then
			rsAppt.Open "Select AppointmentID from tblAppointment Where AppointmentID=" & Request.Form("txtAppointmentID"),cnSQL,adOpenKeyset,adLockPessimistic
			if rsAppt.EOF then 'deleted from under user
				booAddItAnyway = true
			end if
			rsAppt.Close
		end if
		
		If Request.Form("txtAppointmentID") = 0 or booAddItAnyway Then
			intCurrentUserID = Request.Form("txtCreateUserID")
			
			dim cmm
			set cmm = server.CreateObject("adodb.command")
			cmm.ActiveConnection = cnSQL
			cmm.CommandType = adCmdStoredProc
			cmm.CommandText = "sp_AddAppointment"
			Response.Write "CID: " & cid & "<br>"
			cmm.Parameters.append cmm.CreateParameter("@cid",adInteger,adParamInput,,cint(cid))
			cmm.Parameters.append cmm.CreateParameter("@aid",adInteger,adParamOutput)
			cmm.Execute()
			aid = cmm.Parameters("@aid").Value
			set cmm = nothing
			
			rsAppt.Open "select * from tblAppointment where appointmentID = " & aid, cnSQL, adOpenKeyset, adLockPessimistic
			
			'rsAppt.AddNew
				
				rsAppt.Fields("fkCompanyID").Value = cid
				rsAppt.Fields("ApptStartDate").Value = dtStart
				rsAppt.Fields("ApptEndDate").Value = dtStop
				rsAppt.Fields("DateAdded").Value = dtDateAdded
				
				if isnumeric(Request.Form("txtRealGuestID")) then
					rsAppt.Fields("GuestID").Value = Request.Form("txtRealGuestID")
				end if
				rsAppt.Fields("DisplayID").Value = Request.Form("txtGuestID")
				
				rsAppt.Fields("ApptText").Value = RemoveTrailingCRLFS( left(Request.Form("txtSubjectSave"),6442) )
				rsAppt.Fields("Alarm").Value = 0
				rsAppt.Fields("UserID").Value = 1
				rsAppt.Fields("Notes").Value = "Notes"
				rsAppt.Fields("CreateDateTime").Value = Now() + remote.Session("TimeZone")
				rsAppt.Fields("CreateUserID").Value = Request.Form("txtCreateUserID") 'remote.Session("UserID")
				rsAppt.Fields("EditDateTime").Value = Null
				rsAppt.Fields("EditUserID").Value = Null
				rsAppt.Fields("Room").Value = left(Request.Form("txtRoom"),20)
				rsAppt.Fields("Salutation").Value = unescape(Request.Form("txtSalutation"))
				rsAppt.Fields("GuestFirstName").Value = left(Request.Form("txtGuestFirstName"),35)
				rsAppt.Fields("GuestLastName").Value = left(Request.Form("txtGuestLastName"),45)

				rsAppt.Fields("GuestPhone").Value = left(Request.Form("txtGuestPhone"),32)
				rsAppt.Fields("GuestEMail").Value = left(Request.Form("txtGuestEMail"),50)

				rsAppt.Fields("ActionType").Value = 0
				'rsAppt.Fields("People").Value = intPeople
				rsAppt.Fields("EmailConfirm").Value = bitEMail
				
				'If Trim(Request.Form("chkClosed")) <> "on" Then
				'	rsAppt.Fields("Closed").Value = 0
				'Else
				'	rsAppt.Fields("Closed").Value = 1
				'End If
				
				rsAppt.Fields("Status").Value = Request.Form("cboStatus")

				If Request.Form("txtStateChangedToClosed") = "true" then
					rsAppt.Fields("ClosedUserID").Value = Request.Form("txtCreateUserID") 'remote.Session("UserID")
					rsAppt.Fields("ClosedDate").Value = now() + remote.Session("TimeZone")
				else 'if Request.Form("txtStateChangedToClosed") = "false" then
					rsAppt.Fields("ClosedUserID").Value = null
					rsAppt.Fields("ClosedDate").Value = null
				end if

				rsAppt.Fields("CloseNote").Value = Request.Form("txtNotes")

				'Response.AppendToLog "<!-- Debug: 1: " & len(trim(Request.Form("txtRoom"))) = 0 & _
				'				" - 2: " & len(trim(Request.Form("txtSalutation"))) = 0 & _
				'				" - 3: " & len(trim(Request.Form("cboActionType"))) = 0 & _
				'				" - 4: " & len(trim(Request.Form("cboAction"))) = 0 & _
				'				" - 5: " & len(trim(strLocationText)) = 0 & _
				'				" - 6: " & len(trim(Request.Form("LocCity"))) = 0 & _
				'				" - 7: " & len(trim(Request.Form("LocState"))) = 0 & _
				'				" - 8: " & len(trim(Request.Form("LocZip"))) = 0 & _
				'				" - 9: " & len(trim(Request.Form("LocPhone"))) = 0 & _
				'				" - 10: " & len(trim(Request.Form("LocAddress"))) = 0 & _
				'				" -->"
				'Response.End								
				
				if Request.Form("chkRollover") = "on" then
					rsAppt.Fields("Rollover").Value = 1
				else
					rsAppt.Fields("Rollover").Value = 0
				end if
				
				if Request.Form("chkSpan") = "on" then
					rsAppt.Fields("Span").Value = 1
				else
					rsAppt.Fields("Span").Value = 0
				end if
				
								
				if	len(trim(Request.Form("txtRoom"))) = 0 and _
					len(trim(Request.Form("txtSalutation"))) = 0 and _
					Request.Form("cboActionType") = "0" and _
					Request.Form("cboAction") = "0" and _
					len(trim(strLocationText)) = 0 then 'and _
					'len(trim(Request.Form("LocCity"))) = 0 'and _
					'len(trim(Request.Form("LocState"))) = 0 and _
					'len(trim(Request.Form("LocZip"))) = 0 and _
					'len(trim(Request.Form("LocPhone"))) = 0 and _
					'len(trim(Request.Form("LocAddress"))) = 0 then

						rsAppt.Fields("Note").Value = 1
				else
					If Request.Form("chkNoteHidden") = "on" Then
						rsAppt.Fields("Note").Value = 1
					Else
						rsAppt.Fields("Note").Value = 0
					End If
				end if

				' Timeless note
				If Request.Form("chkNoTime") = "on" Then
					rsAppt.Fields("NoTime").Value = 1
				Else
					rsAppt.Fields("NoTime").Value = 0
				End If

				If Request.Form("chkAllDayEvent") = "on" Then
					rsAppt.Fields("AllDayEvent").Value = 1
				Else
					rsAppt.Fields("AllDayEvent").Value = 0
				End If
				rsAppt.Fields("ActionTypeID").Value = Request.Form("cboActionType")
				rsAppt.Fields("ActionID").Value = Request.Form("cboAction")
				rsAppt.Fields("SearchKey").Value = "SearchKey"
				rsAppt.Fields("CompanyName").Value = "CompanyName"
				'Obsolete because users can type in text.
				rsAppt.Fields("LocationID").Value = lngLocationID 'foreign key into tblLocation
				rsAppt.Fields("LocationText").Value = strLocationText
				rsAppt.Fields("Recurrence").Value = 0 'Request.Form("txtRecurrence")
				
				' Added by Ilia 09/24/2001 - Takes care of new Location related Fields
				
				rsAppt.Fields("LocCity").Value = Trim(Request.Form("LocCity"))
				rsAppt.Fields("LocState").Value = Trim(Request.Form("LocState"))
				rsAppt.Fields("LocZip").Value = Trim(Request.Form("LocZIP"))
				rsAppt.Fields("LocPhone").Value = Trim(Request.Form("txtLocPhone"))


				rsAppt.Fields("LocAddress").Value = Trim(Request.Form("Address"))
				rsAppt.Fields("LocSpokeWith").Value = Trim(Request.Form("SpokeWith"))
				rsAppt.Fields("LocConfirm").Value = Request.Form("Confirm")
				'rsAppt.Fields("ChargeType").Value = Request.Form("cboCard")
				
				set rsLoc = cnSQL.Execute("select CrossStreets from vwLocations where LocationID = " & lngLocationID)
				if not rsLoc.EOF then
					rsAppt.Fields("LocXStreet").Value = rsLoc.Fields("CrossStreets").Value
				end if
				rsLoc.Close
				set rsLoc = nothing
				
				if instr(1,Request.Form("CCNumber"),"xxx") < 1 then
					rsAppt.Fields("CCNumber").Value = Request.Form("CCNumber")
				end if
				
				rsAppt.Fields("CCType").Value = Request.Form("cboChargeTo")
				
				'If Request.Form("CCExp") <> "" Then
				rsAppt.Fields("CCExp").Value = Request.Form("CCExp")
				'End If  
				
				If Request.Form("txtAmount") = "" or isEmpty(Request.Form("txtAmount")) Then
					n = null
				else
					n = ccur(Request.Form("txtAmount"))
				end if
				rsAppt.Fields("Amount").Value = n
				
				rsAppt.Fields("GuestMiddleName").Value = Request.Form("txtMI")
				rsAppt.Fields("FGN").Value = Request.Form("txtFGN")
				rsAppt.Fields("GuestNum").Value = Request.Form("txtGuestNum")
				rsAppt.Fields("ResNum").Value = Request.Form("txtResNum")
				rsAppt.Fields("ArrivalDate").Value = Request.Form("txtArrivalDate")
				rsAppt.Fields("DepartDate").Value = Request.Form("txtDepartDate")
				rsAppt.Fields("OTLogID").Value = Request.Form("OTLogID")
				rsAppt.Fields("SSLogID").Value = Request.Form("SSLogID")
				
				If Request.Form("NonGuest") = "on" Then
					rsAppt.Fields("GuestYesNo").Value = 1
				Else
					rsAppt.Fields("GuestYesNo").Value = 0
				End If
				
				'if Request.Form("chkAllDepartments") = "on" then
				'	rsAppt.Fields("DepartmentID").Value = 0
				'else
				'	rsAppt.Fields("DepartmentID").Value = remote.session("FloatingUser_DDID")
				'end if
				
			'rsAppt.Fields("DepartmentID").Value = Request.Form("cmbDepartment")
				
			'Set cmdGN = Server.CreateObject("ADODB.Command")
			'Set cmdGN.ActiveConnection = cnSQL
			'cmdGN.CommandText = "SELECT @@Identity"
			'cmdGN.CommandType = adCmdText

			'rsAppt.Open cmdGN, ,0,2
			'intID = rsAppt.Fields(0).Value
			'rsAppt.Close
			
			'response.write "<BR>intID: " & intID & " or " & s & "<BR>"
			'Response.End
			
			'booFixed = false
			'if lngLocationID = 0 and len(trim(strLocationText)) > 0 then
			'	'check for existence of location workaround
			'	set temprs = server.CreateObject("adodb.recordset")
			'	if len(trim(Request.Form("LocPhone"))) > 0 and len(trim(Request.Form("Address"))) > 0 then
			'		temprs.Open "select locationid from tblLocation where companyname='" & replace(strLocationText,"'","''") & "' and phone='" & Request.Form("LocPhone") & "' and Street = '" & Request.Form("Address") & "'", cnSQL
			'		if not temprs.EOF then
			'			rsAppt.Fields("LocationID").Value = temprs("LocationID").Value
			'			booFixed = true
			'		end if
			'		temprs.Close
			'	else
			'		if len(trim(Request.Form("LocPhone"))) > 0 then
			'			temprs.Open "select locationid from tblLocation where companyname='" & replace(strLocationText,"'","''") & "' and phone='" & Request.Form("LocPhone") & "'", cnSQL
			'			if not temprs.EOF then
			'				rsAppt.Fields("LocationID").Value = temprs("LocationID").Value
			'				booFixed = true
			'			end if
			'			temprs.Close
			'		end if
			'	end if
			'	set temprs = nothing
			'
			'	rsAppt.Update
			'	'send us an e-mail stating this stuff
			'	if booFixed = false then
			'		emailGCN(cnSQL)
			'	end if
			'else
				rsAppt.Update
			'end if

			'Get the Appointment ID
			intID = rsAppt.Fields("AppointmentID")
		Else
			' Edit
			
			'Need the ID for adding/updating recurrence info.
			intID = Request.Form("txtAppointmentID")
			intCurrentUserID = remote.Session("LastEditedID")
			
			'Response.Write Request.Form("txtSubjectSave")
			'Response.End

			rsAppt.Open "Select * from tblAppointment Where AppointmentID=" & Request.Form("txtAppointmentID"),cnSQL,adOpenKeyset,adLockPessimistic

				rsAppt.Fields("fkCompanyID").Value = cid
				if rsAppt.Fields("ApptStartDate").Value <> dtStart then
					rsAppt.Fields("DateAdded").Value = dtStart
				end if

				if isnumeric(Request.Form("txtRealGuestID")) then
					rsAppt.Fields("GuestID").Value = Request.Form("txtRealGuestID")
				end if
				rsAppt.Fields("DisplayID").Value = Request.Form("txtGuestID")
				
				rsAppt.Fields("ApptStartDate").Value = dtStart
				rsAppt.Fields("ApptEndDate").Value = dtStop

				rsAppt.Fields("ApptText").Value = RemoveTrailingCRLFs( Left(Request.Form("txtSubjectSave"),6442) )
				rsAppt.Fields("Alarm").Value = 0
				rsAppt.Fields("UserID").Value = 1
				rsAppt.Fields("Notes").Value = "Notes"
				if Request.QueryString("CopyTask") <> "True" then
					rsAppt.Fields("EditDateTime").Value = Now() + remote.Session("TimeZone")
					rsAppt.Fields("EditUserID").Value = remote.Session("LastEditedID") 'remote.Session("UserID")
				end if
				rsAppt.Fields("Room").Value = left(trim(Request.Form("txtRoom")),20)
				rsAppt.Fields("Salutation").Value = unescape(Request.Form("txtSalutation"))
				rsAppt.Fields("GuestFirstName").Value = left(Request.Form("txtGuestFirstName"),35)
				rsAppt.Fields("GuestLastName").Value = left(Request.Form("txtGuestLastName"),45)

				rsAppt.Fields("GuestPhone").Value = left(Request.Form("txtGuestPhone"),32)
				rsAppt.Fields("GuestEMail").Value = left(Request.Form("txtGuestEMail"),50)

				rsAppt.Fields("ActionType").Value = 0
				'rsAppt.Fields("People").Value = intPeople
				rsAppt.Fields("EmailConfirm").Value = bitEmail
				
				If Trim(Request.Form("chkLetterPrinted")) <> "on" Then
					rsAppt.Fields("LetterPrinted").Value = 0
				Else
					rsAppt.Fields("LetterPrinted").Value = 1
				End If

				rsAppt.Fields("Status").Value = Request.Form("cboStatus")
				rsAppt.Fields("CloseNote").Value = Request.Form("txtNotes")

				if Request.Form("chkRollover") = "on" then
					rsAppt.Fields("Rollover").Value = 1
				else
					rsAppt.Fields("Rollover").Value = 0
				end if

				if Request.Form("chkSpan") = "on" then
					rsAppt.Fields("Span").Value = 1
				else
					rsAppt.Fields("Span").Value = 0
				end if
				

				if	len(trim(Request.Form("txtRoom"))) = 0 and _
					len(trim(Request.Form("txtSalutation"))) = 0 and _
					Request.Form("cboActionType") = "0" and _
					Request.Form("cboAction") = "0" and _
					len(trim(strLocationText)) = 0 then 'and _
					'len(trim(Request.Form("LocCity"))) = 0 and _
					'len(trim(Request.Form("LocState"))) = 0 and _
					'len(trim(Request.Form("LocZip"))) = 0 and _
					'len(trim(Request.Form("LocPhone"))) = 0 and _
					'len(trim(Request.Form("LocAddress"))) = 0 then

						rsAppt.Fields("Note").Value = 1
				else
					If Request.Form("chkNoteHidden") = "on" Then
						rsAppt.Fields("Note").Value = 1
					Else
						rsAppt.Fields("Note").Value = 0
					End If
				end if
				
				' Timeless note
				If Request.Form("chkNoTime") = "on" Then
					rsAppt.Fields("NoTime").Value = 1
				Else
					rsAppt.Fields("NoTime").Value = 0
				End If
 
				If Request.Form("chkAllDayEvent") = "on" Then
					rsAppt.Fields("AllDayEvent").Value = 1
				Else
					rsAppt.Fields("AllDayEvent").Value = 0
				End If
				rsAppt.Fields("ActionTypeID").Value = Request.Form("cboActionType")
				'Response.Write "cboAction: " & Request.Form("cboAction")
				rsAppt.Fields("ActionID").Value = Request.Form("cboAction")
				rsAppt.Fields("SearchKey").Value = "SearchKey"
				rsAppt.Fields("CompanyName").Value = "CompanyName"
				
				rsAppt.Fields("LocationID").Value = lngLocationID 'foreign key into tblLocation
				oldLocation = rsAppt.Fields("LocationText").Value
				rsAppt.Fields("LocationText").Value = strLocationText
				' Added by Ilia 09/24/2001 - Takes care of new Location related Fields
				
				rsAppt.Fields("LocSpokeWith").Value = Trim(Request.Form("LocSpokeWith"))
				rsAppt.Fields("LocCity").Value = Trim(Request.Form("LocCity"))
				rsAppt.Fields("LocState").Value = Trim(Request.Form("LocState"))
				rsAppt.Fields("LocZip").Value = Trim(Request.Form("LocZIP"))
				rsAppt.Fields("LocPhone").Value = Trim(Request.Form("txtLocPhone"))
				rsAppt.Fields("LocAddress").Value = Trim(Request.Form("Address"))
				rsAppt.Fields("LocSpokeWith").Value = Trim(Request.Form("SpokeWith"))
				rsAppt.Fields("LocConfirm").Value = Request.Form("Confirm")
				
				set rsLoc = cnSQL.Execute("select CrossStreets from vwLocations where LocationID = " & lngLocationID)
				if not rsLoc.EOF then
					rsAppt.Fields("LocXStreet").Value = rsLoc.Fields("CrossStreets").Value
				end if
				rsLoc.Close
				set rsLoc = nothing
 
				rsAppt.Fields("ChargeType").Value = Request.Form("cboCard")
				
				if instr(1,Request.Form("CCNumber"),"xxx") < 1 then
					rsAppt.Fields("CCNumber").Value = Request.Form("CCNumber")
				end if
				
				rsAppt.Fields("CCType").Value = Request.Form("cboChargeTo")
				
				'If Request.Form("CCExp") <> "" Then
				rsAppt.Fields("CCExp").Value = Request.Form("CCExp")
				'End If  
				
				If Request.Form("txtAmount") = "" or isEmpty(Request.Form("txtAmount")) Then
					n = null
				else
					n = ccur(Request.Form("txtAmount"))
				end if
				rsAppt.Fields("Amount").Value = n
				
				rsAppt.Fields("GuestMiddleName").Value = Request.Form("txtMI")
				rsAppt.Fields("FGN").Value = Request.Form("txtFGN")
				rsAppt.Fields("GuestNum").Value = Request.Form("txtGuestNum")
				rsAppt.Fields("ResNum").Value = Request.Form("txtResNum")
				rsAppt.Fields("ArrivalDate").Value = Request.Form("txtArrivalDate")
				rsAppt.Fields("DepartDate").Value = Request.Form("txtDepartDate")
				 
				if Request.Form("OTLogID") = "" then
					intOTLID = 0
				else
					intOTLID = Request.Form("OTLogID")
				end if
				rsAppt.Fields("OTLogID").Value = intOTLID
				
				if Request.Form("SSLogID") = "" then
					intSSLID = 0
				else
					intSSLID = Request.Form("SSLogID")
				end if
				rsAppt.Fields("SSLogID").Value = intSSLID
				
				
				If Request.Form("NonGuest") = "on" Then
					rsAppt.Fields("GuestYesNo").Value = 1
				Else
					rsAppt.Fields("GuestYesNo").Value = 0
				End If
				
				rsAppt.Fields("Recurrence").Value = 0 
				
				' IF Rec Edit mode = 1 then update only this Task and set it's exception flag
				If Trim(Request.Form("RecID"))<>"" Then
					If Trim(Request.Form("RecEdit")) = "1" Then
						rsAppt.Fields("RecException").Value=1
					End IF
				End IF
				

			'if Request.Form("chkAllDepartments") = "on" then
			'	rsAppt.Fields("DepartmentID").Value = 0
			'else
			'	rsAppt.Fields("DepartmentID").Value = remote.session("FloatingUser_DDID")
			'end if
			'rsAppt.Fields("DepartmentID").Value = Request.Form("cmbDepartment")
			
			
			'booFixed = false
			'if lngLocationID = 0 and len(trim(strLocationText)) > 0 then
			'	'check for existence of location workaround
			'	set temprs = server.CreateObject("adodb.recordset")
			'	if len(trim(Request.Form("LocPhone"))) > 0 and len(trim(Request.Form("Address"))) > 0 then
			'		temprs.Open "select locationid from tblLocation where companyname='" & replace(strLocationText,"'","''") & "' and phone='" & Request.Form("LocPhone") & "' and Street = '" & Request.Form("Address") & "'", cnSQL
			'		if not temprs.EOF then
			'			rsAppt.Fields("LocationID").Value = temprs("LocationID").Value
			'			booFixed = true
			'		end if
			'		temprs.Close
			'	else
			'		if len(trim(Request.Form("LocPhone"))) > 0 then
			'			temprs.Open "select locationid from tblLocation where companyname='" & replace(strLocationText,"'","''") & "' and phone='" & Request.Form("LocPhone") & "'", cnSQL
			'			if not temprs.EOF then
			'				rsAppt.Fields("LocationID").Value = temprs("LocationID").Value
			'				booFixed = true
			'			end if
			'			temprs.Close
			'		end if
			'	end if
			'	set temprs = nothing
			'
			'	rsAppt.Update
			'	'send us an e-mail stating this stuff
			'	if booFixed = false then
			'		emailGCN(cnSQL)
			'	end if
			'else
				rsAppt.Update
			'end if

			If Request.Form("cboStatus") = "c" AND bitEMail = 1 Then
					AppointmentEmailConfirm intID
			End If

		End If
		
		' save Task Notes
		strTaskNotes = Request.Form("txtTaskNotes")
		cnSQL.Execute "delete tlnkAppointmentNotes where AppointmentID = " & intID
		aTaskNotes = split(strTaskNotes,"|")
		for i = lbound(aTaskNotes) to ubound(aTaskNotes)
			aRec = split(aTaskNotes(i),"~")
			cnSQL.Execute "insert tlnkAppointmentNotes (AppointmentID, ActionID, NotesFieldID, Data) values (" & intID & ", " & Request.Form("cboAction") & ", " & aRec(0) & ",	'" & replace(left(aRec(1),1024),"'","''") & "')"
		next
		
		' do department saving here
		set cmDept = server.CreateObject("adodb.command")
		cmDept.ActiveConnection = cnSQL
		cmDept.CommandType = 4
		cmDept.CommandText = "sp_SaveAppointmentDepartments"
		cmDept.Parameters.Append cmDept.CreateParameter("@Mode",adChar,adParamInput,1,"e")
		cmDept.Parameters.Append cmDept.CreateParameter("@CompanyID",adInteger,adParamInput,,cid)
		cmDept.Parameters.Append cmDept.CreateParameter("@AppointmentID",adInteger,adParamInput,,intID)
		cmDept.Parameters.Append cmDept.CreateParameter("@UserID",adInteger,adParamInput,,intCurrentUserID)
		cmDept.Parameters.Append cmDept.CreateParameter("@DepartmentList",adVarChar,adParamInput,256,ddid)
		cmDept.Execute
		set cmDept = nothing
	End If

   %>
	<!--#include file = "ReminderInc.asp"-->
   <%
	if rsAppt.State=1 Then rsAppt.Close


	' If called from Itinerary...
	if Request.QueryString("iid") <> "" then
		cnSQL.Execute "sp_ItineraryAssign " & intID & ", " & Request.QueryString("iid") & ", 'Assign'"
	end if

	
	' ****** RECURRENCE RELATED CODE
	
	intCompanyID=cid
	intApptID = intID
	intRecEdit = Request.Form("RecEdit")
	intRecID = Request.Form("RecID")
	intOccurences = Request.Form("RecOccurences")
	intRecOnce = Request.Form("RecShowOnce")
	
	If intOccurences = "" Then intOccurences = null
	
	intRecEndDate = Request.Form("RecEndDate")
	
	If intRecEndDate = "" Then intRecEndDate = null
	
	intRecType = Request.Form("RecType")
	intFrequency = Request.Form("RecFrequency")
	intRecDetail = Request.Form("RecDetail")
	If intRecDetail = "" Then intRecDetail = Null
	


If Len(intRecType) > 0 and intRecType <> "0" Then

	Set cm = Server.CreateObject ("Adodb.command")
	cm.ActiveConnection = cnSQL
	cm.CommandType = 4

	If Len(intRecEdit) > 0 and intRecEdit <> "0" Then
	' Existing Recurrence for this Appt (EDIT)
	

		If Trim(intRecEdit) = "2" Then
			cm.CommandText = "sp_Recurrence_UpDateSeries" ' Create the entry for the recurence and get ID

			set pmApptID = cm.CreateParameter ("@apptID",adInteger,adParamInput,,intApptID)
			set pmRecID = cm.CreateParameter ("@RecID",adInteger,adParamInput,,intRecID)
				
			cm.Parameters.Append pmApptID
			cm.Parameters.Append pmRecID
			
			cm.Execute 
		End IF
		
		
		
		If Request.Form("RecUpdate") = "1" and intRecID > 0 Then
		
		Response.Write "Delete tblAppointment where RecID=" & intRecID & " and AppointmentID<>" & intApptID
		
			cnSQL.Execute "Delete tblAppointment where RecID=" & intRecID & " and AppointmentID<>" & intApptID
			
				Set cm3 = Server.CreateObject ("Adodb.command")
				cm3.ActiveConnection = cnSQL
				cm3.CommandType = 4

			
				cm3.CommandText = "sp_Recurrence_Insert" ' Create the entry for the recurence and get ID

				set pmApptID3 = cm3.CreateParameter ("@apptID",adInteger,adParamInput,,intApptID)
				set pmCompanyID3 = cm3.CreateParameter ("@CompanyID",adInteger,adParamInput,,intCompanyID)
				'set pmStartDate = cm.CreateParameter ("@StartDate",adDate,adParamInput,,dtStartDate)
				set pmEndDate3 = cm3.CreateParameter ("@EndDate",adDate,adParamInput,,intRecEndDate)
		
		
				set pmOccurences3 = cm3.CreateParameter ("@Occurences",adInteger,adParamInput,,intOccurences)
				set pmFrequency3 = cm3.CreateParameter ("@Frequency",adInteger,adParamInput,,intFrequency)
				set pmRecType3 = cm3.CreateParameter ("@RecType",adChar,adParamInput,1,intRecType)
				set pmRecDetail3 = cm3.CreateParameter ("@RecDetail",adVarChar,adParamInput,50,intRecDetail)
				set pmRecShowOnce3 = cm3.CreateParameter ("@RecShowOnce",adVarChar,adParamInput,1,intRecOnce)

				cm3.Parameters.Append pmApptID3
				cm3.Parameters.Append pmCompanyID3
				'cm.Parameters.Append pmStartDate
				cm3.Parameters.Append pmOccurences3
				cm3.Parameters.Append pmEndDate3
				cm3.Parameters.Append pmRecType3
				cm3.Parameters.Append pmFrequency3
				cm3.Parameters.Append pmRecDetail3
				cm3.Parameters.Append pmRecShowOnce3
							   
				set rsRec = cm3.Execute 
				set cm3 = Nothing

				intRecEdit = rsRec("RecID")
		
				Response.Write intApptID & "," & intRecEdit
		
				Set cm2 = Server.CreateObject ("Adodb.command")
				cm2.ActiveConnection = cnSQL
				cm2.CommandType = 4
		
				cm2.CommandText = "sp_Recurrence_LinkWithAppt"
		

				set pmApptID2 = cm2.CreateParameter ("@apptID",adInteger,adParamInput,,intApptID)
				set pmRecID2 = cm2.CreateParameter ("@RecID",adInteger,adParamInput,,intRecEdit)
				set pmDeptIDs = cm2.CreateParameter ("@DeptIDs",adVarChar,adParamInput,256,ddid)
		
				cm2.Parameters.Append pmApptID2
				cm2.Parameters.Append pmRecID2
				cm2.Parameters.Append pmDeptIDs
		
				cm2.Execute 
				Set rsRec = Nothing
				Set cm2 = Nothing
				
		End IF
		
	
	Else
	' New Recurrnece

		cm.CommandText = "sp_Recurrence_Insert" ' Create the entry for the recurence and get ID

		set pmApptID = cm.CreateParameter ("@apptID",adInteger,adParamInput,,intApptID)
		set pmCompanyID = cm.CreateParameter ("@CompanyID",adInteger,adParamInput,,intCompanyID)
		'set pmStartDate = cm.CreateParameter ("@StartDate",adDate,adParamInput,,dtStartDate)
		set pmEndDate = cm.CreateParameter ("@EndDate",adDate,adParamInput,,intRecEndDate)
		
		
		set pmOccurences = cm.CreateParameter ("@Occurences",adInteger,adParamInput,,intOccurences)
		set pmFrequency = cm.CreateParameter ("@Frequency",adInteger,adParamInput,,intFrequency)
		set pmRecType = cm.CreateParameter ("@RecType",adChar,adParamInput,1,intRecType)
		set pmRecDetail = cm.CreateParameter ("@RecDetail",adVarChar,adParamInput,50,intRecDetail)
		set pmRecShowOnce = cm.CreateParameter ("@RecShowOnce",adVarChar,adParamInput,1,intRecOnce)

		cm.Parameters.Append pmApptID
		cm.Parameters.Append pmCompanyID
		'cm.Parameters.Append pmStartDate
		cm.Parameters.Append pmOccurences
		cm.Parameters.Append pmEndDate
		cm.Parameters.Append pmRecType
		cm.Parameters.Append pmFrequency
		cm.Parameters.Append pmRecDetail
		cm.Parameters.Append pmRecShowOnce
					   
		set rsRec = cm.Execute 
		set cm = Nothing

		intRecEdit = rsRec("RecID")
		
		Response.Write intApptID & "," & intRecEdit
		
		Set cm2 = Server.CreateObject ("Adodb.command")
		cm2.ActiveConnection = cnSQL
		cm2.CommandType = 4
		
		cm2.CommandText = "sp_Recurrence_LinkWithAppt"
		

		set pmApptID2 = cm2.CreateParameter ("@apptID",adInteger,adParamInput,,intApptID)
		set pmRecID2 = cm2.CreateParameter ("@RecID",adInteger,adParamInput,,intRecEdit)
		set pmDeptIDs = cm2.CreateParameter ("@DeptIDs",adVarChar,adParamInput,256,ddid)
		
		cm2.Parameters.Append pmApptID2
		cm2.Parameters.Append pmRecID2
		cm2.Parameters.Append pmDeptIDs
		

		cm2.Execute 
		Set rsRec = Nothing
		Set cm2 = Nothing
		
		
	
	End IF
	
End If
	
	
sub emailGCN(cn)
	dim rstemp, cdoobj, sfe, userfullname, guestname
	set rstemp = server.CreateObject("adodb.recordset")
	rstemp.Open "select companyname,city,phone from tblCompany where companyid = " & cid,cn,adOpenForwardOnly,adLockReadOnly
	hotelinfo = rstemp("companyname") & ", " & rstemp("city") & "  -  " & PhoneMask(rstemp("phone"))
	rstemp.Close
	set rstemp = nothing
	set cdoobj = Server.CreateObject("CDONTS.NewMail")

	userfullname = remote.Session("FloatingUser_UserName") & " " & remote.Session("FloatingUser_UserLName")
	cdoobj.To = "requests@goconcierge.net"
	if remote.Session("FloatingUser_EMail") = "" then
		sfe = "No Return E-Mail Address"
	else
		sfe = remote.Session("FloatingUser_EMail")
	end if
	cdoobj.BodyFormat = 0
	cdoobj.MailFormat = 0
	cdoobj.From = userfullname & " <" & sfe & ">"
	cdoobj.Subject = formatdatetime(date,2) & " " & formatdatetime(now(),3) & " - GCN Vendor Not-in-List Notice"
    
	guestname = (unescape(Request.Form("txtSalutation")) + " ") & (Request.Form("txtGuestFirstName") + " ") & Request.Form("txtGuestLastName")

	strBody = "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">" & vbCrLf
	strBody = strBody & "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1""><title>Not-in-list notice</title></head>"
	strBody = strBody & "<body style=""font-family:tahoma;font-size:11px"" bgcolor=white>"
	strBody = strBody & "<table cellspacing=2 style=border-style:solid;border-width:2px;border-color:black>"
	strBody = strBody & "<tr><td colspan=2 bgcolor=#FAD667 style=font-size:11px;border-style:outset;border-width:1px>"
	strBody = strBody & "A user has entered a vendor in a task that is not in our location table or has not been assigned to their hotel." & vbcrlf & vbcrlf
	strBody = strBody & "</td></tr>"
	strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>Property:</td><td style=font-size:11px bgcolor=#FAD667>" & hotelinfo & "</td></tr>"
	strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>Guest Name:</td><td style=font-size:11px bgcolor=#FAD667>" & guestname & "</td></tr>"
	strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>Start Date:</td><td style=font-size:11px bgcolor=#FAD667>" & strStartDate & "</td></tr>"
	strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>Start Time:</td><td style=font-size:11px bgcolor=#FAD667>" & strStartTime & "</td></tr>"
	strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>User:</td><td style=font-size:11px bgcolor=#FAD667>" & userfullname & "</td></tr>"
	strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>Vendor:</td><td style=font-size:11px bgcolor=#FAD667>" & strLocationText & "</td></tr>"
	strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>Vendor Address:</td><td style=font-size:11px bgcolor=#FAD667>" & Request.Form("Address") & "</td></tr>"
	strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>Vendor Phone:</td><td style=font-size:11px bgcolor=#FAD667>" & PhoneMask(Trim(Request.Form("LocPhone"))) & "</td></tr>"
	strBody = strBody & "</table>"
	strBody = strBody & "</body></html>"
			
	cdoobj.Body = strBody
	cdoobj.Send
	set cdoobj = nothing
end sub	

'Response.Redirect "Success.asp?TargetDate="	& Request.QueryString("TargetDate")
'Response.End
cnSQL.Close
Set cnSQL = Nothing
%>
