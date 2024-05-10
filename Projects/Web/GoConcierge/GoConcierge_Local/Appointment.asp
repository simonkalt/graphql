<%@ Language=VBScript %>
<%
	Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	Response.Buffer = false

	Set remote = Server.CreateObject ("UserClient.MySession")
	remote.Init (Request.Cookies("UserKey"))
	
	
	cid = remote.Session("CompanyID")
	fua = remote.session("FloatingUser_Admin")	
	fuid = remote.session("FloatingUser_UserID")
	aid = request.querystring("ID")
	copyTask = Request.QueryString("CopyTask")
	booGuestProfile = (remote.Session("UseGuestProfile") = "True")
	gpsid = remote.Session("GPSearchID")
	booUseID = (gpsid <> "0")
	
	'Response.Write Request.Cookies.Count  & " Test"
	'Response.Write Request.QueryString()
	'Response.End 
	
	Randomize()
	intRndNum = rnd() * 10
	
	Dim booNewRec
	Dim cnSQL, rsSQL, rsSalutations, rsVendor, rsFunButtons, rsReminder
	dim rsUser,	rsEditUser, rsClosedUser, rsActionType, rsAction
	Dim strBGColor, strGuestStyle
	
	strGuestStyle = "background-color:#E6AE06"
	strBGColor = "#F9D568" ' "#d4d0c8" '
	
	Set cnSQL = Server.CreateObject("ADODB.Connection")
	Set rsSQL = Server.CreateObject("ADODB.Recordset")
	
	
	Set rsSalutations = Server.CreateObject("ADODB.Recordset")
	Set rsVendor = Server.CreateObject("ADODB.Recordset")
	
	cnSQL.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")

	' Get PeopleID
	Set rsAction = Server.CreateObject("ADODB.Recordset")
	strSQL = "select NotesFieldID from tlkpNotesFields where NotesField = '# People'"
	rsAction.Open strSQL, cnSQL, adOpenDynamic, adLockReadOnly
	if rsAction.EOF then
	    sintPeopleID = 0
	else
	    sintPeopleID = rsAction("NotesFieldID")
	end if
	rsAction.Close

    ' Get List of Salutations
	Set rsSalutations = cnSQL.Execute  ("SELECT 0 as SortOrder, '' as Salutation UNION Select SortOrder, Salutation from tblSalutations Order by SortOrder")
	
	set rsActionDef = cnSQL.Execute ("Select ActionType from tblCompany where CompanyID=" & remote.Session ("CompanyID") )
	strDefaultAction = rsActionDef(0)
	if isnull(strDefaultAction) then
		strDefaultAction = 0
	end if
	
	'Response.Write strDefaultAction
	'Response.End 
	
	set rsActionDef = Nothing
	    
	Set rsSQL = cnSQL.Execute("Select * From tblAppointment Where AppointmentID=" & aid)

	' Check if this is an existing appointment
	If aid <> "0" Then 
			strGuestID = rsSQL.Fields("GuestID").Value
			strDisplayID = rsSQL.Fields("DisplayID").Value
			
			intOTLogID = rsSql("OTLogID")	' assign OT reservation Log ID
			intSSLogID = rsSql("SSLogID")   ' Assign SS Res ID
	
			If intOTLogID <> "" Then
				intOTLogID = CLng(intOTLogID)
			End IF
			
			If intSSLogID <> "" Then
				intSSLogID = CLng(intSSLogID)
			End IF
	else
		strGuestID = Request.QueryString("gid")
		strDisplayID = Request.QueryString("dgid")
	End If
	
	if intOTLogID = "" Then intOTLogID = "0"
	if intSSLogID = "" Then intSSLogID = "0"
	
	if isnumeric(remote.session("FloatingUser_SuperUser")) then
		su = remote.session("FloatingUser_SuperUser")
	else
		if remote.session("FloatingUser_SuperUser") = "True" then
			su = 1
		else
			su = 0
		end if
	end if

	
	set rsTemp = server.CreateObject("adodb.recordset")
	If IsNumeric(aid) And (aid > 0) Then
		booNewRec = false
		fuddid = rsSQL.Fields("DepartmentID").Value
	Else ' this is a new appt.
		booNewRec = true
		fuddid = remote.session("DefaultDepartmentID")
		set rsTemp = cnSQL.Execute("select DepartmentID from tlnkUserDepartment where UserID = " & fuid & " and DepartmentID = " & fuddid)
		if rsTemp.EOF then
			fuddid = remote.session("FloatingUser_DDID")
		end if
	End If
	
	if copyTask = "True" then
		fuddid = remote.session("DefaultDepartmentID")
	end if	
	
	if su = 1 then
		booDeptSelect = true
	else
		set rsTemp = cnSQL.Execute("select count(*) from tlnkUserDepartment where UserID = " & fuid & " and CompanyID = " & cid)
		if rsTemp(0).Value > 1 then
			booDeptSelect = true
		else
			booDeptSelect = false
		end if
		rsTemp.Close
		set rsTemp = nothing
	end if

	'if su = 1 then
	'	fuddid = remote.session("DefaultDepartmentID")
	'end if

	strLongTxt = 266
%>

<!--#INCLUDE file="checkuser.asp"-->
<!--#INCLUDE file="ssglobal.asp"-->
<!--#INCLUDE file="PhoneMask.asp"-->

<style>
	<!--
	.Label				{ font-family: Tahoma; font-size: 11 }
	.Field				{ font-family: Tahoma; font-size: 11; background-color: silver }
	.ShortTxt			{ font-family: Tahoma; font-size: 11; HEIGHT: 19px; LEFT: 335px; TOP: 16px; WIDTH: 150px; background-color: white }
	.ShortestTxt		{ font-family: Tahoma; font-size: 11; HEIGHT: 19px; LEFT: 335px; TOP: 16px; WIDTH: 92px }
	.ShortTxtMargin		{ font-family: Tahoma; font-size: 11; HEIGHT: 19px; LEFT: 335px; TOP: 16px; WIDTH: 150px; padding-left: 5px; border-style: outset; background-color: <%=strBGColor%> }
	.ShortestTxtMargin	{ fo nt-family: Tahoma; font-size: 11; HEIGHT: 19px; LEFT: 335px; TOP: 16px; WIDTH: 80px; padding-left: 5px; border-style: outset; background-color: <%=strBGColor%> }
	.MedTxtMargin		{ font-family: Tahoma; font-size: 11; HEIGHT: 19px; LEFT: 335px; TOP: 16px; WIDTH: 120px; padding-left: 5px; border-style: outset; background-color: <%=strBGColor%> }
	.MedTxt				{ font-family: Tahoma; font-size: 11; HEIGHT: 19px; LEFT: 335px; TOP: 16px; WIDTH: 350px; background-color: white }
	.MedTxtPV			{ HEIGHT: 20px; LEFT: 335px; TOP: 16px; WIDTH: 350px }

	.LongTxt			{ font-family: Tahoma; font-size: 11; WIDTH: <%=strLongTxt%>px; background-color: silver }
		
	.TallLongTxt		{ font-family: Tahoma; font-size: 11; HEIGHT: 246px; WIDTH: 326px; background-color: white }
	.TallMedTxt			{ font-family: Tahoma; font-size: 11; HEIGHT: 246px; WIDTH: 306px; background-color: white }
	.DateCmb			{ font-family: Tahoma; font-size: 11; HEIGHT: 18px; LEFT: 335px; TOP: 16px; WIDTH: 95px }
	.MedTxt2			{ font-family: Tahoma; font-size: 11; HEIGHT: 19px; TOP: 16px; WIDTH: 157px; background-color: white }
	.MedTxt3			{ font-family: Tahoma; font-size: 11; HEIGHT: 19px; TOP: 16px; WIDTH: 300px; background-color: white }
	.txt0			    { font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 25px; TOP: 10px; background-color: white }
	.txt1			    { font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 60px; TOP: 10px; background-color: white }
	.txt2			    { font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 100px; TOP: 10px; background-color: white }
	.txt3			    { font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 220px; TOP: 10px; background-color: white }
	.txt4			    { font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 250px; TOP: 10px; background-color: white }
	.txt5			    { font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 350px; TOP: 10px; background-color: white }
	.col3			    { font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 155px; TOP: 10px; background-color: white }
	.lastname		    { font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 94px; TOP: 10px; background-color: white }
	.col4				{ font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 110px; TOP: 10px; background-color: white }
	.col4half			{ font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 46px; TOP: 10px; background-color: white }
	.exp				{ font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 43px; TOP: 10px; background-color: white }
	.col5			    { font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 76px; TOP: 10px; background-color: white }
	A	{color:blue}
	.txtE				{ font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 210px; TOP: 10px; background-color: white }
	.txtV				{ font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 204px; TOP: 10px; background-color: white }
	.txtX			    { font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 69px; TOP: 10px; background-color: white }
	.txtPhone		    { font-family: Tahoma; font-size: 11; HEIGHT: 19px; WIDTH: 101px; TOP: 10px; background-color: white }
	.txtDateAdded		{ font-family: Tahoma; font-size: 11; font-weight: bold; background-color: transparent }
	-->
</style>

<script language="javascript1.2" src="BrowseSelect2.js?v=2"></script>

<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">

<script src="CheckIfTaskExists.asp" language="javascript"></script>

<script LANGUAGE="vbscript" RUNAT="Server">
Public Function N2Z(pvarIn, pvarDef)
	If IsNull(pvarIn) Then
		N2Z = pvarDef
	Else
		N2Z = pvarIn
	End If
End Function

Public Function CheckNewMode(pbNewMode, pvarVal)
	If pbNewMode Then
		CheckNewMode = ""
	Else
		CheckNewMode = pvarVal
	End If
End Function

</script>


<script Language="JavaScript1.2">
<!--#INCLUDE file="ddEdit.asp"-->
</script>


<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
dim strOpener
dim booToFocus
dim strCalledFrom
dim timerRemind

strCalledFrom = ""
booToFocus = 1

Sub reportViewer_LoadCompleted
	reportViewer.PrintReport False
	window.close
End Sub

Sub RoomValidate()
	document.all("txtRoom").value = Trim(document.all("txtRoom").value)
End Sub

function calert(msg)
    a = msgbox (msg, vbYesNo,"")
	if a = vbYes Then 
		calert = 1 
	else 
		calert = 0 
	End If
End function

function yesno( str, title )
	if msgbox(str,vbYesNo,title) = vbYes then
		yesno = true
	else
		yesno = false
	end if
end function

Function getpeople()

		intPeople = Inputbox("Please enter the number of people:","OpenTable - # of People")
		getpeople = intPeople

End Function

	dim intOldActionID, booDelete
	booDelete = true
	
	sub AdornStatusCombo()
		select case document.all("cboStatus").value
			case "o"
				document.all("cboStatus").style.backgroundColor = "#FFB3B3"
				document.all("txtNotes").style.backgroundColor = "silver"
			case "p"
				document.all("cboStatus").style.backgroundColor = "#FFB353"
				document.all("txtNotes").style.backgroundColor = "#FFB353"
			case "c"
				document.all("cboStatus").style.backgroundColor = "lightgreen"
				document.all("txtNotes").style.backgroundColor = "lightgreen"
			case "x"
				document.all("cboStatus").style.backgroundColor = "lightblue"
				document.all("txtNotes").style.backgroundColor = "lightblue"
			case "r"
				document.all("cboStatus").style.backgroundColor = "#82FFE0"
				document.all("txtNotes").style.backgroundColor = "#82FFE0"
			case "n"
				document.all("cboStatus").style.backgroundColor = "lavender"
				document.all("txtNotes").style.backgroundColor = "lavender"
			case "w"
				document.all("cboStatus").style.backgroundColor = "#cc99ff"
				document.all("txtNotes").style.backgroundColor = "#cc99ff"
				
		end select
	end sub

	sub cboAction_onchange
		dim booGo, z
				
		on error resume next
			set z = document.frames("frameTaskNotes").document.all.tags("INPUT")
		on error goto 0
					
		booGo = true
		if z.length > 0 then
			booAsk = false
			for zz = 0 to z.length-1
				if len(trim(z(zz).value)) > 0 then
					booAsk = true
					exit for
				end if
			next
			if booAsk then
				if msgbox("Clicking 'OK' will overwrite your existing Task Notes." & vbcrlf & vbcrlf & "Proceed?",vbQuestion+vbOKCancel,"Task Note Replace") = vbCancel then
					booGo = false
				end if
			end if
		end if
		if booGo then
			Randomize()
			r = Rnd() * 1000
			
			taskNotes.action = "AppointmentTaskNotes.asp?ActionID=" + document.all("cboAction").value + "&AppointmentID=<%=aid%>&r="+cstr(r)
			taskNotes.target = "frameTaskNotes"
			taskNotes.submit()
			'window.frameTaskNotes.document.location.replace "AppointmentTaskNotes.asp?ActionID=" + document.all("cboAction").value + "&AppointmentID=<%=aid%>&r="+cstr(r)
			intOldActionID = document.all("cboAction").value
		else
			document.all("cboAction").value = intOldActionID
		end if
		'window.event.returnValue = (not booGo) ' ???? What is this for (ilia)
		
		<% if CLng("0" & remote.Session("SS_CompanyID")) > 0 Then %> ' This only applies to hotels who have SS Configured
		
				on error resume next
					se = window.event.srcElement.id ' Source element only to check if the action has actually been changed by the user
				on error resume next 
		
				if se <> "" and document.all("cboAction").options(document.all("cboAction").selectedIndex).text = "Super Shuttle" Then
					  x = calert ("Would you like to make an online Super Shuttle Reservation?")
						if (x=1) Then
						 MakeSSReservation (0)
						else
							document.all("OTButton").innerHTML = "<a href=""javascript:MakeSSReservation(0)"">Create Super Shuttle Reservation</a>"
						End if
				Else
						document.all("OTButton").innerHTML = ""
				End if
				
		<% end if %>
		
	end sub


	dim booFromFrame
	dim lastElement

	Sub window_onload
		<%if booGuestProfile then%>
			set lastElement = document.all("txtGuestID").parentElement
		<%else%>
			set lastElement = document.all("txtRoom").parentElement
		<%end if%>

		booFromFrame = instr(1,parent.document.location,"AppointmentFrame.asp") > 0
		if booFromFrame then
			strOpener = ""
			document.all("cmdSaveAndCopy").style.visibility = "hidden"
		else
			strOpener = window.opener.document.location.href
		end if 
	    
		if instr(1,strOpener,"ItineraryDetailEdit.asp") > 0 or instr(1,strOpener,"ItineraryTask.asp") > 0 then
			strCalledFrom = "IDE"
		end if

	    frames("frm1").cobj.parent = "d1obj"
	    window.divElipse.tabIndex = 0
	    
	    document.all("cboLetterHead").value = "<%=remote.Session("LetterHead")%>"

		txtApptDate = "<%=Request.QueryString("TargetDate")%>"
	    
		<%
		If booNewRec Then
			response.write "document.all(""chkRollover"").checked = " & remote.session("RolloverDefault") & vbcrlf
			
		 %>
			document.all("pvDateStart").value = txtApptDate
			document.all("pvDateEnd").value = txtApptDate
			document.all("txtDateAdded").value = txtApptDate

			d1obj.setVal document.all("pvDateStart").value,1
			
			<% If Len(Request.QueryString("Hour")) > 0 Then %>
          
			for zz = 0 to document.all("txtStartTime").options.length - 1
			
				st = Replace(document.all("txtStartTime").options(zz).text,chr(32),"")
				ct = Replace("<%=AMPM(Request.QueryString("Hour"),"00")%>"," ","")
			
				If len(st) <> len(ct) Then st = "0" & st
			
				If st = ct Then
					document.all("txtStartTime").selectedIndex = zz
					exit for
				End If
			 Next
			 
			for zz = 0 to document.all("txtEndTime").options.length - 1
			
				st = Replace(document.all("txtEndTime").options(zz).text,chr(32),"")
				ct = Replace("<%=AMPM(Request.QueryString("Hour"),"00")%>"," ","")
			
				If len(st) <> len(ct) Then st = "0" & st
			
				If st = ct Then
					document.all("txtEndTime").selectedIndex = zz
					exit for
				End If
			 Next
			 
			<%Else
				dtCurrent = dateadd("h",12,Date())%>
				for zz = 0 to document.all("txtStartTime").options.length - 1
						
					st = Replace(document.all("txtStartTime").options(zz).text,chr(32),"")
					ct = Replace("<%=CustomTime(FormatDateTime(dtCurrent,3))%>"," ","")
			
					If len(st) <> len(ct) Then st = "0" & st
			
					If st = ct Then
						document.all("txtStartTime").selectedIndex = zz
						exit for
					End If
				Next
						 
				for zz = 0 to document.all("txtEndTime").options.length - 1
						
					st = Replace(document.all("txtEndTime").options(zz).text,chr(32),"")
					ct = Replace("<%=CustomTime(FormatDateTime(DateAdd("n",1,dtCurrent),3))%>"," ","")
			
					If len(st) <> len(ct) Then st = "0" & st
			
					If st = ct Then
						document.all("txtEndTime").selectedIndex = zz
						exit for
					End If
				 Next
			<%End If
			if request.querystring("NoTime") = "True" then
				response.write "document.all(""chkNoTime"").checked = true" & vbcrlf
				response.write "document.all(""chkNote"").checked = true" & vbcrlf
				response.write "document.all(""chkNote"").disabled = true" & vbcrlf
				response.write "document.all(""chkNoteHidden"").value = ""on""" & vbcrlf
				response.write "document.all(""txtStartTime"").value = """"" & vbcrlf
				response.write "document.all(""txtEndTime"").value = """"" & vbcrlf
			end if				
				' If New Location with LocationID		
				strPhone = ""
				if Trim(request.querystring("LocID")) <> "" Then
					set rsVendor = cnSQL.Execute("SELECT Locationid,CompanyName,Phone,Street,City,State,OTID FROM tblLocation where locationid = " & Trim(request.querystring("LocID")))
					if rsVendor.eof then
						strLocationText = ""
						strOTID = ""
						strLocationID = 0
						strPhone = ""
						strStreet = ""
						strCity = ""
						strState = ""
					else
						strLocationText = rsVendor("CompanyName")
						strOTID = rsVendor("OTID")
						strLocationID = rsVendor("LocationID")
						strPhone = rsVendor("Phone")
						strStreet = rsVendor("Street")
						strCity = rsVendor("City")
						strState = rsVendor("State")
						
						'If Instr(1,Ucase(strLocationText),"SUPER") > 0 and Instr(1,Ucase(strLocationText),"SHUTTLE") > 0 Then
						
						'End If
						
					end if
					rsVendor.close
					set rsVendor = nothing
				End IF
				%>
				document.all("txtLocation").value = "<%=strLocationText%>"
				document.all("txtLocationID").value = "<%=strLocationID%>"
				document.all("txtVendor").value = "<%=strLocationText%>"
				document.all("txtLocPhone").value = "<%=strPhone%>"
				document.all("txtLocAddress").value = "<%=strStreet%>"
				document.all("txtLocState").value = "<%=strState%>"
		<%Else 

			if rsSQL.fields("Rollover").value then
				booRollover = "true"
				strDateAddedVis = "document.all(""divDateAdded"").style.visibility = ""visible""" & vbcrlf
			else
				strDateAddedVis = "document.all(""divDateAdded"").style.visibility = ""hidden""" & vbcrlf
				booRollover = "false"
			end if
			
			if rsSQL.fields("Span").value then
				booSpan = "true"
				strSpanDisabled = ""
			else
				booSpan = "false"
			end if
			
			response.write "document.all(""chkSpan"").checked = " & booSpan & vbcrlf
			
			' remarked out because Keith does not want to see it here
			'response.write strDateAddedVis 
			
			pubStartTime = CustomTime(FormatDateTime(rsSQL("ApptStartDate"),3))
			pubEndTime = CustomTime(FormatDateTime(rsSQL("ApptEndDate"),3))

			%>

			document.all("pvDateStart").value = "<%=FormatDateTime(rsSQL.Fields("ApptStartDate").Value,2)%>"
			d1obj.setVal document.all("pvDateStart").value,1
			
			document.all("pvDateEnd").value = "<%=FormatDateTime(rsSQL.Fields("ApptEndDate").Value,2)%>"
			
			for zz = 0 to document.all("txtStartTime").options.length - 1
			
				st = Replace(document.all("txtStartTime").options(zz).text,chr(32),"")
				ct = Replace("<%=pubStartTime%>"," ","")
			
				If len(st) <> len(ct) Then st = "0" & st
			
				If st = ct Then
			    	document.all("txtStartTime").selectedIndex = zz
				End If
			Next
			 
			for zz = 0 to document.all("txtEndTime").options.length - 1
			
				st = Replace(document.all("txtEndTime").options(zz).text,chr(32),"")
				ct = Replace("<%=pubEndTime%>"," ","")
			
				If len(st) <> len(ct) Then st = "0" & st
			
				If st = ct Then
					document.all("txtEndTime").selectedIndex = zz
				End If
			Next
			
			<% ' Note check stuff...
			if rsSQL.Fields("NoTime").value then
			response.write "document.all(""chkNoTime"").checked = true" & vbcrlf
			response.write "document.all(""chkNote"").disabled = true" & vbcrlf
			response.write "document.all(""chkNoteHidden"").value = ""on""" & vbcrlf
			response.write "document.all(""txtStartTime"").value = """"" & vbcrlf
			response.write "document.all(""txtEndTime"").value = """"" & vbcrlf
			end if%>

			document.all("txtRoom").value = "<%=rsSQL("Room")%>"
			document.all("pvSalutation").Value = "<%=escape(rsSQL("Salutation"))%>"
	
			<%'sk - 8/27/03 - DateAdded (Rollover Start)
			if isnull(rsSQL.Fields("DateAdded").Value) then
				dDateAdded = FormatDateTime(rsSQL.Fields("ApptStartDate").Value,2)
			else 
				dDateAdded = FormatDateTime(rsSQL.Fields("DateAdded").Value,2)
			end if
			response.write "document.all(""txtDateAdded"").value = """ & dDateAdded & """" & vbcrlf

			if isnull(rsSQL("LocationID")) or rsSQL("LocationID") = 0 then
				strLocationText = rsSQL("LocationText")
				strLocationID = 0
			else
				strLocationID = rsSQL("LocationID").value
				set rsVendor = cnSQL.Execute("SELECT CompanyName, OTID FROM tblLocation where locationid = " & rsSQL("LocationID"))
				if rsVendor.eof then
					strLocationText = ""
					strOTID = ""
				else
					strLocationText = rsVendor("CompanyName")
					strOTID = rsVendor("OTID")
				end if
				rsVendor.close
				set rsVendor = nothing
			end if
			
			' Reminder Get
			Set rsReminder = Server.CreateObject("adodb.recordset")
			set rsReminder = cnSQL.Execute("select * from tblReminder where AppointmentID = " & aid)
			if not rsReminder.EOF then
				sRDate = cstr(rsReminder.Fields("ReminderDateTime").Value)
				booNoTime = cbool(rsReminder.Fields("NoTime").Value)
				if booNoTime then
					sDate = sRDate
					sTime = ""
				else
					'response.write sRDate
					'response.end
					if instr(1,sRDate," ") > 1 then
						sDate = mid(sRDate,1,instr(1,sRDate," ")-1)
						sTime = mid(mid(sRDate,instr(1,sRDate," ")+1),1,instrrev(mid(sRDate,instr(1,sRDate," ")+1),":")-1) & " " & right(trim(sRDate),2)
					else
						' an error occured while saving so fix it here.
						sDate = sRDate
						sTime = ""
						booNoTime = true
					end if
				end if
				response.write "document.all(""txtReminder"").value = """ & sDate & "|" & sTime & "|" & booNoTime & "|" & rsReminder.Fields("Days").Value & "|" & trim(rsReminder.Fields("Type").Value) & "|" & rsReminder.Fields("Note").Value & "|" & cbool(rsReminder.Fields("Rollover").Value) & "|" & rsReminder.Fields("Status").Value & "|" & rsReminder.Fields("DateOrDays").Value & """" & vbcrlf
			end if
			rsReminder.Close
			set rsReminder = nothing
			'
			
			
			' Code to create Task According to Location ID
			if Trim(request.querystring("LocID")) <> "" Then
				set rsVendor = cnSQL.Execute("SELECT Locationid,CompanyName, OTID FROM tblLocation where locationid = " & Trim(request.querystring("LocID")))
				if rsVendor.eof then
					strLocationText = ""
					strOTID = ""
					strLocationID = 0
				else
					strLocationText = rsVendor("CompanyName")
					strOTID = rsVendor("OTID")
					strLocationID = rsVendor("LocationID")
				end if
				rsVendor.close
				set rsVendor = nothing
			End IF
			%>
			
			document.all("txtLocation").value = "<%=strLocationText%>"
			document.all("txtLocationID").value = "<%=strLocationID%>"
			document.all("txtVendor").value = "<%=strLocationText%>"
			document.all("cboChargeTo").value = "<%=rsSQL("CCType")%>"
			if document.all("cboChargeTo").selectedIndex > -1 then
				document.all("txtChargeTo").value = document.all("cboChargeTo").options(document.all("cboChargeTo").selectedIndex).text
			else
				document.all("txtChargeTo").value = ""
			end if
			document.all("txtLocPhone").value = "<%=Trim(CheckNewMode(booNewRec,rsSQL("LocPhone")))%>"
		<%End If
		
		if request.querystring("ButtonID") <> "" then
			set rsFunButtons = server.CreateObject("ADODB.recordset")
			set rsFunButtons = cnSQL.execute("select * from vw_FunButtons where ButtonID=" & request.querystring("ButtonID"))


			if isnull(rsFunButtons("OtID")) then
				strOtID = 0
			else
				strOtID = rsFunButtons("OtID")
			end if
			if isnull(rsFunButtons("LocationID")) then
				intLocationID = 0
			else
				intLocationID = rsFunButtons("LocationID")
			end if
			if isnull(rsFunButtons("DefaultAction")) then
				intDefaultAction = 0
			else
				intDefaultAction = rsFunButtons("DefaultAction")
			end if
			if isnull(rsFunButtons("DefaultType")) then
				intDefaultType = 0
			else
				intDefaultType = rsFunButtons("DefaultType")
			end if
			
			if strOtID > 0 Then
				response.write "document.all(""OTButton"").innerHTML = ""<a href=""""javascript:MakeOTReservation(0)"""">Create OpenTable Reservation</a>""" & vbcrlf
			End IF

			if rsFunButtons.Fields("NoTime").Value then
				rsFNBStartTime = ""
				rsFNBEndTime = ""
				response.write "document.all(""chkNoTime"").checked = true" & vbcrlf
				response.write "document.all(""chkNote"").checked = true" & vbcrlf
				response.write "document.all(""chkNote"").disabled = true" & vbcrlf
				response.write "document.all(""chkNoteHidden"").value = ""on""" & vbcrlf
				response.write "document.all(""txtStartTime"").value = """"" & vbcrlf
				response.write "document.all(""txtEndTime"").value = """"" & vbcrlf
			else
				if rsFunButtons("defaultCurTime") Then
					
					cdt = split(FormatDateTime(Time()+remote.session("TimeZone"),4),":")
					chr1 = cdt(0)
					cmin = cdt(1)
					
					if cmin > 0 and cmin < 16 Then cMinutes = "15"
					if cmin > 15 and cmin < 31 Then cMinutes = "30"
					if cmin > 30 and cmin < 46 Then cMinutes = "45"
					if cmin > 45 and cmin < 60 Then 
						cMinutes = "00"
						cHour = chr1 + 1
					Else
						cHour = Cint(chr1)
					End If
					
					If cHour => 12 Then
						if cHour > 12 then cHour=cHour-12
						cAMPM="PM"
					Else
					    cAMPM="AM"
					End If
					
					cResult = Cstr(cHour) & ":" & cMinutes & " " & cAMPM
					
					
					rsFNBStartTime = cResult
					rsFNBEndTime = cResult
				Else
					rsFNBStartTime = trim(rsFunButtons("DefaultStartTime"))
					rsFNBEndTime = trim(rsFunButtons("DefaultEndTime"))
				End If
			end if			
			
			response.write "selectOptionByValue document.all(""txtStartTime""),""" & rsFNBStartTime & """" & vbcrlf
			response.write "selectOptionByValue document.all(""txtEndTime""),""" & rsFNBEndTime & """" & vbcrlf
			
			
			response.write "document.all(""txtVendor"").value = """ & rsFunButtons("CompanyName") & """" & vbcrlf
			response.write "document.all(""txtLocation"").value = """ & rsFunButtons("CompanyName") & """" & vbcrlf
			response.write "searchDialog.defaultValue = """ & rsFunButtons("CompanyName") & """" & vbcrlf
			response.write "document.all(""txtLocPhone"").value = """ & rsFunButtons("Phone") & """" & vbcrlf
			response.write "document.all(""txtLocAddress"").value = """ & rsFunButtons("Street") & """" & vbcrlf
			response.write "document.all(""txtLocCity"").value = """ & rsFunButtons("City") & """" & vbcrlf
			response.write "document.all(""txtLocationID"").value = " & intLocationID & vbcrlf
			response.write "document.all(""cboAction"").value = " & intDefaultAction & vbcrlf
			response.write "document.all(""cboActionType"").value = " & intDefaultType & vbcrlf
			response.write "cboAction_onchange" & vbcrlf
			'if strOtID > 0 Then
				'response.write "window.MakeOTReservation(0)" & vbcrlf
			'End IF
			
			rsFunButtons.Close
			set rsFunButtons = nothing
		end if
		%>
		
		intOldActionID = document.all("cboAction").value
		AdornStatusCombo

		if document.all("txtLocation").value <> "" Then
			document.all("printloc").disabled = false
		Else
			document.all("printloc").disabled = true
		End If
		
		call SetupOT()
		
		<% if intOTLogID <> "0" Then %>
			document.all("OTButton").innerHTML = "<a href=""javascript:MakeOTReservation(<%=intOTLogID%>)"">Edit OpenTable reservation</a>"
			document.all("cboAction").disabled = true
			document.all("linkrec").disabled = true
		<% end if %>

		<% if intSSLogID <> "0" Then %>
			document.all("OTButton").innerHTML = "<a href=""javascript:MakeSSReservation(<%=intSSLogID%>)"">Edit Super Shuttle Reservation</a>"
			document.all("cboAction").disabled = true
			document.all("linkrec").disabled = true
		<% end if %>
		
		document.all("txtCreatedDateTime").value = DayOfWeek(document.all("txtCreatedDateTime").value) & ", " & document.all("txtCreatedDateTime").value
		If Len(document.all("txtEditDateTime").value) > 0 Then
			document.all("txtEditDateTime").value = DayOfWeek(document.all("txtEditDateTime").value) & ", " & document.all("txtEditDateTime").value
		End If
		
		AppointmentLoginConfirm

		call j_onload()

		FormatReminder
	End Sub
	
	sub FormatReminder
		if document.all("txtReminder").value <> "" then
			document.all("cmdReminder").value = "Edit Reminder"
			timerRemind = window.setInterval("startBlink()",600)
		else
			document.all("cmdReminder").value = "Add Reminder"
			window.clearInterval timerRemind
		end if			
	end sub

	sub startBlink()
		if document.all("cmdReminder").style.borderColor = "yellow" then
			document.all("cmdReminder").style.borderColor = ""
		else
			document.all("cmdReminder").style.borderColor = "yellow"
		end if
	end sub
	
	sub AppointmentLoginConfirm()
		<%
		dim strSQL, booEOF, booCompanyAdmin, strFunctions
		Dim bEmailDisabled, bEmailChecked, booOK

		booOK = true
		bEmailDisabled = "True"
		bEmailChecked = "False"

		strSQL = "SELECT * FROM viewUserCompany WHERE Password = '" & trim(remote.Session("FloatingUser_Password")) & "' AND CompanyID = " & cid
		set rsUser = cnSQL.Execute(strSQL)

		if rsUser.eof then
			strFunctions = "	msgbox ""Invalid Password"",,""Task Login""" & vbCRLF
		else
			'remote.Session("CurrentUserID") = rsUser("UserID")
			remote.Session("FloatingUser_UserID") = rsUser("UserID")
			
			if rsUser("CompanyAdmin") then
				
				if Request.querystring("ID") <> 0 then
					strFunctions = "	DeleteRights()" & vbCrLF
				else
					strFunctions = strFunctions & "	document.all(""txtCreateUserID"").value = " & rsUser("UserID") & vbCrLF
					strFunctions = strFunctions & "	document.all(""lstCreateUserName"").value = """ & rsUser("UserName") & " " & rsUser("UserLName") & """" & vbCrLF
				end if
				strFunctions = strFunctions & "	SubmitComplete()" & vbCrLF
				strFunctions = strFunctions & "	DeleteRights()" & vbCrLF
				remote.Session("LastEditedID") = rsUser("UserID")
			else ' closed fields rights
				if Len(Request.Form("txtClosedUserID")) = 0 then
					booNotClosed = true
					intNotClosedID = 0
				else
					booNotClosed = false
					intNotClosedID = cint(Request.Form("txtClosedUserID"))
				end if
																									'Changed for the new Permissioning Scheme. 10/15/01 IR			
				If  (CInt(Request.Form("txtCreateUserID")) = cint(rsUser("UserID"))) or (Request.querystring("ID") = 0) or remote.Session("ECT") Then 
					if CInt(Request.Form("txtCreateUserID")) = cint(rsUser("UserID")) then
						strFunctions = strFunctions & " DeleteRights()" & vbCrLF
					end if
					strFunctions = strFunctions & "	SubmitComplete()" & vbCrLF
				End If
				    
				if Request.querystring("ID") = 0 then
					remote.Session("CreateUserID") = rsUser("UserID")
					strFunctions = strFunctions & "	document.all(""txtCreateUserID"").value = " & rsUser("UserID") & vbCrLF
					strFunctions = strFunctions & "	document.all(""lstCreateUserName"").value = """ & rsUser("UserName") & " " & rsUser("UserLName") &  """" & vbCrLF
				end if
				remote.Session("LastEditedID") = rsUser("UserID")
			end if
			strFunctions = strFunctions & "	EnableClosedFields()" & vbCRLF

			if booNewRec AND rsUser("EmailAddress") <> "" then
				bEmailDisabled = "False"
				bEmailChecked = "True"
			Else
				if booOK then 
					If rsUser("EmailAddress") <> "" Then
						If 	rsUser("CompanyAdmin") Or rsUser("SuperUser") Or  _
								CInt(Request.Form("txtCreateUserID")) = CInt(rsUser("UserId")) Then
							bEmailDisabled = "False"
						End If
					End If
				end if
			end if
			
			strFunctions = strFunctions & " document.all(""chkEMail"").disabled = " & bEmailDisabled & vbCrLF
		end if

		Response.Write strFunctions & vbCRLF

		rsUser.close
		set rsUser = nothing
		%>
	end sub
	
	Function DayOfWeek(d)
	dt = DatePart("w",d)
		Select Case dt
			case 1: DayOfWeek = "Sun" 
			case 2: DayOfWeek = "Mon" 
			case 3: DayOfWeek = "Tue" 
			case 4: DayOfWeek = "Wed" 
			case 5: DayOfWeek = "Thu" 
			case 6: DayOfWeek = "Fri" 
			case 7: DayOfWeek = "Sat" 
		End Select
	End Function

	function selectOptionByValue(element,target)
	  for xyz = 0 to element.options.length-1
	    if element.options(xyz).value = target then
	      element.selectedIndex = xyz
	      exit for
	    end if
	  next
	end function

	function CopyTask(id)
		dim mode
				
		mode = window.showModalDialog("GetCopyMode.asp","","center:yes;status:no;scroll:no;dialogHeight:100px;dialogWidth:430px")
		if mode <> "" then
			window.opener.frameCalUpdate.location = "CopyTask.asp?ID=" & id & "&curdate=" & parent.d1obj.getVal() & "&mode=" & mode & "&CloseWin=True"
		end if
	end function

	Sub cmdCancel_onclick
		' used to need resume next in case task editing in search
		' I'll keep it in just in case (sk)
		on error resume next
		<%if copyTask = "True" then%>
			document.frames("ifLogin").location = "DeleteApptConfirm.asp?ID=<%=aid%>&Cancel=True"
		<%else%>
			if not booFromFrame then
				if instr(1,strOpener,"TaskSearch.asp") > 0 then
					window.opener.document.all.submit.click
				else
					if (cdate("<%=Request.QueryString("TargetDate")%>") <> cdate(document.all("pvDateStart").value)) then
						if ("<%=Request.QueryString("FromReminder")%>" <> "True") then
							if msgbox("Move calendar to selected task date?",vbYesNo+vbQuestion,"Task Date") = vbYes then
								window.opener.Calobj.onDateChange d1obj.getVal(),1
							end if
						end if
					end if
					d1obj = null
				end if
			end if
		<%end if%>
		if booFromFrame then
			parent.window.close
		else
			window.close
		end if
	End Sub

	Sub chkEMail_onclick
		If document.all("chkEMail").checked then
			document.all("chkEMail").value = "on"
		else
			document.all("chkEMail").value = "off"
		end if
	End Sub
	
	Sub cboStatus_onChange
	    document.all("txtNotes").disabled = (document.all("cboStatus").value <> "c" and document.all("cboStatus").value <> "p" and document.all("cboStatus").value <> "x" and document.all("cboStatus").value <> "r" and document.all("cboStatus").value <> "n" and document.all("cboStatus").value <> "w")
		if document.all("cboStatus").value = "c" or document.all("cboStatus").value = "r" then
			'document.all("txtNotes").style.backgroundcolor = "white"
			document.all("txtStateChangedToClosed").value = "true"
			document.all("txtClosedUserID").value = strLastEditUserID
		else
			document.all("txtNotes").style.backgroundcolor = "silver"
			document.all("txtStateChangedToClosed").value = "false"
		end if
		AdornStatusCombo
	End Sub

	Function NullToZero(pvarIn, pvarDef)
		If IsNull(pvarIn) Then
			NullToZero = pvarDef
		Else
			NullToZero = pvarIn
		End If
	End Function

	
	Sub EnableDelete()
		document.all("cmdDelete").disabled = false 
	End Sub

	Sub EnableClosedFields()
		document.all("cboStatus").disabled = false
		if document.all("cboStatus").value = "c" or document.all("cboStatus").value = "p"  or document.all("cboStatus").value = "r" or document.all("cboStatus").value = "n" or document.all("cboStatus").value = "w" then
			document.all("txtNotes").disabled = false
		end if
	End Sub
	
	Sub SubmitComplete()
		document.all("frmAddEditAppointment").Action = "AppointmentConfirm.asp?ID=<%=aid%>&TargetDate=<%=Request.Querystring("TargetDate")%>&CopyTask=<%=copyTask%>&CalledFrom=" & strCalledFrom & "&iid=<%=Request.QueryString("iid")%>"
		document.all("frmAddEditAppointment").Target = "ifLogin"
		if booToFocus = 1 then
			<%if booUseID then
				response.write "document.all(""txtGuestID"").focus()" & vbcrlf
			else
				response.write "document.all(""txtRoom"").focus()" & vbcrlf
			end if%>
		end if
	End Sub
	
	Sub DeleteRights()
		document.all("cmdDelete").disabled = false
	End Sub

	Sub cmdGuestTaskLetter_onclick
			SubmitTab(3)
	End Sub

	function CustomTime( strTime )
		CustomTime = right("0" & left(strTime,instrrev(strTime,":")-1),5) & " " & right(strTime,2)
	end function

	function cmdRecurrence_on()
		SubmitTab(1)
	end function
	'z = true
	function Reminder()
		aid = "<%=aid%>"
		' need this if task deleted while being edited on other system...
		if not (checkAID(aid) or aid = "0") then
			aid = "0"
		end if
		'
		x = window.showModalDialog("Reminder.asp?NewFromTask=" & document.all("cmdReminder").value & "&FromTask=True&aid=" & aid & "&Reminder=" & escape(document.all("txtReminder").value),window,"dialogHeight:480px;dialogWidth:520px;status:no;scroll:no;center:yes")
		if x <> "close" then
			document.all("txtReminder").value = x
			formatReminder
		end if
	end function

	function SubmitTab( mode )
		StartDate = d1obj.getVal()
	
		document.all("cmdSaveandCopy").disabled = true
		document.all("cmdSaveandPrint").disabled = true
		document.all("cmdSaveandClose").disabled = true
		
		if lastElement.id = "txtVendor" and document.all("txtVendor").value <> document.all("txtLocation").value and document.all("txtVendor").value <> "" then
			'do nothing. handled by body onmousedown handler
			'user must select a Vendor from the Vendor List or leave the field blank.
		else
			select case mode
				case 1
					window.open("recurrence.asp?EndBy=" & StartDate)
				case 2
					if len(window.event.srcElement.getAttribute("value")) > 0 then
						document.frmAddEditAppointment.submit()
				    	booDelete = false
					end if
				case 3 ' Guest Task Letter
					aid = CLng("0" & document.all("cboAction").options(document.all("cboAction").selectedIndex).value)
					vid =  CLng("0" & document.all("txtLocationID").value)
					pstr = "?CompanyID=<%=cid%>&ActionID=" & aid & "&VendorID=" & vid
					
					strTaskNotes = document.frames("frameTaskNotes").getNotes()

					Set xmlHttp2 = CreateObject("Microsoft.XMLHTTP")
					xmlHttp2.open "POST" , "CustomReports/GuestLetterCheck.asp" & pstr, false
					xmlHttp2.send()
					retVal = xmlHttp2.responseText
					set xmlHttp2 = Nothing 
					If instr(1,retVal,"Count:") > 0 then
						numButts = cint(mid(retVal,instr(1,retVal,":")+1))
						maxButts = 16
						if numButts > maxButts then
							numButts = maxButts
						end if
						dHeight = trim(cstr(((numButts*18)+(numButts*5.5)+100)))
						dOptions = "center:yes;status:no;dialogHeight:" & dHeight & "px;dialogWidth:470px;scroll:no"
						pstr = pstr & "&numbutts=" & numButts
						x = window.showModalDialog ("CustomReports/GuestLetterSelect2.asp" & pstr, null, (dOptions))
					Else
						x = retVal
					ENd If
					
					
					If Cstr(x) <> "" Then
							document.all("txtSalutation").value = document.all("pvSalutation").Value
							if vartype(document.frames("frameTaskNotes").document.all("txtData<%=sintPeopleID%>")) = 9 then
								strPeople = ""
							else
								strPeople = document.frames("frameTaskNotes").document.all("txtData<%=sintPeopleID%>").value
							end if
					
							str = "ReportGuestTaskLetter.asp?ID=<%=aid%>&Action=" & document.all("cboAction").options(document.all("cboAction").selectedIndex).text & "&ActionType=" & document.all("cboActionType").options(document.all("cboActionType").selectedIndex).text
							str = str & "&CompanyName=" & escape(document.all("txtLocation").value) & "&Salutation=" & document.all("txtSalutation").Value
							str = str & "&GuestName=" & document.all("txtGuestLastName").value
							str = str & "&StartTime=" & document.all("txtStartTime").value
							str = str & "&StartDate=" & StartDate
							str = str & "&TemplateID=" & escape(x)  
							str = str & "&Notes=" & AutoFormat(document.all("txtNotes").value,"&","%26")
							str = str & "&Room=" & document.all("txtRoom").value
							str = str & "&CP=" & document.all("txtLocPhone").value
							str = str & "&People=" & strPeople
							str = str & "&GP=" & document.all("txtGuestPhone").value
							str = str & "&GE=" & document.all("txtGuestEMail").value
							str = str & "&Sal=" & escape(document.all("txtSalutation").value)
							str = str & "&CA=" & escape(document.all("txtLocAddress").value)
							str = str & "&CC=" & escape(document.all("txtLocCity").value)
							str = str & "&LID=" & escape(document.all("txtLocationID").value) 
							str = str & "&TaskNotes=" & escape(strTaskNotes)
					
							window.showModelessDialog str, null, "center:yes;status:no;dialogHeight:540px;dialogWidth:670px;scroll:no"
					End If

					document.all("cmdSaveandCopy").disabled = false
					document.all("cmdSaveandPrint").disabled = false
					document.all("cmdSaveandClose").disabled = false
				case 4
					if frmAddEditAppointment_onsubmit() then
						doRep(mode)
					end if
				case 5
					document.all("frmAddEditAppointment").Action = "AppointmentConfirm.asp?ID=<%=aid%>&TargetDate=<%=Request.Querystring("TargetDate")%>"
					document.all("frmAddEditAppointment").Target = "_self"
					document.all("frmAddEditAppointment").Submit()
				Case 6
					dim h, w, booPrint, strUrl, strLocationList
		
					h = screen.availHeight*.93
					w = screen.availWidth*.97
	
					strLocationList = document.all("txtLocationID").value
	
					window.status = "Calculating Report..."
					strUrl = "ReportLocation.asp?Mode=v&type=loc&cboLetterHead=" & window.document.all("cboLetterhead").value & "&GridWidth=" & w & "&GridHeight=" & h & "&txtViewLocationList=" & strLocationList
					x = window.showModalDialog(strUrl, window, "dialogheight: " & h & "px; dialogwidth: " & w & "px; center: yes; status: no; scroll: no")
					document.all("cmdSaveandCopy").disabled = false
					document.all("cmdSaveandPrint").disabled = false
					document.all("cmdSaveandClose").disabled = false
				case 7
					if frmAddEditAppointment_onsubmit() then
						doRep(mode)
					end if
				case 8
					if frmAddEditAppointment_onsubmit() then
						document.all("frmAddEditAppointment").Action = "AppointmentConfirm.asp?ID=<%=aid%>&TargetDate=<%=Request.Querystring("TargetDate")%>&CopyTask=<%=copyTask%>&CalledFrom=" & strCalledFrom & "&iid=<%=Request.QueryString("iid")%>&SelfCopy=True"
						document.all("frmAddEditAppointment").Target = "ifLogin"
						document.all("frmAddEditAppointment").submit()
					end if
				case 9
					if frmAddEditAppointment_onsubmit() then
						document.all("frmAddEditAppointment").submit()
					end if
			end select
		end if
	end function

	function doRep( mode )
		document.all("txtSalutation").value = document.all("pvSalutation").Value
		
		document.all("cboAction").disabled = false

		booDelete = false
		
		document.all("txtSubjectSave").value = document.all("txtSubject").value
	
		on error resume next
			If IsObject(Window.opener.Calobj) Then Window.opener.Calobj.setDate(d1obj.getVal())
		on error goto 0

		buildTaskNotes

		booCCMask = not <%=remote.session("FloatingUser_VCCN")%>
		if not booCCMask then
			if trim(document.all("txtNumber").value) <> "" then
				if msgbox("Display Credit Card Number Digits?",vbYesNo+vbQuestion,"Credit Card Information") = vbYes then
					booCCMask = false
				else
					booCCMask = true
				end if
			end if
		end if
		
		if mode = 7 then
			strPS = "p"
			strPSC = "True"
			strTarget = "ifLogin"
		else
			strPS = "v"
			strPSC = "False"
			strTarget = "_self"
		end if
		'x = window.open("tmpPrintWin","_top","toolbar=no,titlebar=no,status=no,height=10,width=10,scrollbars=no,resizable=no,menubar=no,location=no,directories=no")
		document.all("CCNumber").value = document.all("txtNumber").value
		document.all("frmAddEditAppointment").Action = "PrintTaskandSave.asp?Letterhead=" & document.all("cboLetterhead").value & "&Mode=" & strPS & "&ID=<%=aid%>&TargetDate=<%=Request.Querystring("TargetDate")%>&PSC=" & strPSC & "&CCMask=" & booCCMask & "&dids=" & ddo.getSelectedIDs()
		document.all("frmAddEditAppointment").Target = strTarget
		document.all("frmAddEditAppointment").Submit()
	end function
	
	function AutoFormat(strString, strKey, strReplacer)
	    Dim intKeyPos
	    If IsNull(strString) Then
	        strString = ""
	    End If
	    intKeyPos = InStr(1, strString, strKey)
	    Do Until intKeyPos = 0
	        strString = Mid(strString, 1, intKeyPos - 1) & strReplacer & Mid(strString, intKeyPos + Len(strKey))
	        intKeyPos = InStr(1, strString, strKey)
	    Loop
	    AutoFormat = strString
	end function

	'this is the function that works off and onn checking the time on start/end

Sub window_onbeforeunload
	<%if copyTask = "True" then%>
		if booDelete = true then
			document.frames("ifLogin").location = "DeleteApptConfirm.asp?ID=<%=aid%>&Cancel=True"
		end if
	<%end if%>
End Sub

Function rrCheck( o )
	dim retVal
	retVal = true
	select case o.id
		case "chkRollover"
			if document.all("linkRec").innerText = "Edit Recurrence..." then
				if o.checked then
					window.event.returnValue = false
					msgbox "You cannot have a Rollover task if a Recurrence is Setup." & vbcrlf & "If you want to make this a Rollover task, you must edit the Recurrence and remove it.",vbOKOnly+vbExclamation,"Recurrence Detected"
					retVal = false
				end if
			else
				' remarked out because Keith does not want to see it here
				'if o.checked then
				'	document.all("divDateAdded").style.visibility = "visible"
				'else
				'	document.all("divDateAdded").style.visibility = "hidden"
				'end if
			end if
		case "linkrec"
			if document.all("chkRollover").checked then
				msgbox "You cannot have a Recurrence task if Rollover is checked." & vbcrlf & "If you want to make this a Recurrence task, you must uncheck the Rollover checkbox first.",vbOKOnly+vbExclamation,"Recurrence Detected"
				retVal = false
			end if
	end select
	rrCheck = retVal
end function

<%
'Sub ShowTask(param)
'	
'	msgbox "test" 
'	If param = "" Then 	
'		tmpStr = "Appointment.asp?TargetDate=" & window.calObj.getVal() & "&ID=0&Hour=" & FormatDateTime("12:00",4)
'	Else
'		tmpStr = param
'	End If
'	
'		w = 736
'		h = 494
'		
'	wtop = (screen.availHeight - h) / 2
'	wleft = (screen.availWidth - w) / 2
 '
'	param = "Top=" & wtop & ", Left=" & wleft & " ,Height=" & h & ", Width=" & w
'	
'	If Instr(tmpStr,"RecID") > 1 Then
'		xx = window.showModalDialog("RecurrenceDialog.asp",,"dialogheight: 120px; dialogwidth: 200px; status: no; center: yes; scroll: no")
'		If xx > 0 Then
'		
'				' RecEdit Value 1 open this instance 2- open Series
'				tmpStr = tmpStr & "&RecEdit=" & xx
'				xx = window.open (tmpStr, "",param ,null)
'		End If
'	Else
'		'if dialog <> "" then
'		w = "746"
'		h = "520"
'		wtop = (screen.availHeight - h) / 2
'		wleft = (screen.availWidth - w) / 2
'		param = "dialogHeight:" + h + "px;dialogWidth:" + w + "px;scroll:no;center:yes;status:no;"
'		xx = window.showModalDialog(tmpStr, window,param ,null)
'	End If
'End Sub
%>
</script>

<title>Add/Edit Appointment</title>
</head>

<script language="javascript">
var pubStartTime, pubEndTime, pubNoteOnly

<% If remote.Session("AvailWidth") > 800 Then %>
	var searchDialog = new browseSelect(705,968, "", "ID|Location Name" )
<% Else %>
	var searchDialog = new browseSelect(540,840, "", "ID|Location Name" )
<% End If %>
	

function checkReqField (field,setting)
{
	try 
	{
		if ((field.toString().length==0 || field == 0) && setting)
			return false
		else
			return true
	}
	catch (e) { return false }
}

var loaded=false;
var dList = '';

var isSSorOT = false; // To set the flag wheteher the APP is and SS or OT App

function j_onload()
{
	loaded = true
	checkddo(true);
	booToFocus = 1;
	
	<%If booNewRec and intDefaultAction > 0 Then
		response.write "applydeptList();" & vbcrlf
	end if
	
	if strGuestID <> "" then
		response.write "document.all('txtGuestID').value = '" & strGuestID & "';" & vbcrlf
		response.write "lookupGuestID(1,true);" & vbcrlf
		response.write "document.all('txtGuestID').value = '" & strDisplayID & "';" & vbcrlf
	end if%>
	
	document.frames("frameTaskNotes").checkDisableSS();
	document.frames("frameTaskNotes").checkDisableOT();
	
}

function setDPTS(list)
{
	dList = list;
	applydeptList();
}


function applydeptList()
{
	if (loaded)
		ddo.setDepartments(dList);
}


searchDialog.nil = true;
searchDialog.title = 'Location Selection';
searchDialog.label = 'Location:';

var opendiv = null;

function showdiv(d)
{

if (opendiv!=null) { hidediv(opendiv); }

opendiv = d;
d.style.zIndex = 10;
d.style.display = 'inline';

}

function hidediv (d)
{
d.style.zIndex = -10;
d.style.display = 'none';
opendiv=null;

}


function processkey()
{

if ((event.keyCode==9)  || (event.keyCode == 27))
	{
		if (opendiv) hidediv(opendiv);	
	}
	if(window.event.keyCode == 13 || window.event.keyCode == 32)
		lastElement = window.event.srcElement;

}

function buildTaskNotes()
{
	// get Task Notes into Hidden field
	document.all("txtTaskNotes").value =document.frames("frameTaskNotes").buildTaskNotes();
}

	function frmAddEditAppointment_onsubmit()
	{
		
		var booRetVal = false;
		booDelete = false;
		document.all("cboAction").disabled = false;
			
		if(validateForm())
		{
				
			booRetVal = true
				
			document.all('CCNumber').value = document.all('txtNumber').value;
				
			document.all("txtSalutation").value = document.all("pvSalutation").value
				
			document.all("txtLocPhone").value = formatPhone(document.all("txtLocPhone").value)
				
			document.all("txtSubjectSave").value = document.all("txtSubject").value
				
			buildTaskNotes();

			// department id's
			document.all("txtDDID").value = ddo.getSelectedIDs();
				
				
			if(!booFromFrame)
			{
				if(window.opener.Calobj)
					if(document.all("txtReminder").value == "")
						window.opener.Calobj.setDate(d1obj.getVal())
			}
			else
				parent.window.returnValue = "Refresh";
		}
			
		return (booRetVal)
			
	}

	function getValues(sql)
	{
		try {
		var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
			xmlHttp.open( "POST" , "KeyWordOptions.asp?sql=" + escape(sql), false)
			xmlHttp.send()
			return xmlHttp.responseText
			xmlHttp = null
		} catch (e) { }
	}
	
	function validateFields()
	{
	
			var retVal = true

				isANote = (document.all("txtRoom").value.length == valLen &&
							document.all ("txtLocation").value.length == 0 &&
							document.all("cboActionType").selectedIndex == 0 &&
							document.all("cboAction").selectedIndex == 0 ) ||
							document.all("chkNote").checked
	
			if (!isSSorOT && ! isANote)
				{
						var msg='';
						var reqFields = new Array();
				
						var sql = ' select fieldid,required from tblCompanyFields where CompanyID=<%=cid%>';
						var r = getValues(sql);	
	
						var arr= r.split('||')
	
							for (i=0;i<arr.length - 1;i++) // Create the Required Fields array
							{
								ar = arr[i].split('|');
								
								if (ar[1]=='True')
								{
									reqFields[i] = true;
								}
								else
									reqFields[i] = false;
							}
							
							<%if not booGuestProfile then%>
								if (!checkReqField(document.all("txtRoom").value,reqFields[0]))
									msg += 'Room#\n';
								valLen = 0;
							<%else%>
								valLen = document.all("txtRoom").value.length
							<%end if%>
							
							//if (!checkReqField(document.all("pvSalutation").selectedIndex,reqFields[1]))
							if (!checkReqField(document.all("pvSalutation").value,reqFields[1]))
									msg += 'Salutation\n';
							if (!checkReqField(document.all("txtGuestLastName").value,reqFields[2]))
									msg += 'Guest Last Name\n';
							if (!checkReqField(document.all("txtGuestFirstName").value,reqFields[3]))
									msg += 'Guest First Name\n';
							if (!checkReqField(document.all("txtGuestPhone").value,reqFields[4]))
									msg += 'Guest Phone\n';
							if (!checkReqField(document.all("cboActionType").selectedIndex,reqFields[5]))
									msg += 'Action Type\n';
							if (!checkReqField(document.all("cboAction").selectedIndex,reqFields[6]))
									msg += 'Task Type\n';
							if (!checkReqField(document.all("txtGuestEMail").value,reqFields[7]))
									msg += 'Guest E-Mail\n';
							if (!checkReqField(document.all("cboChargeTo").value,reqFields[8]))
									msg += 'Charge Type\n';
							if (!checkReqField(document.all("txtAmount").value,reqFields[9]))
									msg += 'Amount\n';
							if (!checkReqField(document.all("txtNumber").value,reqFields[10]))
									msg += 'Credit Card Number\n';
							
							msg += document.frames("frameTaskNotes").validate();
				
							if (msg.length > 1)
							{
								retVal = false;
								alert ('The Following Fields Must Be Filled In\n      Before You Can Save A Task!\n\n' + msg)
								        
							}
				
				}
					
				return retVal;
	}
	
	function validateForm()
	{
		var retval = false
		<%if booGuestProfile then%>
			valLen = document.all("txtRoom").value.length;
			/*if(GPLookup()=='EOF')
				alert("no match.  add new?");
			else
				{
				alert("the guest is not in the database.  Please select one.")
				x = openGPSearch();
				alert(x);
				}
			*/
		<%else%>
			valLen = 0;
		<%end if%>
		if( document.all("frmAddEditAppointment").action.indexOf("AppointmentConfirm.asp") > -1 && document.all("txtRoom").value.length == valLen && document.all("txtSalutation").value.length == 0 && document.all("txtGuestLastName").value.length == 0 && document.all("txtGuestFirstName").value.length == 0 && document.all("cboActionType").value == 0 && document.all("cboAction").value == 0 && document.all("txtLocation").value.length == 0 && document.all("txtLocPhone").value.length == 0 && document.all("cboChargeTo").value == 0 && document.all("txtNumber").value.length == 0 && document.all("CCExp").value.length == 0 && document.all("txtAmount").value.length == 0 && document.all("txtSubject").value.length == 0 ) //document.all("txtMI").value.length == 0 && 
			alert("You may not submit a blank Task.  Please enter some data or hit Cancel.") 
		else
			{ 
			// last line of defense against blank location id (should never happen)...
			// can't even test it since it should never happen...
			if((document.all("txtLocationID").value == "0" || document.all("txtLocationID").value == "") && (document.all("txtVendor").value != "" || document.all("txtLocation").value != ""))
				vendorLookup(false);
			////
			else
			{
				if( document.all("txtGuestEMail").value.indexOf("@") == -1 && document.all("txtGuestEMail").value.length > 0 )
				{
					alert("The Guest E-Mail Address is invalid.  Please re-enter.");
					document.all("txtGuestEMail").focus();
				}
				else
						{
							var dt = new Date(d1obj.getVal());
							if (dt=='NaN') 
								alert('Please select a valid start date for the task!');
							else
							{
								if(ddo.getSelectedIDs().length == 0)
									alert('At least one department must be selected to save a task.');
								else
									if(dateTimeCheck())
										retval = true;
							} 
	
						}
				 }
			}
			
			retval = (retval && validateFields())
			
		if(!retval)
			enableButtons();
			
		return (retval)
	}
	
	function validateAmount()
	{
	document.all("txtAmount").value = document.all("txtAmount").value.replace(/\$/g,"").replace(/ /g,"");
	if(document.all("txtAmount").value.match("[^0-9.-]"))
		{
			alert('Invalid characters in amount field.')
			document.all("txtAmount").focus();
		} else {
			if(document.all("txtAmount").value.match("-") && document.all("txtAmount").value.lastIndexOf("-") > 0)
				{
				alert('Negative number not valid.  Please edit.');
				document.all("txtAmount").focus();
				}
		}
	}
	
	function validatePhone()
	{
		if(window.event.keyCode == 222) //single quote
			window.event.returnValue = false;
	}

	function validateElement() {
	if(document.all("txtVendor").value)
	{
		if(document.all("txtVendor").value == '')
			clearVendorFields();
		if(lastElement.id == "txtVendor") {
			if(lastElement.value != "") {
				var cid = window.event.srcElement.parentElement.parentElement.id;
				if( cid != "tdElipse" )
					{
					if(window.event.srcElement.id != "cmdCancel" && window.event.srcElement.id != "cmdDelete")
						vendorLookup(false);
					}
				}
			}
		lastElement = window.event.srcElement;
	}
	}

function dateTimeCheck()
{
	var booProceed = true;
	var curDateTime = new Date("<%=dateadd("n",-15,now()) + remote.Session("TimeZone")%>");
	var startDateTime = new Date(document.all("pvDateStart").value+' '+document.all("txtStartTime").value)
	
	if(<%=aid%> == 0)
		if( startDateTime < curDateTime && !(document.all("chkNoTime").checked))
			booProceed = confirm("The date/time selected ocurrs in the past.  Are you sure you want to create it?")
	return(booProceed)
}
</script>

<body onclick="checkddo()" onmousedown="validateElement()" onkeyup="processkey()" bgcolor="#FAD667" topmargin="5" leftmargin="0" marginwidth="0" marginheight="0" link="black" vlink="black" alink="black">
<!--#include file = "Header.inc" -->
<!--#include file=Global.asp -->

<!--div id="divGPLookup" style="z-index:100;display:none;border-style:outset;border-width:2px;position:absolute;top:33px;left:1px">
	<iframe scrolling="no" src="GuestProfileSetup.asp?mode=Appointment&amp;load=1&amp;GPSearchID=<%=gpsid%>" style="height:388px;width:730px" frameborder="no" id="frameGPLookup" name="frameGPLookup"></iframe>
</div-->

<form name=taskNotes id=taskNotes action="" method="POST" target=frmTaskNotes">
<input type=hidden name=test>
</form>

<object style="height:1px;width:1px" id="reportViewer" classid="clsid:8569D715-FF88-44BA-8D1D-AD3E59543DDE" VIEWASTEXT  codebase="../arview2.cab#version=2,0,0,1214"></object>
<iframe src="LoadingAppointments.asp" id="ifLogin" name="ifLogin" frameBorder="0" style="HEIGHT: 1px; width:1px; VISIBILITY: visible"></iframe>

<table valign="top" id="tblTemp" align="center" cellspacing="0" cellpadding="0" border="0" bgcolor="#FAD667">
  <tr>
    <td width="685">
			<div id="divAppointment" style="z-index:10;visibility: visible">
			<form action method="post" id="frmAddEditAppointment" name="frmAddEditAppointment" target="ifLogin" onsubmit="return frmAddEditAppointment_onsubmit();">
			<input type="hidden" name="pageid" id="pageid" value="app">

			<input type="hidden" name="txtTaskDetail" id="txtTaskDetail">
			<input type="hidden" id="txtStateChangedToClosed" name="txtStateChangedToClosed" value="nochange">
			<!--input type=hidden id=txtPasswordSave name=txtPasswordSave-->
			<input type="hidden" id="txtSubjectSave" name="txtSubjectSave">
			<input type="hidden" id="txtTaskNotes" name="txtTaskNotes">
			<input type="hidden" name="txtDDID" id="txtDDID" value="<%=fuddid%>">
			<%
			If aid = 0 Then
				Response.Write "<input type=""hidden"" id=""txtAppointmentID"" name=""txtAppointmentID"" value=""0"">" & vbcrlf
			else
				Response.Write "<input type=""hidden"" id=""txtAppointmentID"" name=""txtAppointmentID"" value=""" & CheckNewMode(booNewRec,aid) & """>" & vbcrlf
			end if

				Dim strCreateUserID
				Dim strCreateUserName
				
				Dim strLastEditUserID
				Dim strLastEditUserName

				Dim strClosedUserID
				Dim strClosedUserName

				
				If Not booNewRec Then 
					strCreateUserID = rsSQL.Fields("CreateUserID")
					
					
					if strCopyTask = "True" then
						strLastEditUserID = 0
					else
						strLastEditUserID = cInt(N2Z(rsSQL.Fields("EditUserID").Value,0))
					end if
					
					strClosedUserID = cInt(N2Z(rsSQL.Fields("ClosedUserID").Value,0))
				else
					strCreateUserID = fuid 'remote.Session("UserID")
					strLastEditUserID = 0 'remote.Session("UserID")
					strClosedUserID = 0
				end if
				
				Set rsUser = Server.CreateObject("ADODB.Recordset")
				Set rsEditUser = Server.CreateObject("ADODB.Recordset")
				Set rsClosedUser = Server.CreateObject("ADODB.Recordset")
			  
				' Create UserName
			    Set rsUser = cnSQL.Execute("SELECT UserName, UserLName FROM tblUser WHERE UserID = " & strCreateUserID)
				if Not rsUser.EOF then
					strCreateUserName = rsUser.Fields("UserName") & " " & rsUser.Fields("UserLName") 
				end if 
				
					
					
				
				rsUser.Close 
				set rsUser = nothing

				Response.Write "<input type=hidden id=txtCreateUserID name=txtCreateUserID value=" & strCreateUserID & ">"

				' Edit UserName	
			    Set rsEditUser = cnSQL.Execute("SELECT UserName, UserLName FROM tblUser WHERE UserID = " & strLastEditUserID)
				if  Not rsEditUser.EOF then
					strLastEditUserName = rsEditUser.Fields("UserName") & " " & rsEditUser.Fields("UserLName")
				end if 
				rsEditUser.Close
				set rsEditUser = nothing
				
				' Closed UserName	
			    Set rsClosedUser = cnSQL.Execute("SELECT UserName, UserLName FROM tblUser WHERE UserID = " & strClosedUserID)
				if  Not rsClosedUser.EOF then
					strClosedUserName = rsClosedUser.Fields("UserName") & " " & rsClosedUser.Fields("UserLName")  
				end if 
				rsClosedUser.Close
				set rsClosedUser = nothing
			%>
			<input id="txtClosedUserID" name="txtClosedUserID" type="hidden" value="<%=CheckNewMode(booNewRec,rsSQL("ClosedUserID"))%>">
			</div>
    <table valign="top" align="center" border="0" cellspacing="0" cellpadding="0" width="730" class="Label">
  <tr>
    <td colspan="8">
	<table class="Label" cellpadding="0" cellspacing="0" border="0">
	<tr>
    <td align="right" height="21"><font color="#000000">&nbsp;Created:</font></td>
    <td width="294" height="21"><input style="width: 126px" class="ShortestTxtMargin" id="lstCreateUserName" name="lstCreateUserName" value="<%=strCreateUserName%>" disabled>&nbsp;
		<% If Not booNewRec Then %>
			<input id="txtCreatedDateTime" style="width: 160px" name="txtCreatedDateTime" class="ShortTxtMargin" value="<%=CheckNewMode(booNewRec,rsSQL("CreateDateTime"))%>" disabled>
		<% Else %>
			<input id="txtCreatedDateTime" style="width: 160px" name="txtCreatedDateTime" class="ShortTxtMargin" value="<%=Now() + remote.Session("TimeZone")%>" disabled>
		<% End If %>
    </td>
    <td width="386" height="21">
		<table class="Label" cellpadding="0" cellspacing="0">
			<tr>
				<td width="144" align="right"><font color="#000000">Last Edited By:&nbsp;</font></td>
				<td><input style="width: 126px" id="lstEditUserName" name="lstEditUserName" class="MedTxtMargin" value="<%=strLastEditUserName%>" disabled>&nbsp;</td>
				<td><input id="txtEditDateTime" style="width: 160px" name="txtEditDateTime" class="ShortTxtMargin" value="<%=CheckNewMode(booNewRec,rsSQL("EditDateTime"))%>" disabled></td>
			</tr>
		</table>
	</td>
	</tr>
	</table>
	</td>
  </tr>
  <tr>
	<td colspan="8">
		<div style="overflow:hidden;height:4px; border-bottom-style:solid;border-bottom-color:gray;border-bottom-width:1px;">&nbsp;</div><!--hr-->
	</td>
  </tr>
  <tr style="padding-top:3px">
		<%' temp for testing
		if booGuestProfile then
			nbsp = "&nbsp;"
			roomlen = 56
			strDisabled = "unselectable=on readonly"
			strSalDisabled = "disabled"
			strDisabledBGColor = "padding-left:2px;background-color:#C6CDC5;"
			strDisabledBGColorE = "background-color:#E0E3DF"
		else
			roomlen = 85
			nbsp = ""
			strDisabled = ""
			strSalDisabled = ""
			strDisabledBGColor = "background-color:white"
			strDisabledBGColorE = ""
		end if


		if booUseID then
			strLabel = "ID:"
			strRoomDisplay = "none"
			strIDDisplay = "inline"
		else
			strLabel = "Room:"
			strRoomDisplay = "inline"
			strIDDisplay = "none"
		end if%>
		<td align="right" class="label"><%=strLabel%></td>
		<td height="21" style="<%=strGuestStyle%>">
			<div style="display:<%=strIDDisplay%>"><input value="<%=strDisplayID%>" onFocus="idVal=this.value;gid=document.all('txtRealGuestID').value" onChange="lookupGuestID(<%=gpsid%>)" onKeyUp="checkEnter(<%=gpsid%>)" name="txtGuestID" id="txtGuestID" style="font-family:Tahoma;font-size:11px;width:56px;background-color:beige">&nbsp;</div>
			<div style="display:<%=strRoomDisplay%>"><input onChange="RoomValidate()" onKeyUp="validLen (this,20)" name="txtRoom" id="txtRoom" style="font-family:Tahoma;font-size:11px;width:<%=roomlen%>px;background-color:white"><%=nbsp%></div>
			<!--div style="display:none"><input onChange="RoomValidate()" onKeyUp="validLen (this,20)" name="txtRoom" id="txtRoom" style="font-family:Tahoma;font-size:11px;width:85px;background-color:white"></div-->
		
		<%if booGuestProfile then%>
		<div style="display:inline;vertical-align:bottom" onclick="openGPSearch()" onmousedown="imgGroup.style.borderStyle='inset'" onmouseup="imgGroup.style.borderStyle='outset'" onmouseout="imgGroup.style.borderStyle='outset'">
			<img title="Lookup Guest Profile" id="imgGroup" style="cursor:hand;border-style:outset;border-width:1px" src="images/GuestProfile.jpg" WIDTH="21" HEIGHT="16">
		</div>
		<%end if%>
		
		<input type=hidden value="<%=strGuestID%>" name=txtRealGuestID id=txtRealGuestID>
		<input type=hidden id=txtSalutation name=txtSalutation>
		<!-- **** SALUTATIONS DROPDOWN ***** -->
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Sal:
		<%if booGuestProfile then
			Response.Write "<input unselectable=on readonly type=text class=txt2 style=width:90px;" & strDisabledBGColor & " name=pvSalutation id=pvSalutation>"
		else
			Response.Write "<select id=pvSalutation class=txt2 style=width:90px name=pvSalutation>"
			do until rsSalutations.EOF
				Response.Write "<option disabled=true value=" & escape(rsSalutations("Salutation")) & ">" & rsSalutations("Salutation") & "</option>"
				rsSalutations.MoveNext
			 loop
			 rsSalutations.Close
			 set rsSalutations = nothing
			Response.Write "</select>"
		end if%>
	</td>
    <td style="<%=strGuestStyle%>;width:45px" align="right" height="21">Last:</td>
    <td style="<%=strGuestStyle%>" height="21">
		<input <%=strDisabled%> id="txtGuestLastName" onblur="properCase(this);" name="txtGuestLastName" style="<%=strDisabledBGColor%>" class="col3" value="<%=Trim(CheckNewMode(booNewRec,rsSQL("GuestLastName")))%>">
    </td>
    <td style="<%=strGuestStyle%>" align="right" height="21" width="54">First:</td>
    <td style="<%=strGuestStyle%>" align="left" height="21">
		<table style="font-face:tahoma;font-size:11px" cellpadding="0" cellspacing="0">
			<tr>
				<td>
					<input <%=strDisabled%> style="<%=strDisabledBGColor%>" id="txtGuestFirstName" onblur="properCase(this);" name="txtGuestFirstName" class="txtX" value="<%=Trim(CheckNewMode(booNewRec,rsSQL("GuestFirstName")))%>">
				</td>
				<td width="40" align="right" height="21">Phone:</td>
				<td>
					<input <%=strDisabled%> style="<%=strDisabledBGColor%>" onmouseover="this.title=this.value" onkeydown="validatePhone();" class="txtPhone" type="text" id="txtGuestPhone" name="txtGuestPhone" value="<%=Trim(CheckNewMode(booNewRec,rsSQL("GuestPhone")))%>">
				</td>
			</tr>
		</table>
		<input type="hidden" id="lstClosedUserName" name="lstClosedUserName" value="<%=strClosedUserName%>">
		<input type="hidden" id="txtClosedDateTime" name="txtClosedDateTime" value="<%=CheckNewMode(booNewRec,rsSQL("ClosedDate"))%>">
    </td>
  </tr>

  <tr height="21">
    <td align="right">Type:</td>
    <td>
		<select id="cboActionType" name="cboActionType" class="txt3">
		<%
		Set rsActionType = Server.CreateObject("ADODB.Recordset")
		'' fill in the ActionTypes!!		
		    
		'  remarked and modified to force Arrange to top (hard coded ID 2)
		'Set rsActionType = cnSQL.Execute("SELECT 0 as ActionTypeID, '' as ActionType UNION SELECT ActionTypeID, ActionType FROM tlkpActionType ORDER BY ActionType")
		Set rsActionType = cnSQL.Execute("SELECT 0 as ActionTypeID, '' as ActionType, 'a' as myIndex UNION select 2 as ActionTypeID, 'Arrange' as ActionType, 'aa' as myIndex UNION SELECT at.ActionTypeID, at.ActionType, at.ActionType as myIndex FROM tlkpActionType at join tblCompanyActionType  cat on cat.ActionTypeID=at.ActionTypeID  where at.ActionTypeID <> 2 and cat.CompanyID=" & cid & "ORDER BY myIndex")
		        
		Do While Not rsActionType.EOF
			If booNewRec = False Then
				If rsActionType.Fields("ActionTypeID") = rsSQL("ActionTypeID") then 
					Response.Write "<OPTION selected "
				Else
					Response.Write "<OPTION "
				End If
			Else
				If rsActionType.Fields("ActionTypeID") = CLng(strDefaultAction) then 
					Response.Write "<OPTION selected "
				Else
					Response.Write "<OPTION "
				End If
			End If

			Response.Write "value=" & rsActionType.Fields("ActionTypeID") & ">" & rsActionType.Fields("ActionType") & "</Option>"
			rsActionType.MoveNext
		Loop
		rsActionType.Close 
		set rsActionType = nothing
		%>
		</select>
    </td>
    <td align="right">Task:</td>
    <td>
		<select id="cboAction" name="cboAction" class="col3"> 
		<%
		Set rsAction = Server.CreateObject("ADODB.Recordset")
		'' fill in the Actions!!		
		Set rsAction = cnSQL.Execute("SELECT 0 as ActionID, '' as Action UNION SELECT a.ActionID, a.Action FROM tlkpAction a join tlnkCompanyAction ca on a.ActionID = ca.ActionID WHERE ca.CompanyID = " & cid & " ORDER BY a.Action")
		        
		Do While Not rsAction.EOF
			If booNewRec = False Then
				if (rsAction.Fields("ActionID") = rsSQL("ActionID")) then 
					Response.write "<OPTION selected "
				else
					Response.Write "<OPTION "
				end if
			Else
					Response.Write "<OPTION "
			End If

			Response.Write "value=" & rsAction.Fields("ActionID") & ">" & rsAction.Fields("Action") & "</Option>"
			rsAction.MoveNext
		Loop
		rsAction.Close
		set rsAction = nothing
		%>
		</select>
    </td>
    <td align="right">E-Mail:<!--# People:--></td>
    <td width="51" height="21">
		<input style="<%=strDisabledBGColor%>" type="text" <%=strDisabled%> id="txtGuestEmail" onKeyUp="validLen (this,50)" name="txtGuestEMail" class="txtE" value="<%=Trim(CheckNewMode(booNewRec,rsSQL("GuestEMail")))%>">
    </td>
  </tr>
  <tr>
<td id="tdVendorLabel" align="right" height="21">Vendor:</td>
<td height="21">
	<input type="hidden" name="txtLocation" id="txtLocation">
	<input type="hidden" id="txtLocationID" name="txtLocationID">

<script Language="JavaScript">
	var idVal = ""; booForceRefresh = false;
	var gid = "<%=strGuestID%>"
	
	function openGPSearch()
	{
		var retval = window.showModalDialog("GuestProfileSetup.asp?mode=Appointment&amp;load=1&amp;GPSearchID=<%=gpsid%>",window,"dialogHeight:410px;dialogWidth:700px;center:yes;scroll:no;status:no")
		if(retval && retval != ',' && retval != '')
		{
			var a = retval.split(',');
			document.all("txtGuestID").value = a[0];
			document.all("txtRealGuestID").value = a[1];
			if(a[0]=='')
				{
				document.all("txtGuestID").value = a[1];
				lookupGuestID(1);
				//if(lookupGuestID(1))
				//	document.all("cboActionType").focus();
				
				document.all("txtGuestID").value = a[0];
				}
			else
				lookupGuestID(<%=gpsid%>);
				//if(lookupGuestID(<%=gpsid%>))
				//	document.all("cboActionType").focus();
		}
		return (retval)
	}
	
	function checkEnter( idType )
	{
		if(window.event.keyCode==13)
			document.all("pvSalutation").focus();
	}
	
	function lookupGuestID( idType, booBypass )
	{
		var retval = true, booProceed = true;
		if((document.all("txtGuestID").value != "") || booForceRefresh)
		{
			document.all("div1").visibility = false;
			booForceRefresh = false;
			idVal = document.all("txtGuestID").value
			gid = document.all("txtRealGuestID").value
			// get data from id here and check for eof too
			var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP")
			var x = xmlHTTP.open("POST","GuestProfileGetGuest.asp?mode="+idType+"&id="+window.frmAddEditAppointment.txtGuestID.value, false)
			xmlHTTP.send();
			var response = xmlHTTP.responseText;
			if(response=='EOF')
				{
				alert('This Guest ID does not exist.  Please try again.');
				document.all("txtGuestID").focus();
				retval = false;
				}
			else
			{
				//alert(response);
				//if(!booBypass)
				//{
				//	if(gid == document.all("txtRealGuestID").value && document.all("txtRealGuestID").value != '')
				//		booProceed = yesno("Update this task with selected guest information?","Guest Profile");
				//}
					
				if(booProceed)
				{
					var a = response.split("|");
					window.frmAddEditAppointment.pvSalutation.value = a[0];
					window.frmAddEditAppointment.txtGuestLastName.value = a[1];
					window.frmAddEditAppointment.txtGuestFirstName.value = a[2];
					if(a[4] != '')
						strExt = ' x'+a[4]
					else
						strExt = ''
					window.frmAddEditAppointment.txtGuestPhone.value = formatPhone(a[3])+strExt;
					window.frmAddEditAppointment.txtGuestEmail.value = a[5];
					window.frmAddEditAppointment.cboChargeTo.value = a[6];
					var x = window.frmAddEditAppointment.cboChargeTo.value;
					if(!x)
						x = '';
					else
						x = document.all("cboChargeTo").options(document.all("cboChargeTo").selectedIndex).text;
					window.frmAddEditAppointment.txtChargeTo.value = x;
					window.frmAddEditAppointment.CCNumber.value = a[7];
					window.frmAddEditAppointment.txtNumber.value = a[7];
					window.frmAddEditAppointment.CCExp.value = a[8];
					window.frmAddEditAppointment.txtRealGuestID.value = a[9];
					window.frmAddEditAppointment.cboActionType.focus();
				}
			}
			xmlHTTP = null;
			return (retval)
		}
	}
	
	function checkddo(force)
	{
		if(force || (event.srcElement.id != 'imgddo' && event.srcElement.id != 'tdddoButton'))
			{
			ddo.hide();
			document.all("divSummary").innerText = ddo.summary();
			if(document.all("tdddoButton"))
				document.all("tdddoButton").style.borderStyle = "outset";
			}
		booToFocus = 0;
	}
	
	function showRecurrence()
	{
	if(rrCheck(document.all.linkrec))
	{
		if (document.all("OTLogID").value == "0" && document.all("SSLogID").value == "0")
		{	
			var h = 385
			var w = 431
			var strUrl = 'RecurrenceFrame.asp?ApptID=' + document.all("txtAppointmentID").value+'&EndBy='+d1obj.getVal(); 
			var x = window.showModalDialog(strUrl, window, "dialogheight: " + h + "px; dialogwidth: " + w + "px; center: yes; status: no; scroll: no");
			//var x = window.open(strUrl)

			if (x=='close')
				{
					try{
						window.opener.Calobj.onDateChange (parent.d1obj.getVal(),1);
					} catch (o) {
						try{
							window.opener.document.all.submit.click()
						}
						catch (o) {}
					}

					window.close();
				}
				 
		}
	}		
	}

	var booLNP = false;
	var booFNP = false;

	function pc( o )
	{
		var str = o.value, a = str.split(/[\s\/-]/); // when we need to add more mod this
		
		o.value = "";
		for(var i=0;i < a.length;i++)
			{
			if(i > 0)
				delimiter = str.substr(o.value.length,1);
			else
				delimiter = ""
			o.value += (delimiter + a[i].substr(0,1).toUpperCase()+a[i].substr(1).toLowerCase());
			}
	}	

	
	function GPLookup()
	{
	<%if booGuestProfile then%>
		//if(document.all("txtGuestID").value != "")
		//{
			var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP")
			var x = xmlHTTP.open("POST","GuestProfileCheck.asp?last="+escape(document.all("txtGuestLastName").value)+"&first="+escape(document.all("txtGuestFirstName").value), false)
			xmlHTTP.send();
			var response = xmlHTTP.responseText;
			return (response);
		//}
	<%end if%>
	}

	function properCase( o )
	{
		if("<%=aid%>" == "0")
		{
			var str = o.value, a;
		
			switch(o.id)
			{
				case "txtGuestFirstName":
				{
					if(!booFNP && str.length > 0)
					{
						pc(o);
						booFNP = true;
					}
					if(o.value.length == 0)
						booFNP = false;
					break;
				}
				case "txtGuestLastName":
				{
					if(!booLNP && str.length > 0)
					{
						pc(o);
						booLNP = true;
					}
					if(o.value.length == 0)
						booLNP = false;
					break;
				}
			}
		}
	}	 
</script>
	<table cellspacing="0" cellpadding="0" border="0">
		<tr>
			<td>
				<input onkeydown="processElipseKeys(0)" id="txtVendor" oonChange="vendorChanged()" ooonBlur="vendorOnBlur()" class="txtV" style="height: 19px; width: 194px; overflow: hidden">
			</td>
			<td align="left">
				<table class="Label" border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td align="center" bgcolor="silver">
							<div onkeydown="processElipseKeys(1)" onselectstart="window.event.returnValue=false" onclick="vendorLookup(true)" id="divElipse" style="cursor:hand" onfocus="setBorder('','on')" onblur="setBorder('','off')">
								<table width="26px" cellpadding="1" cellspacing="0" bgcolor="silver" class="Label"><tr><td id="tdElipse" style="border-color: white; border-width: 2px; border-style: outset" valign="middle" onmousedown="this.style.borderStyle='inset'" onmouseup="this.style.borderStyle='outset';" onmouseout="this.style.borderStyle='outset';"><center><strong>. . .</strong></center></td></tr></table>
							</div>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

    </td>
    <td align="right" height="21">Phone:</td>
    <td>
		<input class="col3" onKeyUp="validLen (this,50)" id="txtLocPhone" name="txtLocPhone">
    </td>
    <td align="right">Street:</td>
    <td>
		<input id="txtLocAddress" name="Address" onKeyUp="validLen (this,256)" class="txtE" value="<%=CheckNewMode(booNewRec,rsSQL("locAddress"))%>">
    </td>
    <td>
		<input style="visibility:hidden;width:0px" id="txtLocCity" name="LocCity" class="col5" value="<%=CheckNewMode(booNewRec,rsSQL("locCity"))%>">
		<input type="hidden" id="txtLocState" name="LocState" value="<%=CheckNewMode(booNewRec,rsSQL("LocState"))%>">
    </td>
  </tr>
  <tr style="overflow:hidden;height:21px">
  <td>&nbsp;</td>
 <td id="OTButton" align="left"></td>
    <td align="right">Charge:</td>
    <td>
    <table cellpadding="0" cellspacing="0" border="0" style="font-face: tahoma; font-size: 11;">
    <tr>
    <td>
		<%if booGuestProfile then
			strDisplaytxt = "inline"
			strDisplaylst = "none"
		else
			strDisplaytxt = "none"
			strDisplaylst = "inline"
		end if%>		
			<input <%=strDisabled%> type=text style="display:<%=strDisplaytxt%>;<%=strDisabledBGColor%> ;font-face: tahoma; font-size: 11; width: 80px" valign="middle" id="txtChargeTo" name="txtChargeTo">
    		<select <%=strDisabled%> style="display:<%=strDisplaylst%>;<%=strDisabledBGColor%>;font-face: tahoma; font-size: 11; width: 80px" valign="middle" id="cboChargeTo" name="cboChargeTo">
				<option value="0"> </option>
				<%dim rsChargeType
				Set rsChargeType = Server.CreateObject("ADODB.Recordset")
				set rsChargeType = cnSQL.Execute("select * from tlkpChargeType order by DisplayOrder")
				do until rsChargeType.EOF
					Response.Write "<option value=" & rsChargeType.Fields("ChargeTypeID").Value & ">" & rsChargeType.Fields("ChargeType").Value & "</option>" & vbcrlf
					rsChargeType.MoveNext
				loop
				rsChargeType.Close
				set rsChargeType = nothing
				%>
			</select>
	</td>
	<td>&nbsp;&nbsp;Amt:&nbsp;</td>
    <td>
		<input style="width:43px" onblur="validateAmount()" onpaste="window.event.returnValue=false;" valign="middle" id="txtAmount" name="txtAmount" class="col4half" value="<%=Trim(CheckNewMode(booNewRec,rsSQL("Amount")))%>">
	</td>	
	<td></td>
	</tr>	
    </table>
    </td>
    <!-- <td width="35" align="right" height="21">Amt:</td> -->
    <td align="right">Charge#:</td>
    <td style="padding-right:0px" align="left">
		<table border="0" align="left" cellpadding="0" cellspacing="0" style="font-face: tahoma; font-size: 11px">
			<tr>
				<td>
					<input <%=strDisabled%> language="javascript" onkeyup="validLen(this,30)" class="col4half" style="<%=strDisabledBGColor%>;width: 137px" valign="left" id="txtNumber" name="txtNumber" value="<% If remote.Session("FloatingUser_VCCN") or (CheckNewMode(booNewRec,rsSQL("CreateUserID")) = fuid) Then Response.write (Trim(CheckNewMode(booNewRec,rsSQL("CCNumber")))) Else Response.Write (CheckNewMode(booNewRec,ccn(CheckNewMode(booNewRec,rsSQL("CCNumber"))))) %>">
				</td>
				<td align="right" width="30px" style="padding-right:0px">
					<input type="hidden" id="CCNumber" name="CCNumber" value="<%=Trim(CheckNewMode(booNewRec,rsSQL("CCNumber")))%>">
					Exp:
				</td>
				<td align="right">
					<input <%=strDisabled%> valign="middle" onKeyUp="validLen (this,16)" id="CCExp" name="CCExp" class="exp" style="<%=strDisabledBGColor%>" value="<%=Trim(CheckNewMode(booNewRec,rsSQL("CCExp")))%>">
				</td>	
			</tr>
		</table>
  </td>
  </tr>
  <tr>
    <td style="padding-top:3px" valign="top" align="right" height="21">Detail:</td>
    <td width="614" height="63" rowspan="3" valign="top" colspan="9">
		<table cellpadding="0" cellspacing="0" style="margin-top: 3px">
			<tr>
				<%if rsSQL.EOF then
					strAppointmentID = "0"
				else
					strAppointmentID = rsSQL("AppointmentID")
				end if%>
				<td><iframe onfocus="checkddo()" onactivate="validateElement()" id="frameTaskNotes" name="frameTaskNotes" src="AppointmentTaskNotes.asp?AppointmentID=<%=strAppointmentID%>&amp;rnd=<%=intRndNum%>&amp;Mode=1" class="TallLongTxt"></iframe></td>
				<td align="right" valign="top" class="Label">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Notes:&nbsp;</td>
				<td><textarea class="TallMedTxt" style="overflow-y:auto" id="txtSubject" name="txtSubject"><%=CheckNewMode(booNewRec,rsSQL("ApptText"))%></textarea></td>
			</tr>
		</table>
    </td>
  </tr>
  <tr>
    <td valign="bottom" align="center" height="21">
		<%if copyTask = "True" then
			Response.Write "<IMG SRC=""images/copy.gif"">" & vbcrlf
		end if%>
	</td>
  </tr>
  <tr>
    <td align="right" height="21"></td>
  </tr>
  <tr>
	<td colspan="8">
		<div style="overflow:hidden;height:4px; border-bottom-style:solid;border-bottom-color:gray;border-bottom-width:1px;">&nbsp;</div><!--hr-->
	</td>
  </tr>
  <tr style="padding-top:3px">

	<td colspan="8">
		<table cellpadding="0" cellspacing="0">
			<tr>
				<td width="70" align="center">
					<img style="border-raised: true" SRC="images/clockYellow.gif" WIDTH="31" HEIGHT="37">
				</td>
				<td>
					<table cellpadding="0" cellspacing="0" class="Label">
						<tr>
							<td style="width:56px">Start Time:</td>
							<td style="width:110px">
								<div style="position:absolute;top:410;left:234;display:none;" name="div1" id="div1">
									<iframe tabindex="-1" allowTransparency="yes" id="frm1" name="frm1" FrameBorder="no" Scrolling="no" style="height:132;width:164" src="dayview.asp" onBlur="d1obj.onBlur()"></iframe>
								</div>

								<input type="hidden" id="pvDateStart" name="pvDateStart">
								<input type="hidden" id="pvDateEnd" name="pvDateEnd">

								<script Language="JavaScript">

								var d1obj = new ddDropDown (85,14,'d1',132,164);
								d1obj.enable();

								d1obj.process = function (str)
								{
									d1obj.setVal(str,1);
									document.all("pvDateStart").value = str;
									document.all("pvDateEnd").value = str;
									<%if aid = 0 then%>
										document.all("txtDateAdded").value = str;
									<%end if%>
									
									hidediv(div1);
									
									if (document.all("OTLogID").value > 0)
									{
											var x = calert ('Would you like to change your Open Table reservation?');
											if (x==1) MakeOTReservation(document.all("OTLogID").value);
									}

								
								if (document.all("SSLogID").value > 0)
									{
											var x = calert ('Would you like to change your Super Shuttle reservation?');
											if (x==1) MakeSSReservation(document.all("SSLogID").value);
									}
								
										
								} 

								d1obj.onBlur = function ()
									{
									}
								
								
								d1obj.blur = function()
								{
								var gv = d1obj.getVal();
								if(!validDate(gv))
									{
									var dx = Date();
									d1obj.setVal(dx,1)
									hidediv(div1)
									document.all.d1.focus();
									alert("Please enter a valid date.");
									}
								else
									d1obj.setVal(gv,1)
								document.all("pvDateStart").value = gv; //(dx.getMonth()+1)+"/"+dx.getDate()+"/"+dx.getFullYear();
								document.all("pvDateEnd").value = gv;
								div1.style.top = 244 //this.posY;
								div1.style.left = this.posX-3;
								frames["frm1"].cobj.setDate(gv)
								hidediv(div1)
								}
									
								d1obj.Click = function ()
								{
									div1.style.top = 244 //this.posY;
									div1.style.left = this.posX-3;

									if (div1.style.display == 'none') 
										{ 
											frames["frm1"].cobj.setDate(d1obj.getVal())
											
											showdiv(div1)}
									else
										{ hidediv(div1) };

								}


								d1obj.textChange = function () 
								{
									//frames["frm1"].cobj.setDate(d1obj.getVal());
								};

								d1obj.keyUp = function () {
															if (d1obj.posX) showdiv(div1)
														  }
								</script>
							</td>
							<td style="width:234px">
								<select onchange="startTimeChange(this)" style="font-family:Tahoma;font-size:11px;background-color:white" name="txtStartTime" id="txtStartTime">
								</select>
	
								<%if request.querystring("RecID") <> "" then
									'Response.Write "<input type=button style=""height:20px;width:90px;font-family:tahoma;font-size:11px"" id=cmdRecEdit value=""Edit Recurrence"" onClick=""showRecurrence()"">" & vbcrlf
									Response.Write "&nbsp<a id=linkrec href=""javascript:showRecurrence()"">Edit Recurrence...</a>" & vbcrlf
								else
								'	Response.Write "<input type=button style=""height:20px;width:90px;font-family:tahoma;font-size:11px"" id=cmdRecEdit value=""Add Recurrence"" onClick=""showRecurrence()"">" & vbcrlf
									Response.Write "&nbsp<a id=linkrec href=""javascript:showRecurrence()"">Add Recurrence...</a>" & vbcrlf
								end if%>

								<script Language="JavaScript1.2">
								<!--#INCLUDE file="arrayTimes.asp"-->
								</script>

								<script Language="JavaScript">

								function startTimeChange(o)
								{
									if(document.all("chkNoTime").checked)
										{
										document.all("chkNoTime").checked = false;
										checkNote(false);
										}

									document.all("txtEndTime").selectedIndex = o.selectedIndex;
									
									if (document.all("OtLogID").value > 0)
										{
											var x = calert ('Would you like to change your Open Table reservation?');
											if (x==1) MakeOTReservation(document.all("OTLogID").value);
										}
										
									if (document.all("SSLogID").value > 0)
										{
											var x = calert ('Would you like to change your Super Shuttle reservation?');
											if (x==1) MakeSSReservation(document.all("SSLogID").value);
										}
									
										
										
								}


								for (i=0;i<at.length;i++)
								{
								var t = document.createElement("OPTION")
								t.text = at[i];
								t.value = at[i];

								document.all("txtStartTime").options.add (t)

								}
								
								function checkNote( booRevertTime )
								{
									if(document.all("chkNoTime").checked)
										{
										pubNoteOnly = document.all("chkNote").checked
										document.all("chkNote").checked = true;
										document.all("chkNoteHidden").value = document.all("chkNote").value;
										document.all("chkNote").disabled = true;
										pubStartTime = document.all("txtStartTime").value
										pubEndTime = document.all("txtEndTime").value
										document.all("txtStartTime").value = "";
										document.all("txtEndTime").value = "";
										}
									else
										{
										document.all("chkNote").disabled = false;
										if( booRevertTime )
											{
											if(!pubStartTime)  // no history yet
												{
												document.all("txtStartTime").value = "12:00 PM"
												document.all("txtEndTime").value = "12:00 PM"
												}
											else
												{
												document.all("txtStartTime").value = pubStartTime
												document.all("txtEndTime").value = pubEndTime
												}
											}
										document.all("chkNote").checked = pubNoteOnly
										document.all("chkNoteHidden").value = pubNoteOnly;
										}
								}
								</script>

								<input type="hidden" id="txtTimeDelta" name="txtTimeDelta" value="0">
							</td>
							<td align="center" rrowspan="2" style="width:116px">
								<div title="Task rolls over to the next day at midnight if not closed">&nbsp;&nbsp;
								<input type="checkbox" id="chkRollover" name="chkRollover" language="vbscript" onclick="rrCheck(document.all.chkRollover)">&nbsp;Rollover task</div>
								<div style="display:none;visibility:hidden" id="divDateAdded"><table id="tblDateAdded" cellpadding="2" cellspacing="0" style="border-style:inset;border-width:1px"><tr><td class="label" align="center" style="background-color:#FAE667;height:38px;width:108px">Original Task Date<br><input type="text" onkeydown="window.event.returnValue=false" class="txtDateAdded" style="text-align:center;width:100px" id="txtDateAdded" name="txtDateAdded"></td></tr></table></div>
							</td>
							<td>
								<p align="left">
									<% If booNewRec Then %>
										<input language="javascript" onchange="saveNoteVal()" id="chkNote" name="chkNote" type="checkbox"><label class="label">&nbsp;Note Only (not a task)</label>
										<input style="display:none" type="text" id="chkNoteHidden" name="chkNoteHidden" value="<%if request.querystring("NoTime") = "True" then Response.Write "on" end if%>">
									<% Else %>
										<input language="javascript" onchange="saveNoteVal()" id="chkNote" name="chkNote" type="checkbox" <%if rsSQL("Note") = true then Response.Write "CHECKED" end if%>>&nbsp;<label class="label">Note Only (not a task)</label>
										<input style="display:none" type="text" id="chkNoteHidden" name="chkNoteHidden" value="<%if rsSQL("Note") = true then Response.Write "on" end if%>">
									<% End If %>
								</p>
							</td>
							<td>
							</td>
						</tr>
						<tr>
							<td>End Time:</td>
							<td>
							</td>
							<td>
								<select onchange="endTimeChange(this)" style="font-family:Tahoma;font-size:11px;background-color:white" name="txtEndTime" id="txtEndTime">
								</select>
								<input onclick="checkNote(true)" type="checkbox" id="chkNoTime" name="chkNoTime">&nbsp;No Specific Time
							</td>
							<td align="Left" rrowspan="2" style="width:116px">
							<div title="Task will show up over multiple time slots if checked">
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" id="chkSpan" name="chkSpan" <%=strSpanDisabled%> language="vbscript" oonclick="rrCheck(document.all.chkRollover)">&nbsp;Span</div>
							</td>
							<td>
								<p align="left">
									<%strEMailConfirm = ""
									If booNewRec Then
										strEMailConfirm = ""
									Else
										if rsSQL("EMailConfirm") then
											strEMailConfirm = "checked"
										end if
									End If
									%><input type="checkbox" id="chkEMail" name="chkEMail" <%=strEMailConfirm%>><font color="#000000">&nbsp;E-Mail Confirmation</font>
								</p>
							</td>
							<td>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</td>
  <tr>

<script Language="JavaScript">

	function endTimeChange(o)
	{
		
		if (o.selectedIndex < document.all("txtStartTime").selectedIndex)
			{
				alert ("End time can not be less than start time.");
				o.selectedIndex = document.all("txtStartTime").selectedIndex;
			}
		else
			{
			if(document.all("chkNoTime").checked)
				{
				document.all("chkNoTime").checked = false;
				checkNote(false);
				document.all("txtStartTime").value = document.all("txtEndTime").value;
				}
			}
			
			if (o.selectedIndex > (document.all("txtStartTime").selectedIndex + 3))
					document.all("chkSpan").checked = true;
			else
				document.all("chkSpan").checked = false;
	}

	for (i=0;i<at.length;i++)
	{
	var t = document.createElement("OPTION")
	t.text = at[i];
	t.value = at[i];
	document.all("txtEndTime").options.add (t)
	}

function setDDID(o)
{
	if(o.checked)
		document.all("txtDDID").value = 0;
	else
		document.all("txtDDID").value = <%=remote.session("FloatingUser_DDID")%>;
		
}
	function clearVendorFields()
	{	
		document.all("txtLocationID").value = 0;
		document.all("printloc").disabled = true;

		document.all("txtVendor").value = '';
		document.all("txtLocation").value = '';
		searchDialog.defaultValue = '';
		document.all("txtLocPhone").value = '';
		document.all("txtLocAddress").value = '';
		document.all("txtLocCity").value = '';
	}

	function vendorLookup(forceIt)
	{
	
	if((document.all("txtVendor").value != document.all("txtLocation").value) || forceIt)
	{
		if(document.all("txtVendor").value == '' && !forceIt)
			clearVendorFields()
		else
		{
			window.event.returnValue = false;
			if (document.all("OTLogID").value != '0') // Check if OT Res exists.
			{
				alert('You must cancel the OT Reservation before changing the Vendor!')
			}
			else // Check for SS also
				
			if (document.all("SSLogID").value != '0') // Check if SS Res exists.
			{
				alert('You must cancel the Super Shuttle Reservation before changing the Vendor!')
			}
			else
				
				{
					if(window.event.srcElement.type == 'button' || window.event.srcElement.id == 'linkrec')  //could be more specific later
						alert("The selected vendor does not exactly match a vendor in your Vendor List.  You must select a vendor from the Vendor List or leave the field blank.\n\nClick OK to continue.")

					searchDialog.defaultValue = escape(document.all("txtVendor").value);
					searchDialog.selectedID = document.all("txtLocationID").value;
					searchDialog.Show();
						
					if(searchDialog.clear)
					{
						clearVendorFields();
						document.all("txtVendor").focus();
					}
					else
					{
						if(searchDialog.canceled)
							{
								document.all("txtVendor").value = document.all("txtLocation").value;
							}
							else
							{
								if(searchDialog.column[0])
									{
									document.all("txtLocationID").value = searchDialog.column[0];
									if (document.all("txtLocationID").value != 0)
									   document.all("printloc").disabled = false
									else
									   document.all("printloc").disabled = true
										   
											
									document.all("txtLocation").value = searchDialog.column[1].replace(/\&amp;/g,'&');
									document.all("txtVendor").value   = document.all("txtLocation").value;

									if(searchDialog.column[6])
										document.all("txtLocPhone").value = searchDialog.column[6];
									else
										document.all("txtLocPhone").value = "";
										
									document.all("txtLocAddress").value = searchDialog.column[2];
									document.all("txtLocCity").value = searchDialog.column[3];
										
									<% If remote.Session("FloatingUser_OTUserID") <> "0" and remote.Session("CompanyOTID") <> "0" Then %>
										
									if(parseInt(searchDialog.column[7])>0)
									{
										if (document.all("OTLogID").value == "0")
										{
											var cname = '<%response.write remote.Session("FloatingUser_UserName") & chr(32) & remote.Session("FloatingUser_UserLName")%>';
											var x = calert ('Do you want to make a reservation in Open Table?' )
											if (x==1) 
												MakeOTReservation(0)
											else
												{
													if (document.all("OTLogID").value == '0')
														document.all("OTButton").innerHTML = '<a href="javascript:MakeOTReservation(0)">Create OpenTable Reservation</a>';
													else
														document.all("OTButton").innerHTML = '<a href="javascript:MakeOTReservation(<%=intOTLogID%>)">Edit OpenTable reservation</a>';
												}	
										}
									}
								<% End IF %>	
								
								// SuperShuttle
								
								<% if CLng("0" & remote.Session("SS_CompanyID")) > 0 Then %>
								{
								
								
									var ss = new String(document.all("txtLocation").value);
									ss = ss.toUpperCase ();
									
									if (ss.indexOf ('SUPER',0) > -1 && ss.indexOf ('SHUTTLE',0) > -1)
									{
											var x = calert ('Would you like to make an online Super Shuttle Reservation?' )
											if (x==1) 
												MakeSSReservation(0)
											else
												document.all("OTButton").innerHTML = '<a href="javascript:MakeSSReservation(0)">Create Super Shuttle Reservation</a>';
									}
									else // if neither SS or OT
										document.all("OTButton").innerHTML = ""	
			
								
								}
								<% end if %>
										
							 
							 
							 <%
							 if booGuestProfile then
								response.write "lastElement = document.all(""txtGuestID"").parentElement" & vbcrlf
							 else
								response.write "lastElement = document.all(""txtRoom"").parentElement" & vbcrlf
							 end if
							 %>
							
							document.all("txtLocPhone").focus();
							}
						}
					}
				}
			}
		}
	}
	
function getPeopleNum()
{
	var npeople = null;
	var r = new RegExp("[0-9]+", "ig");
	while (npeople==null)
	{
		 npeople = getpeople()
		 if (npeople==null) npeople=0;
		 var a = r.exec(npeople);
		 if (a==null)
				alert ("Please enter the number of people");
		 npeople = a;	
	}
	return npeople;
}

var xml='';

function getNode(n)
{
	try {
			return xml.selectSingleNode ("/Reservation/" + n).firstChild.text;
		}
		catch (e)
		{
			return '';
		}
}
	
	function MakeOTReservation(ResID)
	
	{
				if (ResID==0)
				{
	
						if (document.all("cboAction").options(document.all("cboAction").selectedIndex).text!='Restaurant Reservation')
						{						
							for (i=0;document.all("cboAction").options(i).text!='Restaurant Reservation';i++);
							document.all("cboAction").selectedIndex = i;

							for (i=0;document.all("cboActionType").options(i).text!='Arrange';i++);
							document.all("cboActionType").selectedIndex = i;

							cboaction_onchange()
							npeople = getPeopleNum();
								
						}
						else
							if(document.frames("frameTaskNotes").document.all("txtData2").value=='')
										var	npeople = getPeopleNum();
							else
										var npeople = document.frames("frameTaskNotes").document.all("txtData2").value;								
								
							
							//document.all("cboAction").disabled = true; 
						
						var otid='';
						
						if (parseInt(npeople) > 0)
						{	
							
							try 
								{
									otid=searchDialog.column[7].toString();	
								}
							catch (e)
								{
									otid='<%=strOTID%>';
									
								}
							//alert(otid);
							
							var starttime = document.all("txtStartTime").options(document.all("txtStartTime").selectedIndex).text;
							
							var startdate = d1obj.getVal().toString();
							
							var pstr = 'date=' + startdate + '&time=' + starttime + '&otid='+otid;
							
								pstr += '&fname=' + document.all("txtGuestFirstName").value;
								pstr += '&lname=' + document.all("txtGuestLastName").value;
								pstr += '&phone=' + document.all("txtGuestPhone").value;
								pstr += '&ApptID=' + document.all("txtAppointmentID").value;
								pstr += '&people=' + npeople;
								pstr += '&OTLogID=' + document.all("OTLogID").value;
								
							
							
							var h = 440
							var w = 370
							var strUrl = 'OTMainFrame.asp?' + pstr; 

							window.document.title = 'GC - OpenTable';
							var x = window.showModalDialog(strUrl, window, "dialogheight: " + h + "px; dialogwidth: " + w + "px; center: yes; status: no; scrollbars: no");
							
							enableButtons();
							
							window.document.title = 'Add/Edit Appointment';
							if (document.all("OTLogID").value == '0')
								document.all("OTButton").innerHTML = '<a href="javascript:MakeOTReservation(0)">Create OpenTable reservation</a>';
							else
								document.all("OTButton").innerHTML = '<a href="javascript:MakeOTReservation(' + document.all("OTLogID").value + ')">Edit OpenTable reservation</a>';
									
						}
					}	
					else
					{
						
							//var otid=searchDialog.column[0].toString();
							
							var npeople = document.frames("frameTaskNotes").document.all("txtData2").value;														
							var starttime = document.all("txtStartTime").options(document.all("txtStartTime").selectedIndex).text;
							var startdate = d1obj.getVal().toString();
							var pstr = 'date=' + startdate + '&time=' + starttime ;//+ '&otid='+otid;
								pstr += '&fname=' + document.all("txtGuestFirstName").value;
								pstr += '&lname=' + document.all("txtGuestLastName").value;
								pstr += '&phone=' + document.all("txtGuestPhone").value;
								pstr += '&ApptID=' + document.all("txtAppointmentID").value;
								pstr += '&people=' + npeople;
								pstr += '&OTLogID=' + document.all("OTLogID").value;
							
							
							var h = 440
							var w = 370
							var strUrl = 'OTMainFrame.asp?' + pstr; 

							window.document.title = 'GC - OpenTable';
							var x = window.showModalDialog (strUrl, window, "dialogheight: " + h + "px; dialogwidth: " + w + "px; center: yes; status: no; scrollbars: no");
							
							enableButtons();
							
							if (document.all("OTLogID").value == '0')
								document.all("OTButton").innerHTML = '<a href="javascript:MakeOTReservation(0)">Create OpenTable reservation</a>';
							else
								document.all("OTButton").innerHTML = '<a href="javascript:MakeOTReservation(' + document.all("OTLogID").value + ')">Edit OpenTable reservation</a>';
								
							window.document.title = 'Add/Edit Appointment';								
						
					}
	}
	
	function setanddisbale(f,n)
	{
		frameTaskNotes.document.all(f).value = getNode(n);	
		frameTaskNotes.document.all(f).disabled = true;
	}
	
	function MakeSSReservation(ResID)
	
	{
		<% if CLng("0" & remote.Session("SS_CompanyID")) > 0 Then %>
				
						var npeople=1;
	
						if (document.all("cboAction").options(document.all("cboAction").selectedIndex).text!='Super Shuttle')
						{				
							for (i=0;document.all("cboAction").options(i).text!='Super Shuttle';i++);
							document.all("cboAction").selectedIndex = i;

							for (i=0;document.all("cboActionType").options(i).text!='Arrange';i++);
							document.all("cboActionType").selectedIndex = i;

							cboaction_onchange()
								
						}
						else
						{
							try
							{
								if(document.frames("frameTaskNotes").document.all("txtData2").value!='')
									npeople = document.frames("frameTaskNotes").document.all("txtData2").value;								
							}
							catch (e)
									{
									   npeople=1;
									}		
						}	
						
						<!--#include file="TaskNotesFieldIDs.asp"-->
						
							var ssid=ResID;
							
							var starttime = document.all("txtStartTime").options(document.all("txtStartTime").selectedIndex).text;
							
							var startdate = d1obj.getVal().toString();
							
							var pstr = '?Date=' + escape(startdate) + '&Time=' + escape(starttime);							
							
							pstr += '&SSResID=' + ssid;
							pstr += '&CompanyID=<%=remote.session("CompanyID")%>';
							pstr += '&UserID=<%=remote.session("FloatingUser_UserID")%>';
							pstr += '&ApptID=' + document.all("txtAppointmentID").value;							
							pstr += '&FN=' + escape(document.all("txtGuestFirstName").value);
							pstr += '&LN=' + escape(document.all("txtGuestLastName").value);
							pstr += '&Email=' + escape(document.all("txtGuestEMail").value);
							pstr += '&people=' + npeople;
							
							if (document.all("pvSalutation").value != '')
								pstr += '&Sal=' + escape(document.all("pvSalutation").value);
								
							var h = 546
							var w = 600
							var strUrl = 'supershuttle/SuperShuttleMain.aspx' + pstr; 
							
							window.document.title = 'GC - SuperShuttle';
							var x = window.showModalDialog(strUrl, window, "dialogheight: " + h + "px; dialogwidth: " + w + "px; center: yes; status: no; scrollbars: no");
							
							var idSpecInstrID = "txtData43";
							var reslength = 0;
							
							try { reslength = x.length } catch (e) { }
							
							isCancelled = (x=="Canceled");
							
							if ( reslength > 5 && !isCancelled)
							{
								
								xml = new ActiveXObject("Microsoft.xmlDOM"); 
								xml.loadXML (x);
					
								document.all("SSLogID").value = getNode("RezID");
								
								setanddisbale (idConfirmationID,"SSConfirmationCode");
								setanddisbale (idPeople,"adultsN");
								setanddisbale (idKids,"kidsN");
								setanddisbale (idAirline,"airlineCode");
								setanddisbale (idAirport,"airportCode");
					 			setanddisbale (idFlightNO,"flightNumber");
								setanddisbale (idFlightDate,"flightDate");
								setanddisbale (idFlightTime,"flightTime");
								
								frameTaskNotes.document.all(idStatus).value = "Reserved";	
								frameTaskNotes.document.all(idStatus).disabled = true;
								
								var pickupTime = getNode("selPickupTime");
								var pickupDate = getNode("selPickupDate");

								d1obj.setVal(pickupDate);
								document.all("pvDateStart").value = pickupDate;
								document.all("pvDateEnd").value = pickupDate;
	
								for (i=0; document.all("txtStartTime").options.length > i && document.all("txtStartTime").options(i).text != pickupTime; i++);
	
								document.all("txtStartTime").selectedIndex = i;
								document.all("txtEndTime").selectedIndex = i;	
	
								var cboStatus = document.all("cboStatus")
								for (i=0; cboStatus.options.length > i && cboStatus.options(i).text != 'Closed'; i++);
								cboStatus.selectedIndex = i;
								cboStatus = null
									
								frameTaskNotes.document.all(idConfirmationID).disabled = true;
								frameTaskNotes.document.all(idPeople).disabled = true;
					
								document.all("OTButton").innerHTML = '<a href="javascript:MakeSSReservation(' + getNode("RezID") + ')">Edit Super Shuttle Reservation</a>';
								
								isSSorOT = true; // Need to set this to true so the task saves without any validations
								
								document.all("cmdSaveAndClose").click();
							
							}
							else // we are processing the cancellation
							
							if (isCancelled) // only if the user Canceled the reservation specifically
								{
									
									isSSorOT = true;
									document.all("SSLogID").value = 0;
									document.all("OTButton").innerHTML = '<a href="javascript:MakeSSReservation(0)">Create Super Shuttle Reservation</a>';
									
									frameTaskNotes.document.all(idStatus).value = "Canceled";	
									frameTaskNotes.document.all(idStatus).disabled = false;
									
									document.frames("frameTaskNotes").checkDisableSS();
									
									var cboStatus = document.all("cboStatus")
									for (i=0; cboStatus.options.length > i && cboStatus.options(i).text != 'Canceled'; i++);
									cboStatus.selectedIndex = i;
									cboStatus = null
									
									document.all("cmdSaveAndClose").click();
							
								}
							
							enableButtons();
							
		<% end if %>
	}
	
	
	function enableButtons()
	{
		document.all("cmdSaveandCopy").disabled = false
		document.all("cmdSaveandPrint").disabled = false
		document.all("cmdSaveandClose").disabled = false
	}
	
	function processElipseKeys(n)
	{
		if(n==1) // Elipses
		{
			if(window.event.keyCode == 13 || window.event.keyCode == 32)
				vendorLookup(true);
		} else {  // Field
			if(window.event.keyCode == 9 || window.event.keyCode == 13) // tab & enter
				vendorLookup(false);
		}
	}

	function setBorder( i, mode )
	{
		if(mode=='on')
			document.all('tdElipse'+i).style.backgroundColor = 'lightgreen';
		else
			document.all('tdElipse'+i).style.backgroundColor = '';
	}
	
	function SetupOT()
	{
	
			<% IF Cint("0" & strOTID) <> 0 Then %>
	
					if (document.all("OTLogID").value == '0')
						document.all("OTButton").innerHTML = '<a href="javascript:MakeOTReservation(0)">Create OpenTable reservation</a>';
					else
					{
									
						document.all("OTButton").innerHTML = '<a href="javascript:MakeOTReservation(' + document.all("OTLogID").value + ')">Edit OpenTable reservation</a>';
					}	
			<% end if %>							
	
	// Limited Logic for super shuttle
	
			var isSS = false;
			
			<% if CLng("0" & remote.Session("SS_CompanyID")) > 0 Then %> // We don't want any of the bellow if not a SS Hotel
			
					if (document.all("SSLogID").value > 0) isSS = true; // Condition 1
			
					var ss = new String(document.all("txtLocation").value);
					ss = ss.toUpperCase ();
											
					if (ss.indexOf ('SUPER',0) > -1 && ss.indexOf ('SHUTTLE',0) > -1) // Condition 2
					{
						isSS = true;
					}
			
					if (document.all("cboAction").options(document.all("cboAction").selectedIndex).text == "Super Shuttle")
					{
						isSS = true;
					}
						
					if (isSS)
					{
						// isSSorOT = true;
						
						if (document.all("SSLogID").value > 0)
							document.all("OTButton").innerHTML = '<a href="javascript:MakeSSReservation(' + document.all("SSLogID").value + ')">Edit Super Shuttle reservation</a>';
						else
							document.all("OTButton").innerHTML = '<a href="javascript:MakeSSReservation(0)">Create Super Shuttle reservation</a>';
					}
					
			<% end if %> // only for SuperSHuttle HOtels
	
	}
	
	function buttonAction()
	{
		if(document.all("tdddoButton").style.borderStyle != "inset")
			document.all("tdddoButton").style.borderStyle = "inset";
		else
			document.all("tdddoButton").style.borderStyle = "outset";
				
		document.all("divSummary").innerText = ddo.summary();
		document.all("divSummary").title = document.all("divSummary").innerText;
		ddo.toggleShow();
	}

	// sk - 7/1/2003 
	// this function can be used on any field
	// it validates the passed length of the passed object,
	// then displays a message and trims it if it's over.
	function validLen( obj, intLen )
	{
		if(obj.value.length > intLen)
			{
				obj.value = obj.value.substr(0,intLen)
				alert("This field takes a maximum of "+intLen+" characters.");
			}
	}

	function saveNoteVal()
	{
		if(document.all("chkNote").checked)
			document.all("chkNoteHidden").value = "on"
		else
			document.all("chkNoteHidden").value = ""
	}
	</script>
		
  <tr>
	<td colspan="8">
		<div style="overflow:hidden;height:4px; border-bottom-style:solid;border-bottom-color:gray;border-bottom-width:1px;">&nbsp;</div><!--hr-->
	</td>
  </tr>
  <tr style="padding-top:3px">
    <td align="right">
		<font color="red">Status:</font>
    </td>
    <td colspan="8"><table class="Label" cellpadding="0" cellspacing="0"><tr><td>
    <%
		varO = ""
		varP = ""
		varC = "" 
		varX = ""
		if not rsSQL.EOF then
			select case rsSQL.Fields("Status").Value
				case "o"
					varO = " selected"
				case "p"
					varP = " selected"
				case "c"
					varC = " selected"
				case "x"
					varX = " selected"
				case "r"
					varR = " selected"
				case "n"
					varN = " selected"
				case "w"
					varW = " selected"
					
			end select
		end if
		Response.Write "<select  class=""label"" id=""cboStatus"" name=""cboStatus"">" & vbcrlf
		Response.Write "<option" & varO & " value=""o"">Open</option>" & vbcrlf
		Response.Write "<option" & varP & " value=""p"">Pending</option>" & vbcrlf
		Response.Write "<option" & varC & " value=""c"">Closed</option>" & vbcrlf
		Response.Write "<option" & varR & " value=""r"">Reconfirmed</option>" & vbcrlf
		Response.Write "<option" & varX & " value=""x"">Canceled</option>" & vbcrlf
		Response.Write "<option" & varN & " value=""n"">Not Available</option>" & vbcrlf
		Response.Write "<option" & varW & " value=""w"">Wait List</option>" & vbcrlf
		Response.Write "</select>" & vbcrlf
		
		Response.Write "</td><td>"

		Response.Write "<font color=red>&nbsp;&nbsp;Status Notes:</font>&nbsp;"
		If booNewRec Then
			Response.Write "<input type=text disabled id=txtNotes name=txtNotes class=LongTxt>" & vbcrlf
		Else
			Response.Write "<input type=text disabled id=txtNotes name=txtNotes class=LongTxt value=""" & trim(rsSQL("CloseNote").Value) & """>" & vbcrlf
		End If
		Response.Write "</td><td align=right>&nbsp;&nbsp;Applied To:</td><td style=padding-left:2px>"
			Response.Write "<table class=Label bgcolor=lightyellow border=0 style=width:192px;height:18px;border-width:2px;border-style:inset;border-color:white cellpadding=0 cellspacing=1><tr><td nowrap><div id=divSummary style=""padding-left:3px;overflow:hidden;width:170px;height:14px""></div></td>"
			if booDeptSelect then
				 Response.Write "<td onmousedown=buttonAction() id=tdddoButton align=center style=""width:30px;border-style:outset;border-width:1px"" bgcolor=menu><img id=imgddo src=images/arrowdown.gif></td>"
			end if
			Response.Write "</tr></table>"
			Response.Write "<script src=DropDownCheckBox.asp language=javascript></script>"
			Response.Write "<script language=javascript>" & vbcrlf
			Response.Write "var ddo;" & vbcrlf

			if fuddid = 0 and booNewRec then
				dim rsfuddid
				set rsfuddid = server.CreateObject("adodb.recordset")
				rsfuddid.Open "select DepartmentID from tblDepartment where DepartmentName = 'Concierge'",cnSQL
				'fuddid = "1" '"d.DepartmentID"
				fuddid = rsfuddid.Fields("DepartmentID").Value
				rsfuddid.Close
				set rsfuddid = nothing
			end if

			'fuddid = "10" '"d.DepartmentID"
			if booNewRec then
				'if copyTask = "True" then
					'strSQL = "select d.*, case when " & fuddid & " = d.DepartmentID then 1 else 0 end as Checked from tblDepartment d left join tlnkAppointmentDepartment ad on d.DepartmentID = ad.DepartmentID where ud.UserID = " & fuid & " and ud.CompanyID = " & cid
				'else
					if su = 1 then
						strSQL = "select d.*, case when " & fuddid & " = d.DepartmentID then 1 else 0 end as Checked from tblDepartment d join tlnkCompanyDepartment cd on d.DepartmentID = cd.DepartmentID where cd.CompanyID = " & cid
					else
						strSQL = "select d.*, case when " & fuddid & " = d.DepartmentID then 1 else 0 end as Checked from tlnkUserDepartment ud join tblDepartment d on ud.DepartmentID = d.DepartmentID where ud.UserID = " & fuid & " and ud.CompanyID = " & cid
					end if
				'end if
			else
				if su = 1 then
					strSQL = "select d.*, case when d.DepartmentID in (select DepartmentID from tlnkAppointmentDepartment where AppointmentID = " & strAppointmentID & ") then 1 else 0 end as Checked from tlnkCompanyDepartment cd join tblDepartment d on cd.DepartmentID = d.DepartmentID where cd.CompanyID = " & cid & " order by d.DepartmentName"
				else
					strSQL = "select d.*, case when d.DepartmentID in (select DepartmentID from tlnkAppointmentDepartment where AppointmentID = " & strAppointmentID & ") then 1 else 0 end as Checked from tlnkUserDepartment ud join tblDepartment d on ud.DepartmentID = d.DepartmentID where ud.UserID = " & fuid & " and ud.CompanyID = " & cid & " order by d.DepartmentName"
				end if
			end if
			Response.Write "ddo = new dropDownCheckBox(""cmbDepartments"",""" & strSQL & """,301,528,122,200);" & vbcrlf
			Response.Write "ddo.init()" & vbcrlf
			Response.Write "</script>"
			'Response.Write strSQL
			'Response.End
		Response.Write "</td></tr></table>"
		'end if%>
    </td>
  </tr>
  <tr>
	<td colspan="8">
		<div style="overflow:hidden;height:4px; border-bottom-style:solid;border-bottom-color:gray;border-bottom-width:1px;">&nbsp;</div><!--hr-->
	</td>
  </tr>
  <tr>
	<td colspan="8">
		<table cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td align="center">
					<table style="border-style:ridge;border-width:1px;" cellpadding="4" class="Label">
						<tr>
							<td>
								<input id="cmdGuestTaskLetter" name="cmdGuestTaskLetter" style="FONT-SIZE: xx-small; HEIGHT: 22px; WIDTH: 71px" type="button" value="Guest Letter" title="View/Edit Guest Letter">
								<% If booNewRec Then %>
									<input style="display:none" disabled type="checkbox" name="chkLetterPrinted" id="chkLetterPrinted"><!-- &nbsp;Printed -->
								<% Else %>
									<input style="display:none" disabled type="checkbox" name="chkLetterPrinted" id="chkLetterPrinted" <%if rsSQL("LetterPrinted") = true then Response.Write "CHECKED" end if%>><!-- &nbsp;Printed -->
								<% End If%>
							</td>
						</tr>
					</table>
				</td>
				<td align="center">
				
				<table border="0" cellpadding="4" class="Label">
				<tr><td>
				<table style="border-style:ridge;border-width:1px;" cellpadding="4" class="Label">
				<tr><td>
					<select class="Label" size="1" name="cboLetterhead" style="color: darkblue">
						<option value="Yes">Letterhead</option>
						<option selected value="No">Plain Paper</option>
				</select>
					<input type="button" value="View Vendor" style="color: darkblue; FONT-SIZE: xx-small; HEIGHT: 22px; WIDTH: 75px" onclick="SubmitTab(6)" id="printloc" disabled name="printloc" title="View Vendor Report">
					<input type="button" value="Print" style="color: darkblue; FONT-SIZE: xx-small; HEIGHT: 22px; WIDTH: 48px" onclick="SubmitTab(4)" id="printtask" name="printtask" title="Print Task Report">&nbsp;
				</td></tr>
				</table>
				</td><td>
					<input id="txtReminder" name="txtReminder" type="hidden">
					<input id="FromReminder" name="FromReminder" type="hidden" value="<%=Request.Querystring("FromReminder")%>">
					<input id="cmdReminder" name="cmdReminder" style="color: darkblue; FONT-SIZE: xx-small; HEIGHT: 22px; WIDTH: 80px" type="button" value="Add Reminder" onclick="Reminder()" title="Add, Edit, or Remove a Reminder">
					<input id="cmdDelete" name="cmdDelete" style="color:darkred; FONT-SIZE: xx-small; HEIGHT: 22px; WIDTH: 56px" type="button" value="Delete" title="Delete this Task">
					<input id="cmdSaveAndCopy" name="cmdSaveAndCopy" style="FONT-SIZE: xx-small; HEIGHT: 22px; WIDTH: 56px" type="button" value="Copy" title="Copy this Task" onclick="SubmitTab(8)">
					<input id="cmdSaveAndPrint" name="cmdSaveAndPrint" style="color:darkGreen; FONT-SIZE: xx-small; HEIGHT: 22px; WIDTH: 71px" type="button" value="Save &amp; Print" onclick="SubmitTab(7)" title="Save, Print, &amp; Close">
					<input id="cmdSaveAndClose" name="cmdSaveAndClose" style="background-color:lightgreen;FONT-SIZE: xx-small; HEIGHT: 22px; WIDTH: 56px" type="button" value="Save" title="Save &amp; Close" onclick="SubmitTab(9)">
					<input id="cmdCancel" style="FONT-SIZE: xx-small; HEIGHT: 22px; WIDTH: 56px" type="button" value="Cancel" name="cmdCancel" title="Cancel, Lose Changes">
				</td></tr></table>	
				</td>
			<tr>
		</table>
	</td>
  </tr>
</table>
</td>
</tr>
</table>
<script Language="vbScript">
	sub cmdDelete_onclick
		dim intRetVal, booTaskSearch ', booIDE
		booTaskSearch = "False"
		'booIDE = "False"
		
		<% If (CLng("0" & strCreateUserID) = CLng("0" & remote.session("FloatingUser_UserID"))) or (fua) or (remote.session("SuperUser")=1)  Then %>
			intRetVal = MsgBox("Are you sure you want to delete this appointment?", vbYesNo,"Delete Appointment")
		<% Else %>
			intRetVal = vbNo
			
			Alert ("You do not have the rights to delete this task." & vbCrLF & "       Please see the Administrator")
		<% End If %>
		 
	    If intRetVal  = vbYes Then
			if instr(1,strOpener,"TaskSearch.asp") > 0 then
				booTaskSearch = "True"
			end if
			'if instr(1,strOpener,"ItineraryDetailEdit.asp") > 0 then
			'	booIDE = "True"
			'end if
			<%if Request.QueryString("CalledFromIDE") = "True" then
				response.write "booIDE = ""True""" & vbcrlf
			else
				response.write "booIDE = ""False""" & vbcrlf
			end if%>

			call document.frames("ifLogin").location.replace  ("DeleteApptConfirm.asp?ID=<%=aid%>&CalledFromTS=" & booTaskSearch & "&CalledFromIDE=" & booIDE)
		else
			call vendorLookup(false)
	    End If
	end sub
</script>
<!-- RECURRENCE SECTION -->
<%

intRecID = Request.QueryString ("RecID")
If  intRecID <> "" Then
Dim rsRec
Set rsRec = Server.createObject("ADODB.Recordset")

rsRec.open "Select * from tblRecurrence where RecID=" & intRecID, cnSQL


intRecEdit = Request.QueryString("RecEdit")

if not rsRec.EOF Then
	intOccurences = rsRec.Fields("Occurences").Value
	intRecEndDate = rsRec.Fields("Enddate").Value  
	intRecType = rsRec.Fields("type").Value 
	intFrequency = rsRec.Fields("frequency").Value 
	intRecDetail = rsRec.Fields("RecDetail").Value
	intRecShowOnce = rsRec.Fields("RecShowOnce").Value  
End If

rsRec.close
Set rsRec = Nothing

	'Response.Write "RECURRENCE"
End IF

%>

<input type="Hidden" name="RecUpdate" id="RecUpdate">
<input type="Hidden" name="RecID" id="RecID" value="<%=intRecID%>">
<input type="Hidden" name="RecEdit" id="RecEdit" value="<%=intRecEdit%>">
<input type="Hidden" name="RecOccurences" id="RecOccurences" value="<%=intOccurences%>">
<input type="Hidden" name="RecEndDate" id="RecEndDate" value="<%=intRecEndDate%>">
<input type="Hidden" name="RecType" id="RecType" value="<%=intRecType%>">
<input type="Hidden" name="RecFrequency" id="RecFrequency" value="<%=intFrequency%>">
<input type="Hidden" name="RecDetail" id="RecDetail" value="<%=intRecDetail%>">
<input type="Hidden" name="RecShowOnce" id="RecShowOnce" value="<%=intRecShowOnce%>">

<!-- END OF RECURRENCE SECTION -->
<% 	
'Response.Write "<script>alert('" & (len(intOTLOGID & "")) & "')</script>"
if (intOTLogID & "")  = "" Then intOTLogID = "0" 
if (intSSLogID & "")  = "" Then intOTLogID = "0" 


%>

<input type="Hidden" name="OTLogID" id="OTLogID" value="<%=intOTLogID%>">
<input type="Hidden" name="SSLogID" id="SSLogID" value="<%=intSSLogID%>">

</form>
</body>
</html>

<%
Function AMPM( hour, addMinute )
	dim strTime, intHour, strHour
	strHour = left(hour,2)
	intHour = cint(strHour)
	if intHour > 11 then
		if intHour = 12 then
			strTime = "12:" & addMinute & " PM"
		else
			strTime = right("0" & cstr(intHour-12),2) & ":" & addMinute & " PM"
		end if
	else
		if intHour <> 0 then
			strTime = right("0" & cstr(intHour),2) & ":" & addMinute & " AM"
		else
			strTime = "12:" & addMinute & " AM"
		end if
	end if
	AMPM = strTime
End Function

function CustomTime( strTime )
	CustomTime = right("0" & left(strTime,instrrev(strTime,":")-1),5) & " " & right(strTime,2)
end function

rsSQL.Close
set rsSQL = nothing
cnSQL.Close
set cnSQL = nothing

function ccn( s )
	dim a, str, cnt
	str = s
	p = instr(1,str,"-")
	if p > 0 then
		cnt = 1
		a = Array()
		pos1 = p
		do until p < 1
			redim preserve a(cnt)
			a(cnt-1) = p
			cnt = cnt + 1
			p = instr(p+1,str,"-")
		loop
		prefix = left(s,pos1-1)
		l = len(replace(mid(s,pos1+1),"-",""))
		for i = 1 to l
			therest = therest & "x"
		next
		str = prefix & therest
		for i = 0 to ubound(a)-1
			str = left(str,a(i)-1) & "-" & mid(str,a(i))
		next
	end if
	ccn = str
end function

%>
