<%@ Language=VBScript %>
<%
Response.CacheControl = "No-Cache"
Response.AddHeader "Pragma", "No-cache"
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")

Dim uKey 

function initialize()
	uKey = Request.Cookies("UserKey")
	If uKey = "" Then
		uKey = Trim(Request.QueryString ("ukey"))
	End If
	Response.Cookies("UserKey") = uKey
	remote.Init (uKey)
end function

initialize()

'if Request.QueryString("CalledFrom") = "Admin" then
	if remote.Session("_Login") <> "" then
		remote.Session("FloatingUser_Login")     = remote.Session("_Login")
		remote.Session("FloatingUser_Password")  = remote.Session("_Password")
		remote.Session("FloatingUser_UserID")    = remote.Session("_UserID")
		remote.Session("FloatingUser_UserName")  = remote.Session("_UserName")
		remote.Session("FloatingUser_UserLName") = remote.Session("_UserLName")
		remote.Session("FloatingUser_Admin")     = remote.Session("_Admin")
		remote.Session("FloatingUser_SuperUser") = remote.Session("_SuperUser")
		remote.Session("FloatingUser_EMail")     = remote.Session("_EMail")
		remote.Session("FloatingUser_CCPrivate") = remote.Session("_CCPrivate")
		remote.Session("FloatingUser_CCPublic")  = remote.Session("_CCPublic")
		remote.Session("FloatingUser_Title")     = remote.Session("_Title")
	end if
'end if

remote.Session("_Login") = ""
rsa = remote.session("AvailHeight")
rsw = remote.session("AvailWidth")

cid = remote.Session("CompanyID")

strSU = remote.Session("FloatingUser_SuperUser")
booSU = false
if strSU <> "" then
	booSU = cbool(strSU)
end if
strAdmin = remote.Session("FloatingUser_Admin")
booAdmin = false
if strAdmin <> "" then
	booAdmin = cbool(strAdmin)
end if

Set cnSQL = Server.CreateObject("ADODB.Connection")
set rsCalView = Server.CreateObject("ADODB.Recordset")
set rsDepartments = Server.CreateObject("ADODB.Recordset")
cnSQL.Open Application("sqlInnSight_ConnectionString")
set rsCalView = cnsql.Execute("select * from vwCalView where CompanyID = " & cid & " or CompanyID = 999999 order by Name")

if remote.session("SuperUser") = 1 then
	set rsDepartments = cnsql.Execute("select d.* from tblDepartment d join tlnkCompanyDepartment cd on d.DepartmentID = cd.DepartmentID where cd.CompanyID = " & cid)
else
	set rsDepartments = cnsql.Execute("select d.* from tblDepartment d join tlnkUserDepartment ud on d.DepartmentID = ud.DepartmentID where ud.UserID = " & remote.session("UserID") & " and ud.CompanyID = " & cid)
end if
strCalView = ""
strDepartments = ""
booViewsExist = false
booDepartmentsExist = false
do until rsCalView.EOF
	booViewsExist = true
	if rsCalView.Fields("CalViewID").Value = cint(remote.Session("DefaultCalView")) then
		selected = "document.all.cmbCalView.value = " & rsCalView.Fields("CalViewID").Value & ";"
	else
		selected = "document.all.cmbCalView.visible = false;"
	end if
	strCalView = strCalView & "document.all.cmbCalView.length++;document.all.cmbCalView(document.all.cmbCalView.length-1).value = " & rsCalView.Fields("CalViewID").Value & ";document.all.cmbCalView(document.all.cmbCalView.length-1).text = '" & trim(replace(rsCalView.Fields("Name").Value,"'","\'")) & "';" & selected
	rsCalView.MoveNext
loop
fuddid = remote.Session("FloatingUser_DDID")
DepartmentCount = 0
do until rsDepartments.EOF
	booDepartmentsExist = true
	DepartmentCount = DepartmentCount + 1
	if fuddid <> "" then
		if rsDepartments.Fields("DepartmentID").Value = cint(fuddid) then
			selected = "document.all.cmbDepartments.value = " & rsDepartments.Fields("DepartmentID").Value & ";"
		else
			selected = "document.all.cmbDepartments.visible = false;"
		end if
	else
		selected = "document.all.cmbDepartments.visible = false;"
	end if
	strDepartments = strDepartments & "document.all.cmbDepartments.length++;document.all.cmbDepartments(document.all.cmbDepartments.length-1).value = " & rsDepartments.Fields("DepartmentID").Value & ";document.all.cmbDepartments(document.all.cmbDepartments.length-1).text = '" & trim(replace(rsDepartments.Fields("DepartmentName").Value,"'","\'")) & "';" & selected
	rsDepartments.MoveNext
loop
'Response.Write strCalView
rsCalView.Close
set rsCalView = nothing
rsDepartments.Close
set rsDepartments = nothing
cnsql.Close
set cnsql = nothing


if isNull(rsa) or rsa = "" then
	Response.Write "<script language=vbscript>" & vbclf
	Response.Write "msgbox ""Your session has timed out.  You need to re-login.  Click OK to close this window."",vbCritical,""Time Out""" & vbcrlf
	Response.Write "window.close()" & vbcrlf
	Response.Write "</script>"
	Response.End
end if

dim strSource

If Len(Request.QueryString("TargetDate")) > 0 Then
	strSource = "TaskPad.asp?TargetDate=" & Request.QueryString("TargetDate") & "&"
else
	strSource = "TaskPad.asp?"
end if

ah = cint(rsa)
aw = cint(rsw)
'buttonHeight = ah * .048
buttonFontSize = Round(ah/52,0)
buttonFontSize = 12
%>

<!--#INCLUDE file="include/vbfunc.asp"-->
<!--#include file=Global.asp -->

<script Language="JavaScript1.2">
<!--#INCLUDE file="ddCalendar.asp"-->
</script>

<script language="javascript">
var idarr = new Array();

	var TimerID;
	self.moveTo(0,0);
	self.resizeTo(screen.availWidth,screen.availHeight);
</script>


<%
dim intHeightIncrement, intWidthIncrement, searchWinW, searchWinH

If remote.Session("ScreenHeight") < 750 Then
	intHeightIncrement = 230 '170
	intWidthIncrement = 730 '700 '771
	GridHeight = 270 '250
	'AppFrameTop = (ah-540)/2
	'SearchFrameTop = AppFrameTop
	searchWinH = 518
	searchWinW = 756
	LeftColWidth = 129
	LogoIMGHeight = 82
	logoDivWidth = 126
	if remote.Session("UseGuestProfile") = "True" then
		buttonHeight = ah * .041 '.041
	else
		buttonHeight = ah * .046 '.041
	end if
	buttonWidth = 119
	logoDivHeight = 74
	HotelLogoWidth = 127
Else
	intHeightIncrement = 120
	intWidthIncrement = 880
	GridHeight = 540 '375
	'AppFrameTop = (ah-(GridHeight+intHeightIncrement))/2
	'SearchFrameTop = 20
	searchWinH = 640
	searchWinW = 990
	'LeftColWidth = aw * .16
	LeftColWidth = 158
	if remote.Session("UseGuestProfile") = "True" then
		buttonHeight = ah * .041
	else
		buttonHeight = ah * .044
	end if
	buttonWidth = 149
	LogoIMGHeight = 76 '150
	logoDivHeight = 122
	logoDivWidth = 158
	HotelLogoWidth = 158
End If

calloutleft = remote.Session("AvailWidth")-(buttonWidth*2)-60

remote.Session("LeftColWidth") = LeftColWidth
'TaskPadLeft = ((aw-589)/2)+67

GridPadLeft = (aw-intWidthIncrement)/2
GridPadTop = (ah-(GridHeight+intHeightIncrement))/2
If Len(Request.QueryString("TargetDate")) > 0 Then
  remote.Session("TargetDate") = Request.QueryString("TargetDate")
Else
  remote.Session("TargetDate") = Date()
End If

'LeftColWidth = aw * .16
intTaskWidth = rsw-LeftColWidth-200
%>
<html>
<head>
<style> 
	.but {border-width:thin;background-color:silver;border-style:outset;cursor:default;height:10}
	.Norm {cursor:default;background-color:#F9D568;}
	.TS {width:30;font-weight:400;background-color:#F9D568;color:Navy;font-family:Tahoma;font-size:11;height:16} 
	.HS {width:110;text-align:center;vertical-align:middle;font-weight:800;background-color:#F9D568;font-family:Tahoma;color:Navy;font-size:11;height:26}
	.ED {cursor:default;font-family;Tahoma;color:black;background-color:white;height:16}
	.CD {cursor:default;text-align:center;vertical-align:middle;font-family:Tahoma;color:white;background-color:Navy;font-size:11;height:16}
	.MD {cursor:default;text-align:center;vertical-align:middle;font-family:Tahoma;color:gray;background-color:#f0f0f0;font-size:11;height:16}
	.Today {cursor:default;text-align:center;vertical-align:middle;font-family:Tahoma;color:white;background-color:white;font-size:11;height:16;border-style:solid;border-color:magenta;border-size:1px}
	.ND {cursor:default;text-align:center;vertical-align:middle;font-family:Tahoma;color:black;background-color:white;font-size:11;height:16}
	.WE {cursor:default;text-align:center;vertical-align:middle;font-family:Tahoma;color:red;background-color:white;font-size:11;height:20}

	.GreenFont { background-color:menu; font-family: Veranda; FONT-SIZE: <%=buttonFontSize%>px; font-weight:normal; COLOR: green; HEIGHT: <%=buttonHeight%>px; WIDTH: <%=buttonWidth%>px }
	.BlueFont  { background-color:menu;font-family: Veranda; FONT-SIZE: <%=buttonFontSize%>px; font-weight:normal; COLOR: #000080; HEIGHT: <%=buttonHeight%>px; WIDTH: <%=buttonWidth%>px }
	.RedFont   { background-color:menu;font-family: Veranda; FONT-SIZE: <%=buttonFontSize%>px; font-weight:normal; COLOR: red; HEIGHT: <%=buttonHeight%>px; WIDTH: <%=buttonWidth%>px }
	.BlackFont { background-color:#D4D0C8;font-family: Veranda; FONT-SIZE: <%=buttonFontSize%>px; font-weight:normal; COLOR: black; HEIGHT: <%=buttonHeight%>px; WIDTH: <%=buttonWidth%>px }

	<!--#006699
	#ABBA75
	#309E47
	#709E87-->

	.TaskDetailLabel	{ font-family: Tahoma; font-size: 10; font-weight: bold; color: black }
	.TaskDetailField	{ font-family: Tahoma; font-size: 10; background-color: silver; color: blue; border-style: outset; padding-left: 3; padding-right: 3}
	.DateTimeFont		{ font-family: Tahoma; font-size: 11 }
	
</style>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Dim strEasterEgg

Sub ShowTask(param)
	
	If param = "" Then 	
		tmpStr = "Appointment.asp?TargetDate=" & window.calObj.getVal() & "&ID=0&Hour=" & FormatDateTime("12:00",4)
	Else
		tmpStr = param
	End If
	
		w = 736
		h = 494
		
	wtop = (screen.availHeight - h) / 2
	wleft = (screen.availWidth - w) / 2

	param = "Top=" & wtop & ", Left=" & wleft & " ,Height=" & h & ", Width=" & w
	
	If Instr(tmpStr,"RecID") > 1 Then
		xx = window.showModalDialog("RecurrenceDialog.asp",,"dialogheight: 120px; dialogwidth: 200px; status: no; center: yes; scroll: no")
		If xx > 0 Then
		
				' RecEdit Value 1 open this instance 2- open Series
				tmpStr = tmpStr & "&RecEdit=" & xx
				xx = window.open (tmpStr, "",param ,null)
		End If
	Else
		'if dialog <> "" then
		'	xx = window.showModalDialog(tmpStr, window,param ,null)
		'else
			xx = window.open (tmpStr, "",param ,null)
		'end if
	End If
End Sub


sub	cmdLocationRequest_onclick
	<%if remote.Session("BPW") then%>
		'x = window.showModalDialog("LocationRequestFormFrame.asp?v=0625021015",,"dialogheight: 494px; dialogwidth: 537px; status: no; center: yes; scroll: no")
		window.open "LocationRequestFormFrame.asp?v=0625021015","","height=494px,width=537px,status=no,top=" & (screen.availHeight-494)/2 & "px,left=" & (screen.availWidth-537)/2 & "px,scroll=no,toolbar=no,address=no"
	<%else%>
		dim strSQL
		strSQL = "ValidatePassword.asp?v=1&Caption=Administration" 
		x = showModalDialog(strSQL,"","center:yes;status:no;scrollbars:no;dialogHeight:116px;dialogWidth:298px;")
		if x <> "" then
			if x <> "Invalid Password" then
				'x = window.showModalDialog("LocationRequestFormFrame.asp?v=0625021015",,"dialogheight: 494px; dialogwidth: 537px; status: no; center: yes; scroll: no")
				window.open "LocationRequestFormFrame.asp?v=0625021015","","height=494px,width=537px,status=no,top=" & (screen.availHeight-494)/2 & "px,left=" & (screen.availWidth-537)/2 & "px,scroll=no,toolbar=no,address=no"
			else
				msgbox "Invalid password.",vbCritical,"Password"
			end if
		end if
	<%end if%>
	'document.all("frameLogo").focus()
end sub

sub	cmdHelp_onclick
	x = window.showModalDialog("Help.asp",window.document,"dialogheight: 330px; dialogwidth: 420px; status: no; center: yes; scroll: no")
	if x then
		cmdLocationRequest_onclick
	end if
	'document.all("frameLogo").focus()
end sub

sub	cmdFaxCover_onclick
	'window.open "GetFaxCover.asp","","height=650px,width=700px,status=no,top=" & (screen.availHeight-650)/2 & "px,left=" & (screen.availWidth-700)/2 & "px,toolbar=no,address=no"
	'x = window.showModalDialog("GetFaxCover.asp",window.document,"dialogheight: 330px; dialogwidth: 420px; status: no; center: yes; scroll: no")
	'if x then
	'	cmdLocationRequest_onclick
	'end if
	'document.all("frameLogo").focus()
	<%if remote.Session("BPW") then%>
		window.open "GetFaxCover.asp","","height=650px,width=700px,status=no,top=" & (screen.availHeight-650)/2 & "px,left=" & (screen.availWidth-700)/2 & "px,toolbar=no,address=no"
	<%else%>
		dim strSQL
		strSQL = "ValidatePassword.asp?v=1&Caption=Administration" 
		x = showModalDialog(strSQL,"","center:yes;status:no;scrollbars:no;dialogHeight:116px;dialogWidth:298px;")
		if x <> "" then
			if x <> "Invalid Password" then
				window.open "GetFaxCover.asp","","height=650px,width=700px,status=no,top=" & (screen.availHeight-650)/2 & "px,left=" & (screen.availWidth-700)/2 & "px,toolbar=no,address=no"
			else
				msgbox "Invalid password.",vbCritical,"Password"
			end if
		end if
	<%end if%>
end sub

sub	cmdWeather_onclick
	'window.open "http://www.goconcierge.net/NoSSL/Weather.asp?zip=<%=remote.Session("CompanyZip")%>&cid=<%=remote.Session("CompanyID")%>","winWeather","resizable=yes,menubar=yes,toolbar=yes,titlebar=yes,status=yes,location=yes,scrollbars=yes,top=0,left=0,height=" & cstr(screen.availHeight*.74) & "px,width=" & cstr(screen.availWidth*.95) & "px"
	window.open "http://<%=remote.session("WeatherURL")%>","Weather","resizable=yes,menubar=yes,toolbar=yes,titlebar=yes,status=yes,location=yes,scrollbars=yes,top=0,left=0,height=" & cstr(screen.availHeight*.74) & "px,width=" & cstr(screen.availWidth*.75) & "px"
end sub

sub cmdMovies_onclick
	window.open "http://<%=remote.session("MoviesURL")%>","Movies","resizable=yes,menubar=yes,toolbar=yes,titlebar=yes,status=yes,location=yes,scrollbars=yes,top=0,left=0,height=" & cstr(screen.availHeight*.74) & "px,width=" & cstr(screen.availWidth*.75) & "px"
end sub

sub cmdZagat_onclick
	window.open "http://<%=remote.session("ZagatURL")%>","Zagat","resizable=yes,menubar=yes,toolbar=yes,titlebar=yes,status=yes,location=yes,scrollbars=yes,top=0,left=0,height=" & cstr(screen.availHeight*.74) & "px,width=" & cstr(screen.availWidth*.75) & "px"
end sub

sub cmdFlights_onclick
	window.open "http://<%=remote.session("FlightsURL")%>","Flights","resizable=yes,menubar=yes,toolbar=yes,titlebar=yes,status=yes,location=yes,scrollbars=yes,top=0,left=0,height=" & cstr(screen.availHeight*.74) & "px,width=" & cstr(screen.availWidth*.75) & "px"
end sub

sub cmdTickets_onclick
	window.open "http://<%=remote.session("TicketsURL")%>"
end sub

sub cmdCalculator_onclick
	tp = (screen.availHeight-148)/2
	lft = (screen.availWidth-146)/2
	window.open "calculator.asp?v=0","","location=0,resizable=0,titlebar=0,top=" & tp & ",left=" & lft & ",width=146px,height=176px,status=0,menubar=0,scrollbars=0"
	'window.showModelessDialog "calculator.asp?v=0","","dialogTop:" & tp & ";dialogLeft:" & lft & ";dialogHeight=148px;dialogWidth:146px;status:no;center:yes;edge:raised;help:no;scroll:no;unadorned:yes"
end sub

Sub cmdCustomDirections_onclick
	tmpStr = "customdirections.asp?TargetDate=" & window.calObj.getVal()
	window.showModelessDialog tmpStr,"","scroll:no;center:yes;status:no;dialogHeight:520px;dialogWidth:660px"
End Sub

sub cmdSwitchCompany_onclick
  window.parent.location.href = "SelectCompany.asp"
end sub

<% IF remote.Session("VCT") Then %>

Sub cmdGuestEnvelope_onclick
	window.showModelessDialog "BlankEnvelope2.asp?TargetDate=" & window.calObj.getVal(),"","center:yes;status:no;scroll:no;dialogHeight:490px;dialogWidth:646px"
End Sub


Sub cmdPrintCalendar_onclick
	ClearTO()
	window.showModalDialog "PrintCalendarFrame.asp?v=2","","center:yes;dialogHeight:250px;dialogWidth:372px;scroll:no;status:no"
	StartTimer()
End Sub

Sub StartTimer
	nMinutes = 3
	TimerID = window.setInterval("TimerFunction(1)",60 * 1000 * nMinutes)
end sub
sub cmdTaskSearch_onclick
  window.parent.location.href = "TaskSearch.asp?CompanyID=-1"
End Sub

sub cmdGuestProfiles_onclick
	retval = window.showModalDialog("GuestProfileSetup.asp?mode=Switchboard&load=1&GPSearchID=<%=remote.Session("GPSearchID")%>",window,"dialogHeight:410px;dialogWidth:700px;center:yes;scroll:no;status:no")
	if retval <> "," and retval <> "" then
		a = split(retval,",")
		showtask "Appointment.asp?TargetDate=" & window.calObj.getVal() & "&ID=0&Hour=" & FormatDateTime("12:00",4) & "&dgid=" & a(0) & "&gid=" & a(1)
	end if
end sub

sub cmdItinerary_onclick
  ClearTO()
  dim x
  x = ""
  select case showModalDialog("ItineraryAddEditDialog.asp?v=2.4","","dialogHeight:100px;dialogWidth:320px;status:no;scroll:no;center:yes;resizable:no")
	case "Existing"
		x = window.showModalDialog("ItinerarySearch.asp",window,"status:no;DialogHeight:550px;DialogWidth:756px;scroll:no;center:yes")
	case "New"
		<%if remote.Session("BPW") then%>
			call mynavigate(0)
		<%else%>
			strSQL = "ValidatePassword.asp?v=1&Caption=New Itinerary"
			x = showModalDialog(strSQL,"","center:yes;status:no;scroll:no;dialogHeight:116px;dialogWidth:298px;")
			if x <> "" then
				if x <> "Invalid Password" then
					call mynavigate(0)
				else
					alert "Invalid password."
				end if
			end if
		<%end if%>
  end select
  StartTimer
  TimerFunction 1
End Sub

sub mynavigate(iid)
	xx = window.showModalDialog("ItineraryTaskFrame.asp?iid=" & iid & "&CompanyID=<%=cid%>&FloatingUser_ID=<%=remote.Session("FloatingUser_ID")%>&AvailHeight=<%=ah%>&sd=" & formatdatetime(date(),vbShortDate) & "&ed=" & formatdatetime(date()+7,vbShortDate), window,"dialogHeight:522px; dialogWidth:750px;status:no;scroll:no;center:yes")
end sub

sub cmdEventSearch_onclick
  window.parent.location.href = "EventsMain.asp"
End Sub

Sub cmdTaskReport_onclick
	document.parentWindow.location.href ="TaskReport.asp?TargetDate=" & window.calObj.getVal() & "&ID=0"
End Sub

<% End If %>

Sub cmdSearchByCategory_onclick
  x = window.showModalDialog("BrowseLocationsFrame.asp?Mode=Search",window,"status:no;DialogHeight:<%=searchWinH%>px;DialogWidth:<%=searchWinW%>px;scroll:no;center:yes")
  'x = window.open("BrowseLocationsFrame.asp?Mode=Search")
End Sub

Sub cmdAddEditLocation_onclick
	document.parentWindow.location.href = "LocationSetup3.asp"
End Sub

Sub cmdAdmin_onclick
	<%
	if not booSU then%>
		dim strSQL
		strSQL = "ValidatePassword.asp?v=1&Caption=Administration" 
		x = showModalDialog(strSQL,"","center:yes;status:no;scrollbars:no;dialogHeight:116px;dialogWidth:298px;")
		if x <> "" then
			if x <> "Invalid Password" then
				if left(mid(x,instr(1,x,"Admin=")+6,4),4) = "True" or left(mid(x,instr(1,x,"SuperUser=")+10,4),4) = "True" then
			   		document.parentWindow.location.href = "Administration.asp"
			   	else
					msgbox "Sorry, you do not have rights to enter this area.  See the administrator.",48,"Administration"
			   	end if
			else
				msgbox "Invalid password.",vbCritical,"Password"
			end if
		end if
	<%else%>
   		document.parentWindow.location.href = "Administration.asp"
	<%end if%>
End Sub

Sub cmdExit_onclick
	window.close
End Sub

Function document_oncontextmenu
	document_oncontextmenu = false 'event.ctrlKey (for debug mode)
End Function

function checkEasterEgg()
	strEasterEgg = strEasterEgg & chr(window.event.keyCode)
	if instr(1,ucase(strEasterEgg),"SERVER NAME PLEASE") > 0 then
		strEasterEgg = ""
		set xmlHttp = CreateObject("Microsoft.XMLHTTP")
		xmlHttp.open "get", "GetServerName.asp", false
		xmlHttp.send ""
		str = xmlHttp.responseText
		msgbox str,vbOKOnly,"Server Name is..."
	end if
	if len(strEasterEgg) > 200 then
		strEasterEgg = ""
	end if
end function
</script>

<title>GoConcierge.net - <%=remote.Session("CompanyName") & " - User: " & remote.Session("FloatingUser_UserName") & " " & remote.Session("FloatingUser_UserLName")%></title>


<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
function window_onload() {
	//TimerID = window.setInterval("TimerFunction(1)",60 * 60)
	strEasterEgg = "";
	<%=strCalView%>
	<%=strDepartments%>

	var d = new Date('<%=now()%>');
	var tmpTime = (d.getMonth()+1).toString()+(d.getDate()).toString()+(d.getYear().toString())+(d.getHours()).toString()+(d.getMinutes()).toString()+(d.getSeconds()).toString()+(Math.random()*10000).toString()
	Calobj.setDate(d);
	c.value = formatDate(d);
	d = null;
	window.document.all("frameTaskPad1").src = "<%=strSource%>v=" + tmpTime
	setToday(true);

	StartTimer();
	
	//alert('<%=Request.Cookies ("GCNBackup")%>')
	
	<% 
		If Cint("0" & Request.Cookies ("GCNBackup")) = 0 or Cint("0" & Request.Cookies ("GCNBackup")) = 1 Then %>
		// var xx = window.showModalDialog("BackupPrompt.asp" ,null,"dialogheight: 220px; dialogwidth: 300px; status: no; center: yes; scroll: no")	
		// document.all("printbackup").src  = "BackupDownload.asp?action=" + xx
	<% 
		End If %>
}
//-->
</script>
</head>
<div id="printbackupdiv" style="z-index:10;">
			<object id="reportViewer" width="1" height="1" classid="clsid:8569D715-FF88-44BA-8D1D-AD3E59543DDE" VIEWASTEXT codebase="arview2.cab#version=2,0,0,1214">
			</object>

<script language="VBScript">
Dim booPrint

Sub reportViewer_LoadCompleted
	If booPrint=1 Then reportViewer.PrintReport false
	booPrint=0
End Sub

Sub PrintReport(i)
	booPrint=1	
End Sub
</script>

<iframe id="PrintBackup" src="PrintBackup.asp" Height="1" Width="1"></iframe>
</div>
<body onkeyup="vbscript:checkEasterEgg()" scroll="no" bgcolor="silver" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0" link="black" vlink="black" alink="black" LANGUAGE="javascript" onload="return window_onload()">
<!--#include file = "Header.inc"--> 


<input type="hidden" value="0" id="varAction" name="varAction">
<input type="hidden" value="0" id="varActionType" name="varActionType">
<input type="hidden" value id="txtSalutation" name="txtSalutation">
<input type="hidden" id="txtRoom" name="txtRoom">
<input type="hidden" id="txtLocation" name="txtLocation">
<input type="hidden" id="txtTimeEnd" name="txtTimeEnd">
<input type="hidden" id="txtTimeStart" name="txtTimeStart">
<input type="hidden" id="txtDateEnd" name="txtDateEnd">
<input type="hidden" value="0" id="booMustRestore" name="booMustRestore">

<!--div id="divGuestEnvelope" name="divGuestEnvelope" style="position: absolute; top: 12; left: <%=GridPadLeft%>; visibility: hidden;">
	<iframe scrolling="no" AllowTransperency="true" frameborder="0" framespacing="0" height="<%=GridHeight+intHeightIncrement%>" width="<%=intWidthIncrement%>" id="frameGuestEnvelope" name="frameGuestEnvelope" src="LoadingAppointment.asp"></iframe>
</div-->

<!--div id="divSwitchboard" name="divSwitchboard" style="z-index: 100; position: absolute; top: 4; left: 0; visibility: visible"><!-- filter:progid:DXImageTransform.Microsoft.RandomDissolve()"-->

<br>
<table cellpadding="0" cellspacing="0" align="center" width="100%">
<tr>
<td align="center">
<table align="center" style="BORDER-RIGHT: medium none; BORDER-TOP: medium none; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none; BORDER-COLLAPSE: collapse; mso-border-alt: solid windowtext .5pt" cellSpacing="0" cellPadding="0" bgColor="silver" border="1">
<tr>
<td valign="top" align="center" style="border-right-style:solid; WIDTH: <%=LeftColWidth%>px; PADDING-TOP: 0in; BORDER-TOP: windowtext 0.5pt solid; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 0.5pt solid;BORDER-BOTTOM: windowtext 0.5pt solid; padding-top: 2px;">
<table width=100% border=0 cellspacing=0 cellpadding=0>
	<tr>
		<td align=center>
			<div onclick=javascript:TimerFunction(1); style="padding-left:1px;overflow:hidden;height:<%=logoDivHeight%>px;width:<%=logoDivWidth%>px">
			<table bgcolor="<%=remote.Session("LogoBGColor")%>" width="98%" height="100%" valign="middle" align="center" cellpadding="0" cellspacing="0">
				<tr>
				<td align="center" valign="middle">
					<img align="center" SRC="ClientUploads/<%=remote.Session("ScreenLogoLocation")%>" alt="<%=remote.Session("CompanyName")%>">
				</td></tr>
			</table>
			</div>
		</td>
	</tr>
	<tr>
		<td style="padding-left:1px;text-align:center;vertical-align:middle;height: 26px">
			<div style="padding-top:3px;vertical-align:middle;height:100%;width:98%;background-color:#F9D568;"><input language="javascript" onselectstart="{window.event.returnValue=false;}" type="text" class="DateTimeFont" id="t" style="vertical-align:middle;text-align: center;border-style: none;background-color: transparent; width:98%"></div>
		</td>
	</tr>
</table>
<input value type="hidden" id="tcal">

	<table align=center style="border-top-style:solid; border-top-color:black;border-top-width:1px;" width="100%" cellpadding="0" cellspacing="0">
	<tr><td>
	<table border=0 style="padding-left:3px" cellpadding="0" cellspacing="2" width="100%">
	<tr><td style="padding-top:4px" valign="top" height="<%=buttonHeight+10%>">
	<input id="cmdSearchByCategory" type="button" value="Search Locations" name="cmdSearchByCategory" class="GreenFont" onmouseover="cmdSearchByCategory.style.borderWidth=4" onmouseout="cmdSearchByCategory.style.borderWidth=2">
	</td></tr>
	<tr><td>
	<input id="cmdCustomDirections" name="cmdCustomDirections" type="button" value="Custom Directions" class="BlueFont" onmouseover="cmdCustomDirections.style.borderWidth=4" onmouseout="cmdCustomDirections.style.borderWidth=2">
	</td></tr>
	<%If remote.Session("VCT") Then %>
	<tr><td>
	<input id="cmdPrintCalendar" name="cmdPrintCalendar" type="button" value="Print Calendar" class="BlueFont" onmouseover="cmdPrintCalendar.style.borderWidth=4" onmouseout="cmdPrintCalendar.style.borderWidth=2">
	</td></tr>
	<% End IF %>
	<tr><td>
	<input id="cmdLocationRequest" name="cmdLocationRequest" type="button" value="Request Change" class="BlueFont" onmouseover="cmdLocationRequest.style.borderWidth=4" onmouseout="cmdLocationRequest.style.borderWidth=2">
	</td></tr>
	<tr><td>
	<input type="button" value="Task Reports" id="cmdTaskSearch" name="cmdTaskSearch" class="BlueFont" onmouseover="cmdTaskSearch.style.borderWidth=4" onmouseout="cmdTaskSearch.style.borderWidth=2">
	</td></tr>
	<tr><td>
	<!--input disabled type="button" value="Event Search" id="cmdEventSearch" name="cmdEventSearch" class="BlueFont" size="22" onmouseover="cmdEventSearch.style.borderWidth=4" onmouseout="cmdEventSearch.style.borderWidth=2">	<br-->
	<!--input id="cmdWeather" name="cmdWeather" type="button" value="Weather Report" class="BlueFont" size="22" onmouseover="this.style.borderWidth=4" onmouseout="this.style.borderWidth=2">	<br-->
	<input id="cmdCalculator" name="cmdCalculator" type="button" value="Calculator" class="BlueFont" onmouseover="this.style.borderWidth=4" onmouseout="this.style.borderWidth=2">
	</td></tr>

	<tr><td>
	<input type="button" value="Guest Itinerary" id="cmdItinerary" name="cmdItinerary" class="BlueFont" onmouseover="cmdItinerary.style.borderWidth=4" onmouseout="cmdItinerary.style.borderWidth=2">
	</td></tr>
	
	<%if remote.Session("UseGuestProfile") = "True" then%>
	<tr><td>
	<input type="button" value="Guest Profiles" id="cmdGuestProfiles" name="cmdGuestProfiles" class="BlueFont" onmouseover="this.style.borderWidth=4" onmouseout="this.style.borderWidth=2">
	</td></tr>
	<%end if%>

	<tr><td valign="top">
	<input id="cmdHelp" name="cmdHelp" type="button" value="Support" class="BlueFont" onmouseover="cmdHelp.style.borderWidth=4" onmouseout="cmdHelp.style.borderWidth=2">
	</td></tr>
	
	<tr><td valign="top" height="<%=buttonHeight+12%>">
	<input type="button" value="Fax Cover" id="cmdFaxCover" name="cmdFaxCover" class="BlueFont" onmouseover="cmdFaxCover.style.borderWidth=4" onmouseout="cmdFaxCover.style.borderWidth=2">
	</td></tr>
	
	<%if booSU then%>
		<tr><td>
		<input type="button" value="Add/Edit Location" id="cmdAddEditLocation" name="cmdAddEditLocation" class="BlackFont" onmouseover="cmdAddEditLocation.style.borderWidth=4" onmouseout="cmdAddEditLocation.style.borderWidth=2">
		</td></tr>
	<%end if	
	if remote.Session("MultiCo") = "True" then%>
		<tr><td>
		<input type="button" value="Switch Company" id="cmdSwitchCompany" name="cmdSwitchCompany" class="BlackFont" onmouseover="this.style.borderWidth=4" onmouseout="this.style.borderWidth=2">
		</td></tr>
	<%end if
	
	If booAdmin or booSU Then%>
	<tr><td>
		<input type="hidden" value="Login New User" id="cmdLogin" name="cmdLogin" class="BlackFont" onmouseover="cmdLogin.style.borderWidth=4" onmouseout="cmdLogin.style.borderWidth=2">
		<input type="button" value="Administration" id="cmdAdmin" name="cmdAdmin" class="BlackFont" onmouseover="cmdAdmin.style.borderWidth=4" onmouseout="cmdAdmin.style.borderWidth=2">
	</td></tr>
	<%End If
	if remote.session("SuperUser") = 1 then
		incrH = 0
	else
		incrH = 12
	end if
	%>
	</center>

	<tr><td valign="bottom" height="<%=buttonHeight+incrH%>">
	<input type="button" value="Exit" id="cmdExit" name="cmdExit" class="RedFont" onmouseover="cmdExit.style.borderWidth=4" onmouseout="cmdExit.style.borderWidth=2">
	</td></tr>
	<!--tr><td><table cellpadding=0 cellspacing=0-->
	<%
	if booDepartmentsExist or booViewExist then
		Response.Write "<tr><td style=height:" & buttonHeight-12 & "px;vertical-align:bottom></td></tr>"
	end if
	'if booDepartmentsExist = true then
	if DepartmentCount > 1 then
		strDisplay = "inline"
	else
		strDisplay = "none"
	end if%>
	<tr style="display:<%=strDisplay%>"><td class=DateTimeFont valign="bottom">
	<table style="border-style:outset;border-width:1px:border-color:black;" cellpadding=0 cellspacing=0 class=DateTimeFont>
		<tr><td><select onchange=departmentChange(this) class=DateTimeFont style="width:<%=buttonWidth-2%>px" id=cmbDepartments name=cmbDepartments><%if remote.session("Admin") = "True" OR DepartmentCount > 1 then Response.Write "<option value=0>(All Departments)</option>" end if%></select></td></tr>
	</table>
	</td></tr>
	<%if cid = 238 then%>
	<tr><td class=DateTimeFont valign="bottom">
	<br>
	<table style="border-style:outset;border-width:1px:border-color:black;" cellpadding=0 cellspacing=0 class=DateTimeFont>
		<tr><td>
		<select class=DateTimeFont style="width:<%=buttonWidth-2%>px" id=cmbHotels name=cmbHotels>
			<option>Starwood Concierge</option>
			<option>Sheraton San Diego</option>
			<option>W Atlanta</option>
			<option>W Chicago City Center</option>
			<option>W Chicago Lakeshore</option>
			<option>W Los Angeles Westwood</option>
			<option>W Mexico City</option>
			<option>W New Orleans</option>
			<option>W New Orleans French Q</option>
			<option>W New York</option>
			<option>W New York The Court</option>
			<option>W New York Times Square</option>
			<option>W New York Union Square</option>
			<option>W San Diego</option>
			<option>W San Fransisco</option>
			<option>W Seattle</option>
			<option>W Silicon Valley</option>
			<option>Westin St. Francis</option>
		</select></td></tr>
	</table>
	</td></tr>
	<%end if

	if booViewsExist = true then%>
	<tr><td class=DateTimeFont valign="bottom">
	<table style="border-style:outset;border-width:1px:border-color:black;" cellpadding=0 cellspacing=0 class=DateTimeFont>
		<tr><td><select onchange=calViewChange(this) class=DateTimeFont style="width:<%=buttonWidth-2%>px" id=cmbCalView name=cmbCalView><option value=0>(Standard View)</option></select></td></tr>
	</table>
	</td></tr>
	<%end if%>
	<!--/table></td></tr-->
	</table>
	</td>
	<!--tr>
	<td align="center" style="padding-left:1px;vertical-align:bottom;height:<%=LogoIMGHeight%>px" align="center"><img src="images/gcn_bottom_logo.jpg"></td>
	</tr-->
	</table>

</td>

<td valign="top" align="center" style="BORDER-RIGHT: windowtext .5pt solid; PADDING-RIGHT: 0pt; BORDER-TOP: windowtext 0.5pt solid; PADDING-LEFT: 0; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 0.0pt solid; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 0.5pt solid">

	<table cellpadding="0" cellspacing="0" border="0">
	<tr>
	<td>
	
	<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="silver" bordercolorlight="black" bordercolordark="white" bgcolor="#C0C0C0" style="border-style: outset">
		<tr>
			<td bgcolor="#cfcfcf" onselectstart="window.event.returnValue = false" id="tdWeather" align="center" style="height:16;width:46px"><div language="javascript" id="divWeather" onmouseup="tdWeather.style.borderStyle=''" onmouseout="tdWeather.style.borderStyle=''" onmousedown="tdWeather.style.borderStyle='outset'" style="font-family:tahoma;font-size:11px;color:darkred;cursor:hand;width:46px" onclick="cmdWeather_onclick()">Weather</div></td>
			<td bgcolor="#cfcfcf" onselectstart="window.event.returnValue = false" id="tdMovies" align="center" style="height:16;width:44px"><div language="javascript" id="divMovies" onmouseup="tdMovies.style.borderStyle=''" onmouseout="tdMovies.style.borderStyle=''" onmousedown="tdMovies.style.borderStyle='outset'" style="font-family:tahoma;font-size:11px;color:darkred;cursor:hand;width:44px" onclick="cmdMovies_onclick()">Movies</div></td>
			<td bgcolor="#cfcfcf" onselectstart="window.event.returnValue = false" id="tdZagat" align="center" style="height:16;width:42px"><div language="javascript" id="divZagat" onmouseup="tdZagat.style.borderStyle=''" onmouseout="tdZagat.style.borderStyle=''" onmousedown="tdZagat.style.borderStyle='outset'" style="font-family:tahoma;font-size:11px;color:darkred;cursor:hand;width:42px" onclick="cmdZagat_onclick()">Zagat</div></td>
			<td id=tdTodaysDate height=24px>
				<table width=100%>
					<tr>
						<td align=center><input unselectable=on language="javascript" type="text" class="DateTimeFont" name="c" id="c" style="text-align: center; border-style: none; width:100%; background-color: transparent"></td>
						<!--td style="color:white;width:50px" class=DateTimeFont>(today)</td-->
					</tr>
				</table>
			</td>
			<td bgcolor="#cfcfcf" onselectstart="window.event.returnValue = false" id="tdFlights" align="center" style="height:16;width:42px"><div language="javascript" id="divFlights" onmouseup="tdFlights.style.borderStyle=''" onmouseout="tdFlights.style.borderStyle=''" onmousedown="tdFlights.style.borderStyle='outset'" style="font-family:tahoma;font-size:11px;color:darkred;cursor:hand;width:42px" onclick="cmdFlights_onclick()">Flights</div></td>
			<td bgcolor="#cfcfcf" onselectstart="window.event.returnValue = false" id="tdTickets" align="center" style="height:16;width:44px"><div language="javascript" id="divCalc" onmouseup="tdTickets.style.borderStyle=''" onmouseout="tdTickets.style.borderStyle=''" onmousedown="tdTickets.style.borderStyle='outset'" style="font-family:tahoma;font-size:11px;color:darkred;cursor:hand;width:44px" onclick="cmdTickets_onclick()">Tickets</div></td>
			<td bgcolor="#cfcfcf" onselectstart="window.event.returnValue = false" id="tdToday" align="center" style="height:16;width:46px"><div language="javascript" id="divToday" onmouseup="tdToday.style.borderStyle=''" onmouseout="tdToday.style.borderStyle=''" onmousedown="tdToday.style.borderStyle='outset'" style="font-family:tahoma;font-size:11px;color:darkred;cursor:hand;width:46px" onclick="GotoTodaysDate()">Today</div></td>
		</tr>
	</table>
	
	</td>
	</tr>
	<tr>
	<td height="<%=ah-115%>" width="<%=intTaskWidth%>">
	<iframe style="display:none" id="frameTaskPad1" name="frameTaskPad1" src="nullSrc()" allowTransparency="true" wwidth="424" width="<%=intTaskWidth%>" hheight="450" height="<%=ah-103%>" scrolling="yes" align="top" frameborder="0" framespacing="0"></iframe>
	<iframe style="display:none" id="frameTaskPad2" name="frameTaskPad2" src="nullSrc()" allowTransparency="true" wwidth="424" width="<%=intTaskWidth%>" hheight="450" height="<%=ah-103%>" scrolling="yes" align="top" frameborder="0" framespacing="0"></iframe>
	
	</td>
	</tr>
	</table>
</td>

<td style="border-style:solid;border-color:black" valign="top">

<table cellpadding="0" cellspacing="2" border="0">
<tr>
<td>
	<div valign="top" name="CalDIV" Id="CalDIV" style="display:inline"></div>
		
	<script language="javascript1.2">
	var days = new Array();
	days=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
	var months = new Array();
	months = ['January','February','March','April','May','June','July','August','September','October','November','December'];

	function formatDate(dt)
	{
		var d = new Date(dt);
		return days[d.getDay()] +', ' + months[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
		d = null;
	}

	var Calobj = new Calendar('Cal');

	Calobj.render();
	Calobj.onDateChange = function (new_date,x) 
								{
									ClearTO();
									var dt = new Date('<%=now()%>');
									var tmpTime = (dt.getMonth()+1).toString()+(dt.getDate()).toString()+(dt.getYear().toString())+(dt.getHours()).toString()+(dt.getMinutes()).toString()+(dt.getSeconds()).toString()+(Math.random()*10000).toString()
									disableMenu();
									refreshTaskPad('TaskPad.asp?TargetDate=' + new_date + '&LoadCal=False' + '&App=' + x + '&v=' + tmpTime);
									//alert(new_date)
									Calobj.setDate(new_date);
									var tdate = new Date(new_date);
									c.value = formatDate(tdate);
									if((tdate.getMonth()+1).toString()+tdate.getDate().toString()+tdate.getYear().toString()==(dt.getMonth()+1).toString()+(dt.getDate()).toString()+(dt.getYear().toString()))
										setToday(true);
									else
										setToday(false);
									
									StartTimer()
								}
								
	Calobj.onPageChange = function (f,t) 
								{
									document.parentWindow.frames('frameCalUpdate').location = 'CalendarUpdate.asp?CalFrom='+f+'&CalTo='+t
								}
	</script>
<iframe src="nullSrc()" id="frameCalUpdate" style="visibility:hidden;height:0px;width:0px;"></iframe>
</td>
</tr>
<tr>
	<td align="center" valign="middle" id="tdBanner" height="148px">
		<!--iframe allowTransparency="true" scrolling="no" id="frameBanner" src="BannerHolder.asp" width="155" height="148"></iframe-->
		<iframe frameborder="no" allowTransparency="true" scrolling="no" id="frameBanner" src="BannerHolder.asp" width="155" height="<%=ah-235%>"></iframe>
	</td>
</tr>
<tr>
<td>
<%
'If remote.Session("ACT") Then
'Response.Write "<center><input disabled id=cmdAddNewTask name=cmdAddNewTask type=button value=""Add Calendar Task"" Class=BlueFont style=""WIDTH: 155px"" size=22 onmouseover=""cmdAddNewTask.style.borderWidth=4"" onmouseout=""cmdAddNewTask.style.borderWidth=2""></center>"
'End If
%>
</td>
</tr>
</table>	

</td>
</tr>
</table>

</td>
</tr>
</table>

<!--/div-->
<div ID="tooltip" STYLE="left:0px;font-family: Helvetica; font-size: 8pt; position: absolute; z-index: 200; visibility: hidden; width:290px;">
	<iframe height="402" width="620" frameborder="0" style="border-style: none; border-width: 1px;" src="tooltip.asp" id="frameToolTip" scrolling="no"></iframe>
</div>
<div ID="divReminder" STYLE="left:0px;font-family: Helvetica; font-size: 8pt; position: absolute; z-index: 200; visibility: hidden; width:290px;">
	<iframe height="404" width="620" frameborder="0" style="border-style: none; border-width: 1px;" src="ReminderSummary.asp" id="frameReminder" scrolling="no"></iframe>
</div>

<div id="divLoading" name="divLoading" style="position: absolute; top: 200; left: 300; z-index: 201; visibility: hidden">
	<iframe height="69" width="174" frameborder="0" style="border-style: none; border-width: 1px;" src="LoadingDiv.asp?v=1" id="frameLoadingDiv" allowTransperancy="true" scrolling="no"></iframe>
</div>

<div id=divCallOut style="visibility:hidden;position:absolute;top:0;left:<%=calloutleft%>;height:145;width:200"><table cellpadding=10 valign=top align=center height=100% width=86% background=images/QLCallOutSquare.gif><tr><td valign=top id=tdCallOutText>Text will go here...</td></tr></table></div>

</body>
</html>

<script LANGUAGE="JavaScript">
function GotoTodaysDate()
{
	
	var dt = new Date('<%=now()%>');
	var curdt = (dt.getMonth()+1)+'/'+dt.getDate()+'/'+dt.getYear();
	window.top.Calobj.onDateChange (curdt);
	
}

document.all("divLoading").style.pixelTop = (<%=GridHeight+intHeightIncrement%>-142)/2;
document.all("divLoading").style.pixelLeft = (screen.availWidth-174)/2;
document.all("divLoading").style.visibility = "visible";
var booResume = true;
var intBadCount = 0;
var intATF = 1;
//var b = 0;
//refreshTaskPad("<%=strSource%>v=" + tmpTime)

function refreshTaskPad( strPage )
{
	if(booResume)
	{
		intBadCount = 0;
		try
		{
			if(intATF==2)
			{
				if(window.frames("frameTaskPad2").txtTaskPadLoaded.value == "OK")
					intATF = 1;
			}
			else
			{
				if(window.frames("frameTaskPad1").txtTaskPadLoaded.value == "OK")
					intATF = 2;
			}
			window.status = "Refresh succeeded.";
		} catch(e) {}
	}
	else
	{
		window.status = "Refresh pending."
		intBadCount++;
	}
		
	if(intBadCount==3)
	{
		alert('There is a problem with your connection.  Please click OK to update your calendar.');
		window.location.reload();
		intBadCount = 0;
	}
	else
	{
		var myFrame = "frameTaskPad"+intATF
		window.frames(myFrame).location = strPage + "&pad=" + intATF  //+"&b="+b++;
	}
	booResume = false;
}

function viewOK()
{
	if(intATF==2)
	{
		//document.all("frameTaskPad2").style.visibility = "visible";
		//document.all("frameTaskPad1").style.visibility = "hidden";
		document.all("frameTaskPad1").style.display = "none";
		document.all("frameTaskPad2").style.display = "inline";
	}
	else
	{
		//document.all("frameTaskPad2").style.visibility = "hidden";
		//document.all("frameTaskPad1").style.visibility = "visible";
		document.all("frameTaskPad2").style.display = "none";
		document.all("frameTaskPad1").style.display = "inline";
	}
	booResume = true;
}


function TimerFunction(t)
{
	// This is the logic to check if the Taskpad needs to be refreshed.

try
 {	
	var tmpDate = window.Calobj.getVal();
	if (tmpDate == "")
	{
		var dt = new Date('<%=now()%>');
		    tmpDate = (dt.getMonth()+1) + '/' + dt.getDate() + '/' + dt.getFullYear();
    }
	var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
	xmlHttp.open("POST", "TaskPadCheckForUpdate.asp?ukey=<%=uKey%>&date=" + escape(tmpDate), false)
	var tags = document.frames[("frameTaskPad"+intATF)].document.all.tags("input");
	var str = '';
	for(var xx = 0;xx < tags.length;xx++)
		{
		var fpos = tags[xx].id.indexOf("txtTaskText");
		if(fpos==0)
			str += tags[xx].id.substr(11) + ","
		}

	// required for form sending
	xmlHttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
	//
	xmlHttp.send("ids="+str)
	
	//xmlHttp.send();
	//alert(xmlHttp.responseText )
	
if (parseInt(unescape(xmlHttp.responseText)) > 0)
{	

	try
	{
	window.status = "Refreshing Tasks..."
	var tmpDate = window.Calobj.getVal();
	//alert(tmpDate)
	var dt = new Date('<%=now()%>');
 	// create cache busting parameter...
 	var tmpTime = (dt.getMonth()+1).toString()+(dt.getDate()).toString()+(dt.getYear().toString())+(dt.getHours()).toString()+(dt.getMinutes()).toString()+(dt.getSeconds()).toString()+(Math.random()*10000).toString()
	
	if (t) {
			if (tmpDate == "")
				var tmpDate = (dt.getMonth()+1) + '/' + dt.getDate() + '/' + dt.getFullYear();
			refreshTaskPad("TaskPad.asp?TargetDate=" + tmpDate + "&v=" + tmpTime);
			c.value = formatDate(tmpDate)
			}
	else
			{ 
			var dtparam = (dt.getMonth()+1) + '/' + dt.getDate() + '/' + dt.getFullYear();
			refreshTaskPad("TaskPad.asp?TargetDate=" + dtparam + "&v=" + tmpTime);
			c.value = formatDate(dtparam)
			}
	window.status = "Ready"
	}
	catch(e) {}
	}

}
	catch (e)
	{}
	

}
function ClearTO()
{
	window.clearTimeout(TimerID);
}

function calViewChange(o)
{
	var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
	xmlHttp.open("POST", "SetDefaultCalViewID.asp?id="+o.value, false)
	xmlHttp.send()
	window.top.Calobj.onDateChange(window.top.Calobj.getVal());
	xmlHttp = null
}

function departmentChange(o)
{
	var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
	xmlHttp.open("POST", "SetDefaultDepartmentID.asp?id="+o.value, false)
	xmlHttp.send()
	window.top.Calobj.onDateChange(window.top.Calobj.getVal());
	xmlHttp = null
}

function setToday(booToday)
{
	if(booToday)
		{
		document.all("tdTodaysDate").style.backgroundColor = "#00AE57";
		document.all("c").style.color = "white";
		document.all("c").style.fontWeight = "bold";
		document.all("divToday").innerText = "Today";
		document.all("divToday").disabled = true;
		}
	else
		{
		document.all("tdTodaysDate").style.backgroundColor = "";
		document.all("c").style.color = "black";
		document.all("c").style.fontWeight = "normal";
		document.all("divToday").innerText = "Today";
		document.all("divToday").disabled = false;
		}
}
</script>

