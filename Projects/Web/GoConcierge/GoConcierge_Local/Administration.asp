<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))


If remote.Session("ScreenHeight") < 750 Then
	searchWinW = 756
	searchWinH = 518
Else
	searchWinH = 632
	searchWinW = 990
End If

dim strStatus
if Request.Cookies("AllowEMailStatus") = "" or Request.Cookies("AllowEMailStatus") = "no" then
	strStatus = "yes"
else
	strStatus = "no"
end if
%>
<!--#INCLUDE file="checkuser.asp"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<title>Administration</title>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Dim CookieStatus

sub cmdGetNIL_onclick
	window.open "GetNotInListLocations.asp","","center=yes,Width=600px,height=460px,toolbar=no,menubar=yes,scrollbars=yes"
end sub
sub cmdBackupServers_onclick
	window.open "server_phil_validate.asp","ServerBackups"
end sub
Sub cmdHotelSetup_onclick
	document.parentWindow.location.href = "HotelSetup.asp?Mode=S"
End Sub

Sub cmdCS_onclick
	document.parentWindow.location.href = "CSSwitchBoard.asp"
End Sub

sub cmdAllowEmailEdit_onclick
	dim url, x
	url = "<%=Application("HomePage")%>/AllowEMailEdit.asp?Status=" & CookieStatus
	set x = createobject("Microsoft.XMLHTTP")
	x.open "get", url, false
	x.setRequestHeader "Content-Type", "application/x-www-form urlencoded"
	x.send()
	if CookieStatus = "yes" then
		CookieStatus = "no"
		window.cmdAllowEmailEdit.value = "Disallow E-Mail Location Edits"
	else
		CookieStatus = "yes"
		window.cmdAllowEmailEdit.value = "Allow E-Mail Location Edits"
	end if
	set x = nothing
end sub

sub cmdResetReport_onclick
	
	
	x = showModalDialog ("ResetReports.asp","","center:yes;status:no;scrollbars:no;dialogHeight:350px;dialogWidth:400px;")
	
	if x="1" then
			dim url, x
			url = "ResetReports.asp?action=reset"
			set x = createobject("Microsoft.XMLHTTP")
			x.open "get", url, false
			x.setRequestHeader "Content-Type", "application/x-www-form urlencoded"
			x.send()
			if x.responseText = "" Then
				'alert("The object has been reset")
			Else
				'alert(x.responseText)
			End If
			set x = nothing
	end if
end sub


Sub cmdAddEditCategory_onclick
	document.parentWindow.location.href = "AddCat.asp"
End Sub

Sub cmdGoBack_onclick
  dim x
  Randomize 100
  x = 2*rnd(100)
  window.location = "Switchboard3.asp?x=" & cstr(x) & "&CalledFrom=Admin"
End Sub

Sub cmdCustomLocationEdit_onclick
  window.showModalDialog "BrowseLocationsFrame.asp?Mode=Edit","","status:no;DialogHeight:<%=searchWinH%>px;DialogWidth:<%=searchWinW%>px;scroll:no;center:yes"
End Sub

Sub cmdMarqueeAdmin_onclick
  window.parent.location.href = "MarqueeAdmin.asp"
End Sub

Sub cmdRoomMaintenance_onclick
  window.parent.location.href = "RoomSetup.asp"
End Sub

Sub cmdUserSetup_onclick
  window.parent.location.href = "UserSetupNew.asp"
End Sub

Sub cmdGlobalSpellCheckUtility_onclick
  Msgbox "Yet to be implemented."
End Sub

Sub cmdMissingLocationDataReport_onclick
  window.parent.location.href = "ReportMissingLocationData.asp"
End Sub

Sub cmdLocationPrintsReport_onclick
  window.parent.location.href = "ReportLocationPrints.asp"
End Sub

Sub cmdNewLocationWorksheet_onclick
  window.parent.location.href = "ReportNewLocationWorksheet.asp"
End Sub

Sub cmdLocationReportFieldSuppression_onclick
  window.parent.location.href = "SetupSuppression.asp"
End Sub

Sub cmdSetupActionTypes_onclick
  window.parent.location.href = "SetupActionTypes2.asp"
End Sub

Sub cmdSetupActions_onclick
  window.parent.location.href = "SetupActions.asp"
End Sub

Sub cmdDirectionsEdit_onclick
  window.parent.location.href = "DirectionsMain.asp"
End Sub

Sub cmdVenueSetup_onclick
  window.parent.location.href = "VenueSetup.asp"
End Sub

sub cmdSetupCalView_onclick
	window.showModalDialog "SetupCalViewMain.asp","","dialogHeight:390px;dialogWidth:430px;status:no;scroll:no;center:yes"
end sub


sub cmdManageKeyWords_onclick
	window.showModalDialog "KeyWordMain.asp","","dialogHeight:790px;dialogWidth:670px;status:no;scroll:no;center:yes"
end sub

sub cmdManageAreas_onclick
	window.showModalDialog "AreasEdit.asp","","dialogHeight:570px;dialogWidth:340px;status:no;scroll:no;center:yes"
end sub

sub cmdCustomReports_onclick
	window.showModalDialog "CustomReports\TemplateSelect.asp","","dialogHeight:390px;dialogWidth:430px;status:no;scroll:no;center:yes"
end sub



Sub cmdSwitchCompany_onclick
  window.parent.location.href = "SelectCompany.asp"
End Sub

Sub cmdFunButtons_onclick
  window.showModalDialog "SetupFunButtons.asp","","dialogHeight:390px;dialogWidth:636px;center:yes;scroll:no;status:no;"
End Sub
'sub cmdAssignLocsToHotels_onclick
  'window.parent.location.href = "AssignLocationsToHotels.asp"
'  window.parent.location.href = "AssignLocations/AssignLocations.asp"
'End Sub

sub cmdSetupHotels_onclick
  window.parent.location.href = "HotelSetupNewMain.asp"
End Sub

sub cmdSetupCompanies_onclick
  window.parent.location.href = "SetupCompanies.asp"
End Sub

sub cmdSetupGroups_onclick
  window.parent.location.href = "AssignCompanyToHotelGroup.asp?CompanyID=-1"
End Sub

sub cmdstates_onclick
  window.parent.location.href = "SearchbyState.asp"
End Sub

sub cmdEditTemplates_onclick
	 window.parent.location.href = "EditTemplate.asp"
End sub

sub cmdSetupAppointmentNotes_onclick
	x = window.showModalDialog("SetupAppointmentNotes.asp?CompanyID=<%=remote.Session("CompanyID")%>","","center:yes;scroll:no;dialogHeight:420px;dialogWidth:500px;status:no")
	' window.parent.location.href = "SetupAppointmentNotes.asp"
End sub

Sub window_onload
	On Error Resume Next
	<%if strStatus = "yes" then%>
		window.cmdAllowEmailEdit.value = "Allow E-Mail Location Edits"
		CookieStatus = "yes"
	<%else%>
		window.cmdAllowEmailEdit.value = "Disallow E-Mail Location Edits"
		CookieStatus = "no"
	<%end if%>
End Sub

-->
</SCRIPT>
</HEAD>

<style>
	<!--
	.Label	{ font-family: Tahoma; font-size: 11; margin-left: 5 }
	.BUTTON	{ height:20px;width:210px;font-family: Tahoma; font-size: 11 }
	td		{ background-color: #b1b1cb }
	-->
</style>

<body bgcolor=silver topmargin="8" leftmargin="5" marginwidth="0" marginheight="0" link="black" vlink="black" alink="black"><!--#include file = "Header.inc" ---> 

<%
'Response.Write "remote.Session(""FloatingUser_SuperUser""):" & remote.Session("FloatingUser_SuperUser") & ".<br>"

if remote.Session("FloatingUser_SuperUser") then
	Response.Write "<BR>"
end if%>
<center>
<div style="overflow:auto;height:<%=searchWinH-8%>px;width:<%=searchWinW%>px">
<table align="center" style="border-style: solid; border-width: 1px; border-color: black;" cellSpacing=0 cellPadding=1 id=TABLE1>
<tr>
	<td style="border-style: outset; border-width: 2px;" valign="top" colspan=2 background="images/Background_Pinstripe.jpg" align=center>
		<font face="Tahoma" size=4 color=black>A d m i n i s t r a t i o n</font>
	</td>
</tr>
<td valign="top">
<INPUT Class=BUTTON id=cmdUserSetup name=cmdUserSetup type=button value="User Setup">
</td>
<td valign="top">
	<p class="Label">Add and edit user names and passwords. Apply security level.</p>
</td>
</tr>
<!--tr>
<td valign="top">
<INPUT Class=BUTTON id=cmdTaskSearch name=cmdTaskSearch type=button value="Task Search">
</td>
<td valign="top">
	<p class="Label">Search for tasks based on a wide range of criteria.</p>
</td>
</tr-->
<tr>
<td valign="top">
<INPUT Class=BUTTON id=cmdCustomLocationEdit name=cmdCustomLocationEdit type=button value="Location Edit">
</td>
<td valign="top">
	<p class="Label">Customize Location information specific to your hotel.</p>
</td>
</tr>

<tr>
	<td valign="top">
		<INPUT Class=BUTTON id=cmdSetupActions name=cmdSetupActions type=button value="Setup Actions/Notes">
	</td>
	<td>
		<p class="Label">Populate the listing of Actions/Notes used in Task assignment.</p>
	</td>
</tr>

<!--tr>
	<td valign="top">
		<INPUT Class=BUTTON id=cmdSetupAppointmentNotes name=cmdSetupAppointmentNotes type=button value="Setup Appointment Notes">
	</td>
	<td>
		<p class="Label">Setup the master list of available Appointment Note fields.</p>
	</td>
</tr-->

<%if remote.Session("FloatingUser_SuperUser") then%>

<!--	<tr>
		<td valign="top">
		<INPUT Class=BUTTON id=cmdVenueSetup name=cmdVenueSetup type=button value="Event Venues Setup">
		</td>
		<td valign="top">
			<p class="Label">Customize Venues For Events Display</p>
		</td>
	</tr>
	
		<tr>
		<td valign="top">
		<INPUT Class=BUTTON id=cmdCS name=cmdCS type=button value="CitySearch Setup">
		</td>
		<td valign="top">
			<p class="Label">Setup CitySearch</p>
		</td>
	</tr>  -->



	<tr>
		<td valign="top">
		<INPUT Class=BUTTON id=cmdDirectionsEdit name=cmdDirectionsEdit type=button value="Directions Edit">
		</td>
		<td valign="top">
			<p class="Label">Customize Driving Directions Text.</p>
		</td>
	</tr>
<!--	<tr>
		<td valign="top">
			<INPUT Class=BUTTON id=cmdMissingLocationDataReport name=cmdMissingLocationDataReport type=button value="Missing Location Data Report">
		</td>
		<td valign="top"><p class="Label">This report lists all the Locations that have missing data.</p>
		</td>
	</tr> -->

<!-- <tr>
<td valign="top">
<INPUT Class=BUTTON id=cmdLocationPrintsReport name=cmdLocationPrintsReport type=button value="Location Prints Report">
</td>
<td>
<p class="Label">This report counts the number of times a location has been printed based on user-defined criteria.</p>
</td>
</tr> -->

	<tr>
		<td valign="top">
			<INPUT Class=BUTTON id=cmdFunButtons name=cmdFunButtons type=button value="Setup Quick Links">
		</td>
		<td valign="middle"><p class="Label">Setup the Switchboard "Quick Link" Buttons.</p>
		</td>
	</tr>

<!-- <tr>
<td valign="top">
<INPUT Class=BUTTON id=cmdNewLocationWorksheet name=cmdNewLocationWorksheet type=button value="New Location Worksheet">
</td>
<td>
<p class="Label">This worksheet allows users to log Locations that are to be entered by the Administrator at a later date.</p>
</td>
</tr> -->

<!--tr>
<td valign="top">
<% 
	'If remote.Session("FloatingUser_SuperUser") Then 
	'	Response.Write "<INPUT Class=BUTTON id=cmdLocationReportFieldSuppression name=cmdLocationReportFieldSuppression type=button value=""Location Report Field Suppression"">"
	'else
	'	Response.Write "<INPUT Class=BUTTON id=cmdLocationReportFieldSuppression name=cmdLocationReportFieldSuppression type=button value=""Location Report Field Suppression"" disabled>"
	'End If 
%>
</td>
<td>
<p class="Label">Configure the Location Report to suppress selected fields no data.</p>
</td>
</tr-->

<tr>
<td valign="top">
<% 
	If remote.Session("FloatingUser_SuperUser") Then 
		Response.Write "<INPUT Class=BUTTON id=cmdSetupActionTypes name=cmdSetupActionTypes type=button value=""Setup Action Types"">"

	else
		Response.Write "<INPUT Class=BUTTON id=cmdSetupActionTypes name=cmdSetupActionTypes type=button value=""Setup Action Types"" disabled>"
	End If 
%>
</td>
<td>
<p class="Label">Populate the listing of Action Types used in Task assignment.</p>
</td>
</tr>

<!-- <tr>
<td valign="top">
<% 
	If remote.Session("FloatingUser_SuperUser") Then 
		Response.Write "<INPUT Class=BUTTON id=cmdMarqueeAdmin name=cmdMarqueeAdmin type=button value=""Marquee Admin"">"

	else
		Response.Write "<INPUT Class=BUTTON id=cmdMarqueeAdmin name=cmdMarqueeAdmin type=button value=""Marquee Admin"" disabled>"
	End If 
%>
</td>
<td>
<p class="Label">Admin the Marquee Messages.</p>
</td>
</tr> 

<tr>
<td valign="top">
<% 
	If remote.Session("FloatingUser_SuperUser") Then 
		Response.Write "<INPUT Class=BUTTON id=cmdRoomMaintenance name=cmdRoomMaintenance type=button value=""Room Maintenance"">"

	else
		Response.Write "<INPUT Class=BUTTON id=cmdRoomMaintenance name=cmdRoomMaintenance type=button value=""Room Maintenance"" disabled>"
	End If 
%>
</td>
<td>
<p class="Label">Populate the room dropdown.</p>
</td>
</tr> -->

<tr>
<td valign="top">
<% 
	If remote.Session("FloatingUser_SuperUser") Then 
		Response.Write "<INPUT Class=BUTTON id=cmdSwitchCompany name=cmdSwitchCompany type=button value=""Switch to Another Company"">"

	else
		Response.Write "<INPUT Class=BUTTON id=cmdSwitchCompany name=cmdSwitchCompany type=button value=""Switch to Another Company"" disabled>"
	End If 
%>
</td>
<td>
<p class="Label">Switch to another company.</p>
</td>
</tr>

<tr>
<td valign="top">
<INPUT class="button" id=cmdBackupServers name=cmdBackupServers type=button value="Backup Server Status">
</td>
<td>
<p class="Label">Check the status of the on-site (future off-site as well) backup servers.</p>
</td>
</tr>
<tr>
<td valign="top">
<INPUT class="button" id=cmdGetNIL name=cmdGetNIL type=button value="Not In List Locations">
</td>
<td>
<p class="Label">Get all the locations that have been entered into a task yet do not exist in our location table.</p>
</td>
</tr>

<!--tr>
<td valign="top">
<INPUT id=cmdstates name=cmdstates type=button value="Choose States">
</td>
<td>
<p class="Label">Choose your states</p>
</td>
</tr-->
<tr>
<td valign="top">
<% 
	If remote.Session("FloatingUser_SuperUser") Then 
		Response.Write "<INPUT Class=BUTTON id=cmdSetupHotels name=cmdSetupHotels type=button value=""Setup Hotels"">" 
	else
		Response.Write "<INPUT Class=BUTTON id=cmdSetupHotels name=cmdSetupHotels type=button value=""Setup Hotels"" disabled>" 
	End If 
%>
</td>
<td>
<p class="Label">Setup any of our client hotels.</p>
</td>
</tr>
<!--tr>
<td valign="top">
<% 
'	If remote.Session("FloatingUser_SuperUser") Then 
'		Response.Write "<INPUT Class=BUTTON id=cmdAssignLocsToHotels name=cmdAssignLocsToHotels type=button value=""Assign Locations to Hotels"">" 
'	else
'		Response.Write "<INPUT Class=BUTTON id=cmdAssignLocsToHotels name=cmdAssignLocsToHotels type=button value=""Assign Locations to Hotels"" disabled>" 
'	End If 
%>
</td>
<td>
<p class="Label">Assign multiple locations to multiple hotels.</p>
</td>
</tr-->

<tr>
<td valign="top">
<% 
	If remote.Session("FloatingUser_SuperUser") Then 
		Response.Write "<INPUT Class=BUTTON id=cmdSetupGroups name=cmdSetupGroups type=button value=""Add or Delete Hotel Groups"">"
	else
		Response.Write "<INPUT Class=BUTTON id=cmdSetupGroups name=cmdSetupGroups type=button value=""Add or Delete Hotel Groups"" disabled>"
	End If 
%>
</td>
<td>
<p class="Label">Add or delete hotel groups.</p>
</td>
</tr>

<tr>
<td valign="top">
<%
    If remote.Session("FloatingUser_SuperUser") Then
		Response.Write "<INPUT Class=BUTTON id=cmdEditTemplates name=cmdEditTemplates type=button value=""Edit Guest Report Templates"">"
	Else
		Response.Write "<INPUT Class=BUTTON id=cmdEditTemplates name=cmdEditTemplates type=button value=""Edit Guest Reports"" disabled>"
	End If
%>
</td>

<td>
<p class="Label">Edit Guest Task Report Templates.</p>
</td>
</tr>

<tr>
<td valign="top">
<%
    If remote.Session("FloatingUser_SuperUser") Then
		Response.Write "<INPUT Class=BUTTON id=cmdAddEditCategory name=cmdAddEditCategory type=button value=""Add/Edit Category"">"
	Else
		Response.Write "<INPUT Class=BUTTON id=cmdAddEditCategory name=cmdAddEditCategory type=button disabled value=""Add/Edit Category"">"
	End If
%>
</td>

<td>
<p class="Label">Add and Edit Category information.</p>
</td>
</tr>

<tr>
<td valign="top">
<%
    If remote.Session("FloatingUser_SuperUser") Then
		Response.Write "<INPUT Class=BUTTON id=cmdHotelSetup name=cmdHotelSetup type=button value=""Setup This Hotel"">"
	Else
		Response.Write "<INPUT Class=BUTTON id=cmdHotelSetup name=cmdHotelSetup type=button value=""Setup This Hotel"" disabled>"
	End If
%>
</td>
<td>
<p class="Label">Edit Hotel information for <%=remote.Session("CompanyName")%>.</p>
</td>
</tr>
<tr>
<td valign="top">
<%
    If remote.Session("FloatingUser_SuperUser") Then
		Response.Write "<INPUT Class=BUTTON id=cmdSetupCalView name=cmdSetupCalView type=button value=""Setup Calendar Views"">"
	End If
%>
</td>
<td>
<p class="Label">Edit Calendar Views for <%=remote.Session("CompanyName")%>.</p>
</td>
</tr>

<tr>
<td valign="top">
<%
    If remote.Session("FloatingUser_SuperUser") Then
		Response.Write "<INPUT Class=BUTTON id=cmdManageKeyWords name=cmdManageKeyWords type=button value=""Manage Keywords"">"
	End If
%>
</td>
<td>
<p class="Label">Manage Keywords</p>
</td>
</tr>

<tr>
<td valign="top">
<%
    If remote.Session("FloatingUser_SuperUser") Then
		Response.Write "<INPUT Class=BUTTON id=cmdManageAreas name=cmdManageAreas type=button value=""Manage Areas"">"
	End If
%>
</td>
<td>
<p class="Label">Manage Areas</p>
</td>
</tr>

<%If remote.Session("FloatingUser_SuperUser") Then%>
<tr>
<td valign="top">
	<INPUT Class=BUTTON id=cmdSetupLookups name=cmdSetupLookups type=button value="Setup Lookup Tables">
</td>
<td>
<p class="Label">Setup Lookup Tables</p>
</td>
</tr>
<%End If%>

<tr>
<td valign="top">
<%
    If remote.Session("FloatingUser_SuperUser") Then
		Response.Write "<INPUT Class=BUTTON id=cmdCustomReports name=cmdCustomReports type=button value=""Manage Custom Reports"">"
	End If
%>
</td>
<td>
<p class="Label">Custom Reports</p>
</td>
</tr>


<tr>
<td valign="top">
<INPUT Class=BUTTON style=color:purple id=cmdAllowEmailEdit name=cmdAllowEmailEdit type=button value="Allow E-Mail Location Edits">
</td>
<td>
<p class="Label">Allow this computer to edit locations from request e-mails.</p>
</td>
</tr>

<tr>
<td valign="top">
<INPUT Class=BUTTON style=color:purple id=cmdResetReport name=cmdResetReport type=button value="Reset Reports">
</td>
<td>
<p class="Label">Reset Report Object.</p>
</td>
</tr>




<%end if%>
<tr>
<td valign="top">
<%
	Response.Write "<INPUT Class=BUTTON id=cmdGoBack name=cmdGoBack style=""COLOR: #007500;"" type=button value=""Back to Home Page"">"
%>
</td>
<td>
<p class="Label">Return to Main Menu.</p>
</td>
</tr>
</table>
<!--
<font face="Tahoma" size="2"><A href="SwitchBoard3.asp" >Back to Home Page</A></font>
-->
</div>
</center>
</body>
</HTML>
