<%@ Language=VBScript %>
<%
Response.AddHeader "Content-Encoded", "gzip"
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))

If Request.Cookies ("CompanyID") = "" Then Response.Cookies ("CompanyID") = remote.Session ("CompanyID")

If Request.Cookies ("CompanyID") = remote.Session ("CompanyID") Then
	'Response.Write "Pad: OK" 
Else
	remote.Session ("CompanyID") = Request.Cookies ("CompanyID")
	emailGCN()
	'Response.Write "Pad: NOT OK" 
	' Do Emailing....
End If	

CompanyID = remote.Session ("CompanyID")




'remote.Session.Timeout = 60*24

dim intDaysPlus 
dim intDaysMinus
intDaysPlus = 32
intDaysMinus = 32

tmpACT = remote.Session("ACT")
tmpSuperUser = remote.Session("SuperUser")
rsah = remote.session("AvailHeight")
'if cint(Request.QueryString("b")) >= 2 then
'	Response.End
'end if
%>

<!--#INCLUDE file="include/vbfunc.asp"-->

<script LANGUAGE="vbscript" RUNAT="Server">
Server.ScriptTimeout = 2147483647
Function FormatDateTimeJA(pdtIn)

			Select Case remote.Session ("TimeFormat") 

				Case 1 ' if Military
							strOut = pdtIn
				Case Else ' All other instances
				
					Select Case pdtIn
						Case "00:00"
							strOut = "12 am"
						Case "12:00"
							strOut = "Noon"
						Case Else
							intHour = cInt(Left(pdtIn,2))
							If intHour > 12 Then
								intHour = intHour - 12
								strOut = cStr(intHour) & " pm"
							else
								strOut = cStr(intHour) & " am"
							End If
							'strOut = strOut & ":00"
					End Select
					
			End Select ' end of the outer select

				FormatDateTimeJA = strOut
End Function

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
</script>

<%
dim intTaskHeight, intTaskRowWidth, vbDarkGreen

vbDarkGreen = rgb(0,174,0)

intTaskRowWidth = 408 '424

Set cnSQL = Server.CreateObject("ADODB.Connection")
'Set cnSQLTemp = Server.CreateObject("ADODB.Connection")
'Set rsSQL = Server.CreateObject("ADODB.Recordset")
Set rsBooked = Server.CreateObject("ADODB.Recordset")
Set rsApp = Server.CreateObject("ADODB.Recordset")
'Set rsQuarter = Server.CreateObject("ADODB.Recordset")
  
cnSQL.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")
cnsql.CursorLocation = 3
%>
<html>
<head>
<!--META HTTP-EQUIV="REFRESH" Content="5"-->
<script src=CheckIfTaskExists.asp language=javascript></script>

<script LANGUAGE="JavaScript">

var  xx =0, tdbgColor = "", tdborderstyle = "";

parent.disableMenu()

// from below
function hide(intAppointmentID)
	{
	window.status = "";

		for (var i=0; i<24; i++)
		{
			try
				{
					//document.all("taskspan"+intAppointmentID+"t"+i).style.borderStyle="none";
					document.all("taskspan"+intAppointmentID+"t"+i).style.backgroundColor = tdbgColor;
					document.all("taskspan"+intAppointmentID+"t"+i).style.borderStyle = tdborderstyle;
				}
				catch (e)
				{
				}

		}
	}
		
function show(intAppointmentID, bSpan)
{
	if(document.all("txtTaskText"+intAppointmentID).value)
	window.status = document.all("txtTaskText"+intAppointmentID).value.replace(" <font style=background-color:red color=white>","").replace("</font> ","").replace("<b style=color:maroon>Reminder:</b>","Reminder:");
	//alert(document.all("taskspan"+intAppointmentID+"t"+i));
		
		for (var i=0; i<24; i++)
		{
			try
				{
					//document.all("taskspan"+intAppointmentID+"t"+i).style.borderWidth="1px";
					//document.all("taskspan"+intAppointmentID+"t"+i).style.borderStyle="solid";
					//document.all("taskspan"+intAppointmentID+"t"+i).style.borderColor="yellow";
					if(bSpan=='True') 
						{
						document.all("taskspan"+intAppointmentID+"t"+i).style.backgroundColor = "#FF00FF";
						//document.all("taskspan"+intAppointmentID+"t"+i).style.borderStyle = "outset"
						}
				}
				catch (e)
				{
		
				}
		}
}

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
		
function showTextActivate(strField) { 
	document.all("frameSTA").src = "TaskPadToolTip.asp?id=" + strField +"&UserID=" + <%=remote.session("FloatingUser_UserID")%>;
 }
 
 var booTTLoaded = true;
 
 function sta() { 
	if(booTTLoaded)
		booTTLoaded = false;
	else
		staContinue();
		//alert(unescape(window.frames("frameSTA").document.body.innerHTML));
 }

 function staContinue() {

  code = unescape(window.frames("frameSTA").document.body.innerHTML);
  if(code == "EOF")
  {
	alert ("The task you selected has been deleted by another user.  The calendar will now refresh.");
	parent.TimerFunction(1);
  }
  else
  {
	eval(code)

	var el = parent.document.all.tooltip //, a = new Array(), alen = 0;
	var rel = parent.document.all.divReminder //, a = new Array(), alen = 0;
	var frel = parent.document.frames.frameToolTip.document.all.tbdyToolTip
	var frid = parent.document.frames.frameToolTip.document.all.txtAppointmentID
	var booLastWasNote = false;
  
	el.style.pixelTop = (screen.availHeight-400)/2;
	el.style.pixelLeft = ((screen.availWidth-620)/2)-6;
	var j = frel.rows.length;
	for(i=0;i<j;i++)
		frel.deleteRow(0);
		  
	var str = "", cnt = 1;
  
	a = aTT;
	frid.value = a[0];
	strColor = a[1];

		  
	var oRow = frel.insertRow();
	var oCell = oRow.insertCell();
	oCell.innerText = "Task ID:";

	oCell = oRow.insertCell();
	oCell.innerText = a[0];
  
	var oRow = frel.insertRow();
	var oCell = oRow.insertCell();
	oCell.innerText = "Date:";
	oCell = oRow.insertCell();
	oCell.innerText = formatDate(a[2]);

	parent.document.frames.frameToolTip.txtReminderStatus.value = a[3];
  
	if(a[3] == " ")
 		parent.document.frames.frameToolTip.cmdReminder.value = "Add Reminder";
	else
		parent.document.frames.frameToolTip.cmdReminder.value = "Edit Reminder";
  
	if(strColor=="lightgreen" || strColor=="white")
		oRow.style.backgroundColor = "#C2F5C2";
	else
		oRow.style.backgroundColor = strColor;  	

	var cboStatus = parent.document.frames.frameToolTip.document.all("cboStatus");
	cboStatus.style.backgroundColor = strColor;
  
	for(i=4; i<a.length-1; i++)
	{
		var b = Array();
		b = a[i].split("|");
		b[1] = b[1].replace(/\&lt;<td>\&gt;/gi,"<<td>>").replace(/\&lt;<sq>\&gt;/gi,"<<sq>>")
		if(b[0]=='Status')
		{
			
			switch (b[1])
			{
				case 'Open' :
						cboStatus.selectedIndex = 0;
						break

				case 'Pending' :
						cboStatus.selectedIndex = 1;
						break

				case 'Closed' :
						cboStatus.selectedIndex = 2;
						break

				case 'Re-Confirmed' :
						cboStatus.selectedIndex = 3;
						break
				case 'Canceled' :
						cboStatus.selectedIndex = 4;
						break
				case 'Not Available' :
						cboStatus.selectedIndex = 5;
						break

				case 'Wait List' :
						cboStatus.selectedIndex = 6;
						break
						
						
			}
		}

		if(b[1].length > 0)
			{
			oRow = frel.insertRow();
				if(b[1].indexOf("<<td>>") == -1)
				{
					if(booLastWasNote)
						{
						if(cnt == 0)
							cnt = 1;
						else
							cnt = 0;
						} 
					if(cnt == 0)
						{
						if(strColor=="lightgreen" || strColor=="white")
							strColor = "#C2F5C2"
						oRow.style.backgroundColor = strColor; //"#C2F5C2";
						cnt++;
						}
					else
						{
						oRow.style.backgroundColor = "#FFFFE1";
						cnt = 0;
						}
					booLastWasNote = false;
				}
				else
				{
					if(cnt == 0)
						{
						if(strColor=="lightgreen" || strColor=="white")
							strColor = "#C2F5C2"
						oRow.style.backgroundColor = strColor; //"#C2F5C2";
						}
					else
						{
						oRow.style.backgroundColor = "#FFFFE1";
						}
					booLastWasNote = true;
				}

				if(b[0].indexOf("Reminder") > -1)
					oRow.style.backgroundColor = "#FCE6A3"

			oCell = oRow.insertCell();
			oCell.innerText = b[0]+":";
			oCell.vAlign = "top";
			oCell.width = "100px";
				
			oCell = oRow.insertCell();
			oCell.innerHTML = b[1].replace(/<<sq>>/g,"'").replace(/<<td>>/g,"")
			oCell.vAlign = "top"
			}
		}
  }
  
  var perm = a[a.length-1].split('|');
  
  
  
  var tooltip = parent.document.frames.frameToolTip.document;
  
  if (parseInt(perm[1]) > 0)
	{
		tooltip.all("cmdEdit").disabled = false
		tooltip.all("cmdDelete").disabled = false
		tooltip.all("cmdReminder").disabled = false
	}
  else
  {
	
		tooltip.all("cmdEdit").disabled = true
		tooltip.all("cmdDelete").disabled = true
		tooltip.all("cmdReminder").disabled = true
	
	}
	
  if (parseInt(perm[2]) > 0)
	{
		tooltip.all("cboStatus").disabled = false
	}
  else
  {
		tooltip.all("cboStatus").disabled = true
	
  }
	
	
	
	  
  
  
	
  el.style.visibility = "visible"
  rel.style.visibility = "hidden"
  booToolTipOpen = true;
}

function showActivate(str)
{
	var w = str.split('|');
	
	if (w[1]=='0')
		showTextActivate (w[0]);	
	else
		showReminderActivate  (w[0]);
	
}

function showReminderActivate(strField) { //, x, y) {
	document.all("frameReminderSTA").src = "ReminderToolTip.asp?id=" + strField;
	parent.document.frames.frameReminder.document.all.txtRAID.value = strField;
 }
 
 var booRTTLoaded = true;
 
 function staReminder() { 
	if(booRTTLoaded)
		booRTTLoaded = false;
	else
		staReminderContinue();
 }

 function staReminderContinue() {

  code = unescape(window.frames("frameReminderSTA").document.body.innerHTML);
  //alert(code)
  if(code == "EOF")
  {
	alert ("The reminder you selected has been deleted by another user.  The calendar will now refresh.");
	parent.TimerFunction(1);
  }
  else
  {
	eval(code)

	var el = parent.document.all.divReminder // for showing
	var ttel = parent.document.all.tooltip // for hiding
  
	var frameDoc = parent.document.frames.frameReminder.document.all;
	var frbd = frameDoc.bdy;
	var frio = frameDoc.imgBGOpen;
	var fric = frameDoc.imgBGClosed;
	var frel = frameDoc.tbdyReminder;
	var frdr = frameDoc.divReminderNote;
	var fren = frameDoc.tbdyReminderNote;
	var frid = frameDoc.txtAppointmentID;
	var frtb = frameDoc.divTaskBody;
  
	var booLastWasNote = false;
  
	el.style.pixelTop = (screen.availHeight-400)/2;
	el.style.pixelLeft = ((screen.availWidth-620)/2)-6;
	//el.style.pixelLeft = (screen.availWidth-540)/2;
	var j = frel.rows.length;
	for(i=0;i<j;i++)
		frel.deleteRow(0);
		  
	var str = "", cnt = 1;
  
	a = aTT;
	frid.value = a[0];
	strColor = a[1];

	// reminder note
	var oRow = frel.insertRow();
	oRow.style.height = 8;
	oCell = oRow.insertCell();
	oCell.colSpan = 2;
	var rn = a[3].replace(/\&lt;<sq>\&gt;/gi,"<<sq>>").replace(/<<sq>>/g,"'").replace(/\&amp;/,"&");
	if(rn.length > 0)
		frdr.innerText = (" - " + rn);
  
	// set height based on size of possibly scrolled reminder header.
	if(frdr.offsetHeight > 13)
		n = frdr.offsetHeight-10
	else
		n = 13;
	frtb.style.height = (334-(n-13));
  
	// task header
	var oRow = frel.insertRow();
	TaskHeaderRow = oRow;
	var oCell = oRow.insertCell();
	oCell.colSpan = 2;
	oCell.style.height = 24;
	oCell.style.textAlign = "center";
	oCell.innerText = "Task";
	oCell.style.fontWeight = "bold";

	var oRow = frel.insertRow();
	oRow.style.backgroundColor = strColor;
	var oCell = oRow.insertCell();
	oCell.innerText = "Task ID:";

	oCell = oRow.insertCell();
	oCell.innerText = a[0];
  
	var oRow = frel.insertRow();
	var oCell = oRow.insertCell();
	oCell.innerText = "Date:";
	oCell = oRow.insertCell();
	oCell.innerText = formatDate(a[2]);
  
	if(strColor=="#E3E1FF" || strColor=="white")
		oRow.style.backgroundColor = "" //"#C2F5C2";
	else
		oRow.style.backgroundColor = strColor;  	

	var cboStatus = parent.document.frames.frameReminder.document.all("cboStatus");
	var imgBellAni = parent.document.frames.frameReminder.document.all("imgBellAni");
	var imgBell = parent.document.frames.frameReminder.document.all("imgBell");
	cboStatus.style.backgroundColor = strColor;
			  
	for(i=4; i<a.length; i++)
	{
		var b = Array();
		b = a[i].split("|");
		b[1] = b[1].replace(/\&lt;<td>\&gt;/gi,"<<td>>").replace(/\&lt;<sq>\&gt;/gi,"<<sq>>")
		if(b[0]=='Status')
		{
			switch (b[1])
			{
				case 'Open' :
						cboStatus.selectedIndex = 0;
						imgBell.style.display = 'none';
						imgBellAni.style.display = 'inline';
						frbd.background = frio.src;
						TaskHeaderRow.style.backgroundColor = "#FF9595";
						break
				case 'Closed' :
						cboStatus.selectedIndex = 1;
						imgBell.style.display = 'inline';
						imgBellAni.style.display = 'none';
						frbd.background = fric.src;
						TaskHeaderRow.style.backgroundColor = "lightgreen";
						break
				case 'Canceled' :
						cboStatus.selectedIndex = 2;
						imgBell.style.display = 'inline';
						imgBellAni.style.display = 'none';
						frbd.background = fric.src;
						TaskHeaderRow.style.backgroundColor = "lightblue";
						break
			}
		}

		if(b[1].length > 0)
			{
			oRow = frel.insertRow();
				if(b[1].indexOf("<<td>>") == -1)
				{
					if(booLastWasNote)
						{
						if(cnt == 0)
							cnt = 1;
						else
							cnt = 0;
						} 
					if(cnt == 0)
						{
						if(strColor=="#E3E1FF" || strColor=="white")
							strColor = "#C2F5C2"
						oRow.style.backgroundColor = strColor; //"#C2F5C2";
						cnt++;
						}
					else
						{
						oRow.style.backgroundColor = strColor //"#FF9595" //"#E3E1FF";
						cnt = 0;
						}
					booLastWasNote = false;
				}
				else
				{
					if(cnt == 0)
						{
						if(strColor=="#E3E1FF" || strColor=="white")
							strColor = "#C2F5C2"
						oRow.style.backgroundColor = strColor; //"#C2F5C2";
						}
					else
						{
						oRow.style.backgroundColor = strColor //"#FF9595" //"#E3E1FF";
						}
					booLastWasNote = true;
				}
			oCell = oRow.insertCell();
			oCell.innerText = b[0]+":";
			oCell.vAlign = "top";
			oCell.width = "100px";
				
			oCell = oRow.insertCell();
			oCell.innerHTML = b[1].replace(/<<sq>>/g,"'").replace(/<<td>>/g,"")
			oCell.vAlign = "top"
			}
	}
	
	el.style.visibility = "visible"
	ttel.style.visibility = "hidden"
		  
	booReminderOpen = true;
	}
}

function window_onload() {
	window.scrollTo(0,window.screen.height)
	hide()
	//window.document.body.style.visibility = "visible"
	parent.viewOK();
}
</script>

<script LANGUAGE="vbscript" RUNAT="Server">
Public Function FormatAppointment(booShowLocation, pstrStartTime, pstrEndTime, pstrRoom, pstrSalutation, pstrName,pstrFirstName,pstrLastName, pstrActionType, pstrAction, pstrLocation, pstrSubject, pstrStatus, pstrNoTime, pstrNotesDetails, pbooReminder, pstrReminderNote)
	
	if pstrNoTime then
		strAppt = "Note: "
		strDash = ""
		strSpacer = "&nbsp;"
	else
		strAppt = CustomTime( pstrStartTime, "" )
		if pstrStartTime <> pstrEndTime then
			strAppt = strAppt & "-" & CustomTime( pstrEndTime, "" )
		end if
		strDash = " - "
		strSpacer = ""
	end if
		
	if pbooReminder then
		c = ""
		if pstrStatus="x" then
			c = "<font style=background-color:red color=white>&nbsp;Canceled&nbsp;</font>&nbsp;-&nbsp;"
		end if
		if pstrNoTime then
			strAppt = "<b style=color:maroon>Reminder:</b> " & c & pstrReminderNote
			strDash = " - "
			strSpacer = "&nbsp;"
		else
			strAppt = trim(CustomTime( pstrStartTime, "" )) & " - <b style=color:maroon>Reminder:</b> " & c & pstrReminderNote
			strDash = " - "
			strSpacer = ""
		end if
	else
		if len(trim(pstrRoom))=0 then 
			pstrRoom = null
		else
			pstrRoom = trim(pstrRoom)
		end if
		if len(trim(pstrSalutation))=0 then 
			pstrSalutation = null
		else
			pstrSalutation = trim(pstrSalutation)
		end if
		if len(trim(pstrName))=0 then 
			pstrName = null
		else
			pstrName = trim(pstrName)
		end if
		if len(trim(pstrFirstName))=0 then 
			pstrFirstName = null
		else
			pstrFirstName = trim(pstrFirstName)
		end if
		if len(trim(pstrLastName))=0 then 
			pstrLastName = null
		else
			pstrLastName = trim(pstrLastName)
		end if
		if len(trim(pstrActionType))=0 then 
			pstrActionType = null
		else
			pstrActionType = trim(pstrActionType)
		end if
		if len(trim(pstrAction))=0 then 
			pstrAction = null
		else
			pstrAction = trim(pstrAction)
			if pstrAction = "Restaurant Reservation" then
				pstrAction = "Restaurant"
			end if
		end if
		if len(trim(pstrLocation))=0 then 
			pstrLocation = null
		else
			pstrLocation = trim(pstrLocation)
		end if
	
		if len(trim(pstrSubject))=0 then 
			pstrSubject = null
		else
			'replace special html chars: <, >, &
			pstrSubject = replace(replace(replace(trim(pstrSubject),">","&gt;"),"<","&lt;"),"&","&amp;")
		end if

		if len(pstrSalutation) = 0 then
			vName = trim(pstrFirstName & " " & pstrLastName)
		else
			if len(pstrLastName) = 0 then
				vName = pstrFirstName
			else
				vName = pstrLastName
			end if
		end if
	
		if len(vName) = 0 then
			vName = null
		end if

		if pstrStatus="x" then
			strAppt = strAppt & strDash & "<font style=background-color:red color=white>&nbsp;Canceled&nbsp;</font>"
		end if

		strAppt = strAppt & (strDash + ((pstrSalutation & " ") + (vName)))
		strAppt = strAppt & (" - "+("Rm "+pstrRoom))
		strAppt = strAppt & (" - "+pstrActionType)
		strAppt = strAppt & (" - "+pstrAction)
	
		if booShowLocation then
			strAppt = strAppt & (" - "+pstrLocation)
		end if
	
		if Len(trim(pstrNotesDetails)) > 0 Then
			strAppt = strAppt & (" - ["+pstrNotesDetails +"]")
		End If
		strAppt = strAppt & (" - "+pstrSubject)
	end if
	
	FormatAppointment = strSpacer & trim(strAppt)
End Function

function CustomTime( strTime, padding )
	CustomTime = GCFormatTimeQ (strTime) 'right("0" & left(strTime,instrrev(strTime,":")-1),5) & lcase(right(strTime,2))
end function

</script>


<meta name="VI60_defaultClientScript" content=JavaScript>
<style type="text/css">
<!--
.opentask { color: #000000; font-size: 11px; font-family: Tahoma; text-decoration: none }
.colWhite { color: #FFFFFF; font-size: 11px; font-family: Tahoma; text-decoration: none }
.pendingtask { color: #000000; font-size: 11px; font-family: Tahoma; text-decoration: none }
.recontask { text-align:center;color: #FFFFFF; font-size: 11px; font-family: Tahoma; text-decoration: none; background-color:darkgreen;border-style:raised;border-width:1px }

a.opentask:link { color: #FFFFFF; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.opentask:vlink { color: #FFFFFF; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.opentask:link { color: #FFFFFF; font-size: 11px; font-family: Tahoma; text-decoration: none }

.closedtask { color: #000000; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.closedtask:link { color: #000000; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.closedtask:vlink { color: #000000; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.closedtask:link { color: #000000; font-size: 11px; font-family: Tahoma; text-decoration: none }

.notes { color: #000FFF; font-size: 11px; font-family: Tahoma; text-decoration: none }
.hours { color: #0000FF; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.hours:link { color: #0000FF; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.hours:vlink { color: #0000FF; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.hours:link { color: #0000FF; font-size: 11px; font-family: Tahoma; text-decoration: none }

.red { color: #FF0000; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.red:link { color: #FF0000; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.red:vlink { color: #FF0000; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.red:link { color: #FF0000; font-size: 11px; font-family: Tahoma; text-decoration: none }
.blue { color: #000080; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.blue:link { color: #000080; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.blue:vlink { color: #000080; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.blue:link { color: #000080; font-size: 11px; font-family: Tahoma; text-decoration: none }
.green { color: #007500; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.green:link { color: #007500; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.green:vlink { color: #0075000; font-size: 11px; font-family: Tahoma; text-decoration: none }
a.green:link { color: #007500; font-size: 11px; font-family: Tahoma; text-decoration: none }
}
-->
</style>


<script LANGUAGE="vbscript">
<!--

Sub jaDivClick(pstrUrl)
	if parent.cmdSearchbyCategory.value = "Search Locations" then
		str = "&v=" & right("0" & cstr(Month(Now())),2) & right("0" & cstr(Day(Now())),2) & right("0" & cstr(Year(Now())),2) & right("0" & cstr(Hour(Now())),2) & right("0" & cstr(Minute(Now())),2) & right("0" & cstr(Second(Now())),2)
		addTask pstrUrl & str
	end if
End Sub

function addTask(s)
	<%if remote.Session("BPW") then%>
		window.top.showTask s
	<%else%>
		dim strSQL
		startpos = instr(1,s,"ID=")+3
		endpos = instr(startpos,s,"&")
		aid = mid(s,startpos,endpos-startpos)
		strSQL = "ValidatePassword.asp?v=2&aid=" & aid & "&Caption=Add/Edit Task&TargetDate=" & window.top.calObj.getVal()
		x = showModalDialog(strSQL,"","center:yes;status:no;scrollbars:no;dialogHeight:116px;dialogWidth:298px;")
		if x <> "" then
			select case x
				case "Invalid Password"
					msgbox "Invalid password.",vbCritical,"Password"
				case "Invalid Department"
					msgbox "You do not have rights to modify a task for this department.",vbCritical,"Department Validation"
				case else
					window.top.showTask s
			end select
		end if
	<%end if%>
end function

' hide on every load in case user hit back button
' parent.divGuestTaskLetter.style.visibility = "hidden"
' parent.divSwitchboard.style.visibility = "visible"
parent.CalDIV.style.display = "inline"

-->
</script>

</head>
<body bgcolor="silver" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" link="blue" vlink="white" alink="blue" LANGUAGE=javascript onload="return window_onload()">
<!--#include file = "Header.inc" ---> 
<iframe onload=sta() src="LoadingAppointment.asp" id=frameSTA style=display:none;visibility:hidden></iframe>
<iframe onload=staReminder() src="LoadingAppointment.asp" id=frameReminderSTA style=display:none;visibility:hidden></iframe>
<%

'for i = 1 to Request.ServerVariables.Count
'	Response.Write Request.ServerVariables.Item(i) & vbcrlf
'next
'Response.Write "App: " & Request.QueryString("App")
'If remote.Session("ScreenHeight") < 750 Then
	intLoadTop = 140
	intLoadLeft = 140
'else
'	intLoadTop = 140
'	intLoadLeft = 140
'end if
aw = remote.session("AvailWidth")

'LeftColWidth = aw * .16
LeftColWidth = remote.Session("LeftColWidth")

intWidthFactor = aw-LeftColWidth-272
'intWidthFactor = aw-LeftColWidth-328

Response.Write "<table id=""tblMain"" name=""tblMain"" width=""" & intTaskRowWidth & """ style=""BORDER-RIGHT:medium none;BORDER-TOP:medium none;BACKGROUND:silver;BORDER-LEFT:medium none;BORDER-BOTTOM:medium none;BORDER-COLLAPSE:collapse;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0in 5.4pt 0in 5.4pt"" cellSpacing=""0"" cellPadding=""0"" bgColor=""silver"" border=""1"">"

        Dim TaskPadTargetDate
        
        Dim PrevApptID ' Added to avoid duplicate display of tasks in Calendar Screen...
        PrevApptID = -9 ' Some initial dummy value
        
        'Set TaskPadTargetDate
        If IsDate(Request.QueryString("TargetDate"))  Then
			TaskPadTargetDate = Request.QueryString("TargetDate")
        Else
        
			If IsDate(remote.Session("TargetDate")) Then
				TaskPadTargetDate = remote.Session("TargetDate")
			Else
				TaskPadTargetDate = Date()
			End If
        End If
        
        ' Response.write "<SCRIPT language=javascript>alert('" & TaskPadTargetDate & "')</SCRIPT>"

		'Set Start Time
		'dtStartTime = DateAdd("h",0,Month(TaskPadTargetDate) & "/" & Day(TaskPadTargetDate) & "/" & Year(TaskPadTargetDate))
		dtStartTime = cdate(Month(TaskPadTargetDate) & "/" & Day(TaskPadTargetDate) & "/" & Year(TaskPadTargetDate))

		'strSQL= "sp_GetAppointmentsWithRecurrence " & CompanyID & ", '" & TaskPadTargetDate & "', '" & DateAdd("D",1,TaskPadTargetDate) & "', 0"
		'strSQL= "sp_Appointments '" & TaskPadTargetDate & "', '" & DateAdd("D",1,TaskPadTargetDate) & "'," & CompanyID
		strSQL= "sp_Appointments"
		
		'Response.Write strSQL
		
		Dim aHourCnt
		aHourCnt = Array(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
		
		Dim SpanArr (24)
		
		
		dim cmd
		
		'Initialize Array
		'For intCnt = LBound(aHourCnt) to UBound(aHourCnt)
		'	aHourCnt(intCnt) = 0
		'Next 
		
		
		'Response.End 
		
		
		set cmd = server.CreateObject("ADODB.command")
		cmd.ActiveConnection = cnSQL
		cmd.CommandText = strSQL
		cmd.CommandType = adCmdStoredProc
		
		cmd.Parameters.Append cmd.CreateParameter("@startDate",adDBTimeStamp,adParamInput,,TaskPadTargetDate)
		cmd.Parameters.Append cmd.CreateParameter("@endDate",adDBTimeStamp,adParamInput,,TaskPadTargetDate) 'DateAdd("D",1,TaskPadTargetDate))
		cmd.Parameters.Append cmd.CreateParameter("@CompanyID",adInteger,adParamInput,,CompanyID)
		
		strFromDate = DateAdd("D",-intDaysMinus,TaskPadTargetDate)
		strToDate = DateAdd("D",1+intDaysPlus,TaskPadTargetDate)
		
		'Response.Write strToDate
		
		cmd.Parameters.Append cmd.CreateParameter("@CalFrom",adDBTimeStamp,adParamInput,,strFromDate)
		cmd.Parameters.Append cmd.CreateParameter("@CalTo",adDBTimeStamp ,adParamInput,,strToDate)
		' & TaskPadTargetDate & "', '" & DateAdd("D",1,TaskPadTargetDate) & "'," & CompanyID
		'Response.Write remote.Session("DefaultCalView")
		dcv = remote.Session("DefaultCalView")
		did = remote.Session("DefaultDepartmentID")
		
		cmd.Parameters.Append cmd.CreateParameter("@CalViewID",adInteger,adParamInput,,dcv)
		cmd.Parameters.Append cmd.CreateParameter("@DepartmentID",adInteger,adParamInput,,did)
		cmd.Parameters.Append cmd.CreateParameter("@UserID",adInteger,adParamInput,,remote.session("FloatingUser_UserID"))
		
		'Response.Write "@DaysPlus: " & intDaysPlus
		'Response.End
		
		'set rs = server.CreateObject("ADODB.recordset")
		set rsApp = cmd.Execute
		
		'set rsApp = rs
		'rsApp.Open strSQL, cnSQL , adOpenStatic, adLockReadOnly, adCmdStoredProc
		
		'Response.Write rsApp.recordcount & "<br>"
		'Response.Write cnSQL.ConnectionString
		'Response.End 
	

		' Getting permissions to the Department
		
		strPerm = "select * from tlnkUserDepartment where UserID=" & remote.session("FloatingUser_UserID") & " and CompanyID=" & CompanyID & " and DepartmentID=" & did
		
		set rsPerm = CreateObject("ADODB.Recordset")
		set rsPerm = cnSQL.Execute (strPerm)
		
		
		
		if tmpSuperUser = 1 or did = 0 Then
			CanView = True
			CanAdd = True
			CanEdit = True
			CanClose = True
		Else
			
				If not rsPerm.eof Then
					CanView = Not rsPerm("ViewTasks").Value 
					CanAdd = Not rsPerm("AddTasks").Value 
					CanEdit = Not rsPerm("EditTasks").Value 
					CanClose = Not rsPerm("CloseTasks").Value 
				End If
		
		End If
		
		'Response.Write canView
		'Response.End 
		
		
		
		
		
		'Couldn't get recordcount for this rs so I will set flag
		
		' Changed to reflect the new permission Scheme 10/15/01 IR
		bHasRecs = False
		If Not rsApp.EOF and remote.Session("VCT")  Then bHasRecs = True
			
		'set rsTemp = server.CreateObject("ADODB.recordset")
		'set rsTemp = rsApp.Clone
		
		' calculate for sizing
		tcnt = 0
		strIDArray = "['" 
		Do While Not rsApp.EOF
		
			if isnull(rsApp.Fields("Alarm").Value) then
				Alarm = "0"
			else
				If rsApp.Fields("Alarm").Value Then
					Alarm = "1"
				Else
					Alarm = "0"
				End If
			end if
		
			tcnt = tcnt + 1
			strIDArray = strIDArray & rsApp("AppointmentID") & "|" & Alarm  & "','" 
			intHour = Hour(rsApp.Fields("ApptStartDate"))
			aHourCnt(intHour) =  aHourCnt(intHour) + 1
			rsApp.MoveNext
		Loop
		
		strIDArray = Left(strIDArray,len(strIDArray)-2) & "]"
		'Response.Write strIDArray
		'Response.End
		
		b1cnt = 0
		bcnt = 0
		for i = 0 to 23
			if aHourCnt(i) < 2 then
				b1cnt = b1cnt + 1
			else
				mcnt = mcnt + (aHourCnt(i)-1)
			end if
		next

		' Response.Write "<br>" & strIDArray & "<br>"
		
		'Response.Write "Recordcount: " & rsApp.RecordCount & "<BR>"
		
		If bHasRecs Then
			rsApp.MoveFirst
		End If
		
		If IsDate(Request.QueryString("TargetDate")) Then
			dtToUse = Request.QueryString("TargetDate")
		Else
			dtToUse = TaskPadTargetDate
		End If

		strCalColor = vbBlack
		
		cnt = 0
		aAppts = Array()
		redim aAppts(500)
				
		set rsAppointmentNotes = server.CreateObject("ADODB.recordset")
		set rsAppointmentNotes = cnSQL.Execute("sp_TaskNotesFieldsFilter " & CompanyID & ", '" & dtToUse & "'")
				
        For intCnt = -2 to 23
			if intCnt < 12 then
				if dcv = 0 then
					strHourBGColor = "#F9D568"
				else
					strHourBGColor = "#99D8F0"
				end if
			else
				if dcv = 0 then
					strHourBGColor = "#FCE6A3"
				else
					strHourBGColor = "#BCE6F5"
				end if
			end if
				
			' Calculate Task Height based on all kinds of stuff...
			DefaultTaskHeight = ((rsah-90)/26)
			'intTaskHeight = DefaultTaskHeight
			intTaskHeight = Round(DefaultTaskHeight-(((15*(mcnt+1)))/26),0)
			strClass = "hours"
			strAlign = ""
			select case intCnt
				case -2
					strAlign = "center"
					'intTaskHeight = Round(DefaultTaskHeight-(((15*mcnt))/26),0)
					strTimeCol =  "<img title=""Reminders"" src=images/remind.gif>"
					strNoTime = "True"
					hbgcolor = "#FDF0C8"
				case -1
					strTimeCol =  "Notes"
					'intTaskHeight = Round(DefaultTaskHeight-(((15*mcnt))/26),0)
					strNoTime = "True"
					hbgcolor = "#FCE6A3"
				case else
					strTimeCol =  Trim(FormatDateTimeJA(FormatDateTime(dtStartTime,4)))
					'if aHourCnt(intCnt) < 2
					'	'intTaskHeight = Round(DefaultTaskHeight-(((15*mcnt))/26),0)
					'end if
					strNoTime = "False"
					hbgcolor = "Silver"
			end select
			' done with calc
	        
	        Response.Write "<tr height=""" & intTaskHeight & """ valign=middle>" & vbcrlf
			Response.Write "<td style=""WIDTH: 35px"" align=""right"" noWrap bgColor=""" & hbgcolor & """ bordercolor=""Black"">" & vbcrlf
			if intCnt > -2 then
				strString = "jaDivClick(" & """Appointment.asp?NoTime=" & strNoTime & "&TargetDate=" & dtToUse & "&ID=0&Hour=" & FormatDateTime(dtStartTime,4) & "&AppointmentID=0"")"
			else
				strString = ""
			end if
			
			If (CanAdd) and intCNT > -2 Then 
				strHand = "style=cursor:hand;padding-right:2px onclick=" & strString
			else
				strHand = ""
			end if
			
			Response.Write "<div align=" & strAlign & " name=""divTime" & intCnt & """ id=""divTime" & intCnt & """ " & strHand & ">" & vbcrlf
			Response.Write "<p ilia valign=middle class=" & strClass & ">" & strTimeCol & "</p>" & vbcrlf
			Response.Write "</div>" & vbcrlf
			Response.Write "</td>" & vbcrlf
			Response.Write "<td valign=top width=" & intWidthFactor+19 & "px noWrap bgColor=""" & strHourBGColor & """ bordercolor=""Black"">" & vbcrlf
			Response.Write "<table border=1 cellpadding=0 cellspacing=0 width=""100%"">" & vbcrlf

				'intLoopTrap = 1
				PrevApptID = -9 
				strTaskWidth = intWidthFactor '"330"
				strStatusImage = ""
				
				If intCnt > 0 Then ' This is the logic to show the spanned tasks
									' Can't do it any later cause if there are no records besides the original spoanned one the while loop will end.
				'Response.Write Len(SpanArr (intCnt+1))
					If Len(SpanArr (intCnt)) > 1 Then
						Response.Write SpanArr (intCnt)
					End If
				End If
				
				
				' Response.Write 1
				Do While Not rsApp.EOF
				
				
					
					
					strTaskNote = ""
					intTNCount = 0
					raid = rsApp.Fields("AppointmentID")


					Response.Write "<TR>"
					''Response.Write "Inside While loop PrevApptID: " & PrevApptID & "<BR>"
					' Skip duplicates!
					if (PrevApptID = raid) then
						rsApp.MoveNext
						
						'do while ( (Not rsApp.EOF) and (PrevApptID = raid) )
						do while Not rsApp.EOF
							if (PrevApptID = raid) then
								''Response.Write "Skipping MORE......." & "<BR>"
								rsApp.MoveNext 
							else
								exit do ' exit inner do
							end if
						loop

						if rsApp.EOF then
							''Response.Write "ExiTING Do While loop inside skipping condition..." & "<BR>"
							exit do ' exit outer do 
						end if 
					end if
					

					'Figure out colors
					'If rsApp.Fields("Closed") <> 0 Then
					strStyle = ""
					strTaskWidth = intWidthFactor-15 '"335"
					'strStatusStyle = ""
					factor = 0
					select case rsApp.Fields("Status").Value
						case "c"
							'Closed Task
							strColor="closedtask"
							strBackColor="lightgreen"
							strStatusImage = "<img title=""Closed"" border=0 src=""images/closed.gif"">"
						case "p"
							'Pending Task
							strColor="pendingtask"
							strBackColor= "#FFB353" '"#FFFFA8" '"#FFD2D2" '"#FF6666" '"red" '"#C10000" '
							strStatusImage = "<img title=""Pending"" border=0 src=""images/pending.gif"">"
						case "o"
							'Open Task
							strColor="opentask"
							strBackColor= "#FFB3B3" '"#FFD2D2" '"#FF6666" '"red" '"#C10000" '
							strStatusImage = "<img title=""Open"" border=0 src=""images/open.gif"">"
						case "x"
							'Canceled Task
							'strStyle = ";background-image:url(images/canceledBG.gif)"
							strColor="opentask"
							strBackColor= "lightblue" '"#FFD2D2" '"#FF6666" '"red" '"#C10000" '
							strStatusImage = "<img title=""Canceled"" border=0 src=""images/canceled.gif"">"

						case "r"
							'ReConfirmed Task
							'strStyle = ";background-image:url(images/canceledBG.gif)"
							strColor="opentask"
							strBackColor= "#82FFE0" '"#FFD2D2" '"#FF6666" '"red" '"#C10000" '
							strStatusImage = "<img title=""Re-Confirmed"" border=0 src=""images/Recon.gif"">"
							'strStatusImage = "Re-Confirmed"
							'strStatusStyle = ""
						case "n"
							'Not Available Task
							'strStyle = ";background-image:url(images/canceledBG.gif)"
							strColor="pendingtask"
							strBackColor= "lavender" '"#FFFFA8" '"#FFD2D2" '"#FF6666" '"red" '"#C10000" '
							strStatusImage = "<img title=""Not Available"" border=0 src=""images/NA.gif"">"
							factor = 20
							
						case "w"
							'Not Available Task
							'strStyle = ";background-image:url(images/canceledBG.gif)"
							strColor="pendingtask"
							strBackColor= "#cc99ff" '"#FFFFA8" '"#FFD2D2" '"#FF6666" '"red" '"#C10000" '
							strStatusImage = "<img title=""Wait List"" border=0 src=""images/WL.gif"">"
							factor = 20
							
					end select					

					if isnull(rsApp.Fields("Alarm").Value) then
						booAlarm = false
					else
						booAlarm = rsApp.Fields("Alarm").Value
					end if
					
					if rsApp.Fields("Note") <> 0 and not booAlarm then
						strColor="closedtask"
						select case rsApp.Fields("Status").Value
							case "c"
								'Closed Task
								strBackColor="#CAF7CA"
							case "p"
								strBackColor="#FFB353"
							case "n"
								strBackColor="lavender"
							case "o"
								strBackColor="white"
							case "w"
								strBackColor="#cc99ff"
						end select
						strStatusImage = "<img title=""Note"" border=0 src=""images/note.gif"">"
					end if 

					rcwidth = 0					
					If IsNull(rsApp.Fields("TaskID")) Then
						strRecImage = ""
						strTaskWidth = intWidthFactor+1 '"351"
					Else
						rcwidth = 15
						If rsApp.Fields("RecException") Then
							strRecImage = "<img title=""Reccurence (Task Edited)"" border=0 src=""images/Recur_excp.gif"">"
						Else
							strRecImage = "<img title=""Reccurence"" border=0 src=""images/Recur_norm.gif"">"
						End IF
					End If
					
					strReminderImage = ""
					strReminderBG = ""
					if booAlarm then
						'Reminder
						if rsApp.Fields("Status").Value = "c" or rsApp.Fields("Status").Value = "x" then
							strReminderImage = "<img title=""Reminder"" border=0 src=""images/remindoff.gif"">"
							'strBackColor= "#CCFF99"
							'strReminderBG = " background=images/ReminderBG_Closed_Green.jpg "
						else
							strReminderImage = "<img title=""Reminder"" border=0 src=""images/remindani.gif"">"
							'strBackColor= "#FFB3B3"
							'strReminderBG = " background=images/ReminderBG_Open.jpg "
						end if
						strColor="pendingtask"
						factor = factor + 20
					end if
					
					strItineraryImage = ""
					if rsApp.Fields("ItineraryID").Value <> 0 then
						strItineraryImage = "<img title=""Itinerary"" border=0 src=""images/itinerary.gif"">"
						factor = factor + 20
					end if

					if rsApp.Fields("ReminderID").Value <> 0 then
						strReminderImage = "<img title=""Reminder"" border=0 src=""images/remindoff.gif"">"
						factor = factor + 20
					end if

					If not IsNull(rsApp.Fields("OTLogID")) Then
						If rsApp.Fields("OTLogID") > 0 Then
							strTaskWidth = intWidthFactor-15-rcwidth '"335"
							strRecImage = "<img title=""OpenTable Task"" height=12 border=0 src=""images/ot.gif"">"
						End IF
					End If
					
					If not IsNull(rsApp.Fields("SSLogID")) Then
						If rsApp.Fields("SSLogID") > 0 Then
							strTaskWidth = intWidthFactor-17-rcwidth '"335"
							strRecImage = "<img title=""SuperShuttle Task"" height=12 border=0 src=""images/ss.gif"">"
						End IF
					End If
					
					
					if isnull(rsApp.Fields("DateAdded").Value) then
						strDateAdded = "01/01/03"
					else
						strDateAdded = formatdatetime(rsApp.Fields("DateAdded").Value,vbShortDate)
					end if

					if (rsApp.Fields("Rollover").Value and strDateAdded <> formatdatetime(rsApp.Fields("ApptStartDate").Value,vbShortDate)) then 'or (rsApp.Fields("Rollover").Value and (rsApp.Fields("Status").Value = "p" or rsApp.Fields("Status").Value = "o")) then
						strTaskWidth = intWidthFactor-33-rcwidth
						strRecImage = "<img title=""Rollover Task"" height=12 border=0 src=images/roll.gif>" & strRecImage
					end if

					strTaskWidth = strTaskWidth - factor

					''''' End of icons
					
					appStart = Hour(rsApp.Fields("ApptStartDate"))
					appEnd   = Hour(rsApp.Fields("ApptEndDate"))
					
					
					'intCnt > (appStart-1) and intCnt < (appEnd + 2) and intCnt > -1
					
					If (intCnt = appEnd or intCnt=AppStart) or (intCnt = -1 and rsApp.Fields("NoTime").Value) or (intCnt = -2 and booAlarm and rsApp.Fields("NoTime").Value) Then
					
							'if rsApp.Fields("Note").Value = true then
							if rsApp.Fields("Note") <> 0 then
								strTitleText = "Note"
								'strTaskText = FormatAppointment(true,FormatDateTime(rsApp.Fields("ApptStartDate"),3), rsApp.Fields("Room"), rsApp.Fields("Salutation"), rsApp.Fields("GuestName"),rsApp.Fields("GuestFirstName"), rsApp.Fields("GuestLastName"),rsApp.Fields("ActionType"), rsApp.Fields("Action"), rsApp.Fields("LocationText"), rsApp.Fields("ApptText").Value) 'rsApp.Fields("ApptText"))
								strTaskText = FormatAppointment(true,FormatDateTime(rsApp.Fields("ApptStartDate"),3),FormatDateTime(rsApp.Fields("ApptEndDate"),3), rsApp.Fields("Room"), rsApp.Fields("Salutation"), rsApp.Fields("GuestName"),rsApp.Fields("GuestFirstName"), rsApp.Fields("GuestLastName"),rsApp.Fields("ActionType"), rsApp.Fields("Action"), rsApp.Fields("DisplayVendor"), rsApp.Fields("ApptText").Value, rsApp.Fields("Status").Value, rsApp.Fields("NoTime").Value,"",booAlarm,rsApp.Fields("ReminderNote").Value) 'rsApp.Fields("ApptText"))
								strTaskTextWOLocation = FormatAppointment(true,FormatDateTime(rsApp.Fields("ApptStartDate"),3),FormatDateTime(rsApp.Fields("ApptEndDate"),3), rsApp.Fields("Room"), rsApp.Fields("Salutation"), rsApp.Fields("GuestName"),rsApp.Fields("GuestFirstName"), rsApp.Fields("GuestLastName"),rsApp.Fields("ActionType"), rsApp.Fields("Action"), rsApp.Fields("DisplayVendor"), rsApp.Fields("ApptText").Value, rsApp.Fields("Status").Value, rsApp.Fields("NoTime").Value, "",booAlarm,rsApp.Fields("ReminderNote").Value) 'rsApp.Fields("ApptText"))
								if len(trim(strTaskText)) = 0 then
									strTaskText = "&nbsp;"
									strTaskTextWOLocation = "&nbsp;"
								end if
							else
								' Show notes that have "Calendar" set  to 1
								
								sql = "select an.Data, display from tlnkAppointmentNotes an join tlnkActionNotesFields anf on anf.ActionID=an.ActionID and anf.NotesFieldsID=an.NotesFieldID "
								sql = sql & " where anf.Calendar = 1 and an.AppointmentID=" & rsApp("AppointmentID") & " and CompanyID=" & CompanyID & " order by anf.ListIndex asc "
								
								strNotesDetails = ""
								set rsNotes = cnSQL.Execute (sql)
								
								do while not rsNotes.eof
									if trim(rsNotes(0)) = "" then
										v = null
									else
										if trim(rsNotes(1))="" Then
											v = rsNotes(0)
										Else
											if isnull(rsNotes(1)) then
												srsNotes = ""
											else
												srsNotes = rsNotes(1) & " : "
											end if
											v = srsNotes & rsNotes(0)
										End If
									end if
									strNotesDetails = strNotesDetails & (v + ", ")
									rsNotes.movenext 
								Loop
								
								set rsNotes = Nothing

								if not isnull(strNotesDetails) and len(strNotesDetails) > 2 then
									nnn = left(strNotesDetails,len(strNotesDetails)-2)
									strNotesDetails = nnn
								end if
							                    					
								'strTaskText = FormatAppointment(true,FormatDateTime(rsApp.Fields("ApptStartDate"),3), rsApp.Fields("Room"), rsApp.Fields("Salutation"), rsApp.Fields("GuestName"),rsApp.Fields("GuestFirstName"), rsApp.Fields("GuestLastName"),rsApp.Fields("ActionType"), rsApp.Fields("Action"), rsApp.Fields("LocationText"), "") 'rsApp.Fields("ApptText"))
								strTaskText = FormatAppointment(true,FormatDateTime(rsApp.Fields("ApptStartDate"),3), FormatDateTime(rsApp.Fields("ApptEndDate"),3), rsApp.Fields("Room"), rsApp.Fields("Salutation"), rsApp.Fields("GuestName"),rsApp.Fields("GuestFirstName"), rsApp.Fields("GuestLastName"),rsApp.Fields("ActionType"), rsApp.Fields("Action"), rsApp.Fields("DisplayVendor"), rsApp.Fields("ApptText").Value, rsApp.Fields("Status").Value, rsApp.Fields("NoTime").Value, strNotesDetails, booAlarm,rsApp.Fields("ReminderNote").Value)
								'strTaskTextWOLocation = FormatAppointment(true,FormatDateTime(rsApp.Fields("ApptStartDate"),3), rsApp.Fields("Room"), rsApp.Fields("Salutation"), rsApp.Fields("GuestName"),rsApp.Fields("GuestFirstName"), rsApp.Fields("GuestLastName"),rsApp.Fields("ActionType"), rsApp.Fields("Action"), rsApp.Fields("LocationText"), "" ) 'rsApp.Fields("ApptText"))
								strTaskTextWOLocation = FormatAppointment(true,FormatDateTime(rsApp.Fields("ApptStartDate"),3), FormatDateTime(rsApp.Fields("ApptEndDate"),3), rsApp.Fields("Room"), rsApp.Fields("Salutation"), rsApp.Fields("GuestName"),rsApp.Fields("GuestFirstName"), rsApp.Fields("GuestLastName"),rsApp.Fields("ActionType"), rsApp.Fields("Action"), rsApp.Fields("DisplayVendor"), rsApp.Fields("ApptText").Value, rsApp.Fields("Status").Value, rsApp.Fields("NoTime").Value, strNotesDetails, booAlarm,rsApp.Fields("ReminderNote").Value)
								strTitleText = "Task"
							end if

							if rsApp.Fields("Alarm").Value then
								strTitleText = "Reminder"
							end if
							
							
							strStr = "Appointment.asp?TargetDate=" & dtToUse & "&ID=" & raid 
							
							If Not IsNull(rsApp.Fields("TaskID").Value) Then
								strStr = strStr & "&RecID=" & rsApp.Fields("TaskID").Value 
							End If
							
							'strStr2 = "javascript:jaDivClick('" & strStr & "')"
							if booAlarm then
								strStr2 = "javascript:showActivate('" & raid & "|1')"
							else
								strStr2 = "javascript:showActivate('" & raid & "|0')"
							end if
							
							timespan = (appEnd - AppStart)
							
							if Minute(rsApp.Fields("ApptEndDate")) = 0 Then
								timespan = timespan - 1
							End If
							
							if timespan > 0 Then
								strSpan = "taskspan" & raid & "t0"
								'strSpan = "alert(1)"
								'taskSPanStart = "<div id=""taskspan" & raid & "t0"">"
								'taskSapnEnd = "</div>"
							Else
								strSpan = ""
								'taskSPanStart = ""
								'taskSapnEnd = ""
							End If
							
							strStr3 = "javascript:show('" & raid & "', '" & rsApp.Fields("span").Value & "')"' & strSpan
							strStr4 = "javascript:hide('" & raid & "')"
														
							strEvents = "onclick=""" & strStr2 & """" & " onmouseover=""" & strStr3 & """" & " onmouseout=""" & strStr4 & """" & " title=""Click to view/edit " & strTitleText & """"
							
							strResult = "<td id=" & strSpan & " " & strReminderBG & " style=border-right-style:none" & strStyle & " align=left bgcolor=""" & strBackColor & """>"
							strResult = strResult & "<span style=""cursor: hand"">"
							strResult = strResult & "<div id=divEditTask" & raid & " onclick=javascript:jaDivClick('" & strStr & "')></div>"
							
							if rsApp.Fields ("CanViewTasks") > 0 or CInt(rsApp("CreateUserID")) = Cint(remote.session("FloatingUser_UserID")) Then
							
									strResult = strResult & "<table style=""padding:0px"" cellpadding=0 cellspacing=0 border=0 width=""100%""><tr><td>"
									'strResult = strResult & taskSpanStart
									strResult = strResult &	"<div name=""divTask" & raid & """ id=""divTask" & raid & """ nowrap style=""padding: 0px; z-index: 10; OVERFLOW: hidden; width:" & strTaskWidth & "px""" &  strEvents & ">"
									strResult = strResult & "<font face=""Tahoma"" size=""2"">"
									strResult = strResult & "<input type=""hidden"" id=""txtTaskText" & raid & """ name=""txtTaskText" & raid & """ value=""" & replace(strTaskText,"'","''")& """>"
									strResult = strResult & "<p style=""padding: 0px"" class=""" &  strColor & """>" & strTaskTextWOLocation & "</p>"
									strResult = strResult & "</font></div>"
									strResult = strResult & "</td><td width=80px align=""right"" style=""padding-right:1px"">" & "<div " & strEvents & ">" & strItineraryImage & strReminderImage & strRecImage & strStatusImage & "</div></td></tr>"
									'strResult = strResult & taskSpanEnd
									strResult = strResult & "</table>"
					
							End If 
							
							strResult = strResult & "</span></td></tr>"
							
							If (timespan > 0 and rsApp.Fields ("Span")) Then
							   	for tmsp = intCnt + 1 to intCnt + timespan
							   		spantmp = Replace(strResult,"divEditTask" & raid,"divEditTask" & raid & "a")
							   		spantmp = Replace(spantmp,"taskspan" & raid & "t0","taskspan" & raid & "t" & tmsp)
									SpanArr (tmsp) = SpanArr (tmsp) & spantmp
							   	next
							End If
							
							 Response.Write strResult
							'Response.End 
							

							
							%>
								<!--td id="tdExpandTask<%'=raid%>" width="16px" style="font-family: tahoma; font-size: 11px; border-left:none;" bgcolor="<%'=strBackColor%>"><span title="Click for details" id="spanExpandTask" style="cursor: hand" onclick="javascript:showTextActivate('<%'=raid%>')"><img src="images/Magnifying_Glass.gif" WIDTH="16" HEIGHT="16"></span></td>
								</tr-->
								<%
						
						PrevApptID = raid 
						rsApp.MoveNext
					Else
						Exit Do
					End If

				Loop
				 %>
				</tr>
				</table>
				</font>
				</td>
			</tr>
			<%
			if strNoTime = "False" then
				dtStartTime = DateAdd("n",60,dtStartTime)
			end if
        Next
		rsAppointmentNotes.Close
		set rsAppointmentNotes = nothing

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
		
sub emailGCN()
	'dim rstemp, cdoobj, sfe, userfullname, guestname
	'set rstemp = server.CreateObject("adodb.recordset")
	'rstemp.Open "select companyname,city,postalcode,phone from tblCompany where companyid = " & remote.Session("companyid"),cn,adOpenForwardOnly,adLockReadOnly
	'hotelinfo = rstemp("companyname") & ", " & rstemp("city") & "  " & rstemp("postalcode") & "  -  " & PhoneMask(rstemp("phone"))
	'rstemp.Close
	'set rstemp = nothing
	
	On Error Resume Next
	
	set cdoobj = Server.CreateObject("CDONTS.NewMail")

	
	cdoobj.To = "iliar@goconcierge.net"
	'cdoobj.To = "simon@goconcierge.net"
	
	
	sfe = "gcnerros@goconcierge.net"

	
	cdoobj.BodyFormat = 0
	cdoobj.MailFormat = 0
	cdoobj.Importance = 1
	cdoobj.From = sfe
	cdoobj.Subject = "A switch has occured"
    
	

	'strBody = "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">" & vbCrLf
	'strBody = strBody & "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1""><title>Not-in-list notice</title></head>"
	'strBody = strBody & "<body style=""font-family:tahoma;font-size:11px"" bgcolor=white>"
	'strBody = strBody & "<table cellspacing=2 style=border-style:solid;border-width:2px;border-color:black>"
	'strBody = strBody & "<tr><td colspan=2 bgcolor=#FAD667 style=font-size:11px;border-style:outset;border-width:1px>"
	'strBody = strBody & "A Company switch has occured." & vbcrlf & vbcrlf
	'strBody = strBody & "</td></tr>"
	'strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>Property:</td><td style=font-size:11px bgcolor=#FAD667>" & hotelinfo & "</td></tr>"
	'strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>User:</td><td style=font-size:11px bgcolor=#FAD667>" & userfullname & "</td></tr>"
	'strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>Vendor:</td><td style=font-size:11px bgcolor=#FAD667>" & strLocationText & "</td></tr>"
	'strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>Vendor Address:</td><td style=font-size:11px bgcolor=#FAD667>" & strAddress & "</td></tr>"
	'strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>Vendor Phone:</td><td style=font-size:11px bgcolor=#FAD667>" & PhoneMask(strPhone) & "</td></tr>"

	'shref = Application("HomePage") & "/Login.asp?url=" & escape(Application("HomePage") & "/LocationSetup3.asp?lid=" & strID & "&cid=" & remote.Session("CompanyID"))

	'strBody = strBody & "<tr><td bgcolor=#FAD667 style=font-size:11px;font-weight:bold>Edit Link:</td><td style=font-size:11px bgcolor=#FAD667><A href='" & shref & "'>Edit This Location </a></td></tr>"
	'strBody = strBody & "</table>"
	strBody = strBody & "UserKey:" & Request.Cookies("UserKey")
			
	cdoobj.Body = strBody
	cdoobj.Send
	set cdoobj = nothing
end sub	
		
		
		%>
        </table>
<input type="hidden" value="OK" id="txtTaskPadLoaded">
</body>
</html>


<script id="VariousWithToolTip" language="javascript">
	var intHourDivTop = 0, intDelay;
	var popup, booToolTipOpen = false, booOverFrame = false;
	<%
	for i = lbound(aAppts) to ubound(aAppts)
		if len(aAppts(i)) > 0 then
			response.write aAppts(i) & vbcrlf
		end if
	next
	%>
	
	/*if( window.parent.frames.frameSearchCache.document.location == "about:blank" )		
		window.parent.frames.frameSearchCache.document.location.replace("BrowseLocationsMain.asp")
	//else
	//	enableMenu()
	*/
</script>

<script language="JavaScript">
	
	<%if len(strIDArray) > 2 Then%>
		parent.idarr = <%=strIDArray%>; // for implementing the prev/next functionality
	<%end if%>	

	parent.Calobj.colarr = null
	parent.Calobj.colarr = new Array()
		<%
		Dim rsBooked 
		Dim dteDate
		Dim intDaysPlusMinus
		
		dteDate = Date()
		set rsBooked = rsApp.NextRecordSet
		
		Do While Not rsBooked.EOF
		  Response.write ("parent.Calobj.colarr['" & Replace(rsBooked.Fields(0).Value,"/","") & "'] = '" & rsBooked("BackGroundColor") & "';" & vbCrLf)
		  rsBooked.MoveNext
		Loop

		rsApp.Close
		set rsApp = nothing
		rsBooked.Close
		set rsBooked = nothing
		cnSQL.Close
		set cnSQL = nothing
		%>
		
	parent.Calobj.setDate('<%=TaskPadTargetDate%>');
	//alert('<%=TaskPadTargetDate%>');
	parent.enableMenu()
	
	//delete parent.Calobj.colarr;
	
	//parent.Calobj.colarr['120701'] = 'test'
	
</script>

<script language="vbscript">
	' NEEDS TO BE AT THE VERY END OF THIS PAGE
	dim td, m, d, y, strMDY, strmonth
	
	td = cdate("<%=date%>")
	m = right("0" & month(td),2)
	d = right("0" & day(td),2)
	y = right(year(td),2)
	
	std = formatdatetime(td,1)

	strMDY = m & d & y
	
	strmonth = mid(std,instr(1,std," ")+1)
	strmonth = left(strmonth,instr(1,strmonth," ")-1)
	
	parent.document.all("t").value = strmonth & " " & day(td) & ", " & year(td) 'mid(std,instr(1,std,",")+2)
	parent.document.all("tcal").value = strMDY
</script>
