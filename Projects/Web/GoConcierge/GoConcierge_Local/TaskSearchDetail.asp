<%@ Language=VBScript %>

<%
Response.CacheControl = "No-Cache"
Response.AddHeader "Pragma", "No-cache"
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))

REC_NUM = 16

LastName = replace(Request.Form("LastName"),"'","''")
FirstName = replace(Request.Form("FirstName"),"'","''")
Status = Request.Form("Status")
Vendor = Request.Form("ddLoc")
Room = Request.Form("ddRoom")
StartDate = Request.Form("d1")
EndDate = Request.Form("d2")
SortColumn = Request.form("Sort") 
SortDir = Request.Form("SortDir")
CreateUserID = Request.Form("CreateUserID")
ClosedUserID = Request.Form("ClosedUserID")
Notes = Request.Form("Notes")
strAction = Request.Form("cboAction")
strTaskID = Request.Form("txtTaskID") 'sk 9/23/03

strWhereClause = "a.fkCompanyID=" & remote.Session("CompanyID")

If LastName<>"" Then 
	if strWhereClause <> "" then strWhereClause = strWhereClause & " and "
	
	strWhereClause = strWhereClause & " GuestLastName like '" & LastName & "%'"
end if

If FirstName<>"" Then 
	
	if strWhereClause <> "" Then strWhereClause = strWhereClause & " and "

	strWhereClause = strWhereClause & " GuestFirstName like '" & FirstName & "%'"

End If


If Status <> "" Then 
	if strWhereClause <> "" Then strWhereClause = strWhereClause & " and "
	strWhereClause = strWhereClause & " Status='" & Status & "'"
End If

If Trim(Vendor)<>"" and Vendor <> "0" Then 
	if strWhereClause <> "" Then strWhereClause = strWhereClause & " and "
	strWhereClause = strWhereClause & " l.LocationID = " & Vendor
End If

If Trim(Room)<>"" Then 
	
	if strWhereClause <> "" Then strWhereClause = strWhereClause & " and "

	strWhereClause = strWhereClause & " Room like '" & Room & "%'"

End If


If CreateUserID<>"" Then 
	
	if strWhereClause <> "" Then strWhereClause = strWhereClause & " and "

	strWhereClause = strWhereClause & " CreateUserID=" & CreateUserID

End If

If ClosedUserID<>"" Then 
	
	if strWhereClause <> "" Then strWhereClause = strWhereClause & " and "

	strWhereClause = strWhereClause & " ClosedUserID=" & ClosedUserID

End If

'Response.Write "Action: " & strAction

If strAction > 0 Then 
	
	if strWhereClause <> "" Then strWhereClause = strWhereClause & " and "

	strWhereClause = strWhereClause & " a.ActionID=" & strAction

End If

if strTaskID <> "" then
	if strWhereClause <> "" Then strWhereClause = strWhereClause & " and ("
	aTaskID = split(strTaskID,",")
	for zyx = lbound(aTaskID) to ubound(aTaskID)
		strWhereClause = strWhereClause & " a.AppointmentID = " & aTaskID(zyx) & " or "
	next
	strWhereClause = left(strWhereClause,len(strWhereClause)-4) & ")"
end if


if len(trim(EndDate)) > 0 then
	strEndDate = formatdatetime(EndDate,2) & " 23:59:59"
else
	strEndDate = ""
end if

if strWhereClause <> "" Then strWhereClause = strWhereClause & " and "

strWhereClause = strWhereClause & " apptStartDate between '" & StartDate & "' and '" & strEndDate & "'"

If SortColumn <> "" Then strOrderByClause = " order by " & SortColumn & " " & SortDir

Dim rsAction, rs, cn, SQL, SQLNOORDER, sintPeopleID, sintKidsID, sintAdultsID
prl = Request.QueryString("prl")

Set cn = CreateObject("Adodb.connection")
cn.CursorLocation = 3
cn.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")

Set rsAction = Server.CreateObject("ADODB.Recordset")
strSQL = "select NotesFieldID from tlkpNotesFields where NotesField = '# People'"
rsAction.Open strSQL, cn, adOpenDynamic, adLockReadOnly
if rsAction.EOF then
    sintPeopleID = 0
else
    sintPeopleID = rsAction("NotesFieldID")
end if
rsAction.Close
strSQL = "select NotesFieldID from tlkpNotesFields where NotesField = '# Kids'"
rsAction.Open strSQL, cn, adOpenDynamic, adLockReadOnly
if rsAction.EOF then
    sintKidsID = 0
else
    sintKidsID = rsAction("NotesFieldID")
end if
rsAction.Close
strSQL = "select NotesFieldID from tlkpNotesFields where NotesField = '# Adults'"
rsAction.Open strSQL, cn, adOpenDynamic, adLockReadOnly
if rsAction.EOF then
    sintAdultsID = 0
else
    sintAdultsID = rsAction("NotesFieldID")
end if
rsAction.Close

strSQL = "select NotesFieldID from tlkpNotesFields where NotesField = 'Confirmation #'"
rsAction.Open strSQL, cn, adOpenDynamic, adLockReadOnly
if rsAction.EOF then
    sintConfID = 0
else
    sintConfID = rsAction("NotesFieldID")
end if
rsAction.Close
set rsAction = nothing


If Request.QueryString ("p") = "" Then 

	Set rs = CreateObject("Adodb.recordset")

	SQL = "Select l.CompanyName as Vendor, a.*,convert(varchar,a.ApptStartDate,1) as DateOnly,"

	' had to add this crap for the ugly Single Task Report
	SQL = SQL & "CASE a.AllDayEvent WHEN 0 THEN 'No' ELSE 'Yes' END as AllDayEventState,"
	SQL = SQL & "CASE a.Status WHEN 'o' THEN 'Open' WHEN 'p' THEN 'Pending' WHEN 'c' THEN 'Closed' WHEN 'n' Then 'Not Available' END as ClosedState,"
	SQL = SQL & "stuff(stuff(stuff('123',3,1,ut2.UserLName),2,1,' '),1,1,ut2.UserName) AS ClosedUserFullName,"
	SQL = SQL & "stuff(stuff(stuff('123',3,1,ut3.UserLName),2,1,' '),1,1,ut3.UserName) AS EditUserFullName, "
	SQL = SQL & "stuff(stuff(stuff('123',3,1,ut1.UserLName),2,1,' '),1,1,ut1.UserName) AS CreatedUserFullName, "
	SQL = SQL & "CASE a.EMailConfirm WHEN 0 THEN 'No' ELSE 'Yes' END as EMailConfirmState, "
	SQL = SQL & "CASE a.LetterPrinted WHEN 0 THEN 'No' ELSE 'Yes' END as LetterPrintedState, "
	''''''''''''''''''''''''''''''''''''''
	SQL = SQL & "CASE a.CCType WHEN 0 THEN '' WHEN 1 THEN 'Room'WHEN 2 THEN 'Credit Card' END as ChargeTo, "
	SQL = SQL & "CASE a.ChargeType WHEN 0 THEN '' WHEN 1 THEN 'Visa' WHEN 2 THEN 'MasterCard' WHEN 3 THEN 'AMEX' WHEN 4 THEN 'Discover' WHEN 5 THEN 'Diner''s Club' END as Card, "
	''''''''''''''''''''''''''''''''''''''

	SQL = SQL & "dbo.NumberOnly(an.Data)+dbo.NumberOnly(ank.Data)+dbo.NumberOnly(ana.Data) as tnPeople, "
	SQL = SQL & "LTrim(RTrim(anconf.Data)) as ConfId, "

	'SQL = SQL & "an.Data as tnPeople, "
	'SQL = SQL & "ana.Data as tnAdults, "
	'SQL = SQL & "ank.Data as tnKids, "
	
	SQL = SQL & " convert(datetime,right(convert(varchar,a.ApptStartDate),7)) as TimeOnly,"
	SQL = SQL & " ut1.UserName as CreatedBy,ut2.UserName as ClosedBy,ut3.UserName as EditedBy," 
	SQL = SQL & " l.FaxNumber as LocFax, a.locAddress, a.NoTime, l.Phone, "
	SQL = SQL & " act.*,ac.* " 
	
	remote.Session("SQLSelectClause") = SQL
	
	SQL1 = SQL
	SQL = ""
	
	SQL = SQL & " from vw_Appointment a left join tblUser ut1 on a.CreateUserID = ut1.UserID left join tblUser ut2 on ut2.UserID=a.ClosedUserID left join tblUser ut3 on ut3.UserID=a.EditUserID"
	SQL = SQL & " left join tlkpActionType act on act.ActionTypeID = a.ActionTypeID left join tlkpAction ac on ac.ActionID = a.ActionID"
	SQL = SQL & " left outer join tblLocation l on a.LocationID = l.locationid"
	SQL = SQL & " left outer join tlnkAppointmentNotes an on a.AppointmentID = an.AppointmentID and an.NotesFieldID = " & sintPeopleID
	SQL = SQL & " left outer join tlnkAppointmentNotes ana on a.AppointmentID = ana.AppointmentID and ana.NotesFieldID = " & sintAdultsID
	SQL = SQL & " left outer join tlnkAppointmentNotes ank on a.AppointmentID = ank.AppointmentID and ank.NotesFieldID = " & sintKidsID
	SQL = SQL & " left outer join tlnkAppointmentNotes anconf on a.AppointmentID = anconf.AppointmentID and anconf.NotesFieldID = " & sintConfID
	
	remote.Session("SQLFromClause") = SQL
	
	SQL1 = SQL1 & SQL
	SQL = ""
	

	SQL = SQL & " where "
	SQL = SQL & strWhereClause & " and IsNull(a.RecShow,0)<>1"
	
	remote.Session("SQLWhereClause") = SQL
	
	SQL1 = SQL1 & SQL
	SQL = SQL1
	
	SQL = SQL & strOrderByClause 	
'	Response.Write SQL
	rs.PageSize = REC_NUM
	rs.CacheSize = 1000
	rs.Open  SQL, cn,adOpenForwardOnly,adLockReadOnly
	'Response.write rs.PageSize  & " " & rs.PageCount & rs.AbsolutePage  
	'Response.End 
	
	'tmp = rs.GetString()
	Set Session("TaskRS") = rs
Else

'Response.Write "Cached.."
	Set rs = Session("TaskRS")
	
	if Request.QueryString("p") = "p" Then
		
		if rs.EOF then
			rs.MoveLast
			rs.Move -(REC_NUM+prl)
		else
			j = 0
			do while not rs.BOF and j < 2*REC_NUM
				rs.MovePrevious
				j = j + 1
			loop	
		end if		
	End If
	If rs.BOF Then rs.MoveFirst
	
End If

SQLNOORDER = rs.Source
if instr(1,SQLNOORDER,"order by") > 0 then
	SQLNOORDER = mid(SQLNOORDER,1,instr(1,SQLNOORDER,"order by")-1)
end if
remote.Session("SQLNOORDER") = SQLNOORDER


%>
<script LANGUAGE="vbscript" RUNAT="Server">

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

</script>

<STYLE>
A:active { font-family:Tahoma;font-size:11px;font-weight:bold; color:blue }
A:link {  font-family:Tahoma;font-size:11px;font-weight:bold; color:blue }
A:hover {  font-family:Tahoma;font-size:11px;font-weight:bold; color:blue }
A:visited {  font-family:Tahoma;font-size:11px;font-weight:bold; color:blue }
.bord {border-bottom-width:1px;border-bottom-style:solid;border-bottom-color:#D8BFD8;border-right-width:1px;border-right-style:solid;border-right-color:#D8BFD8 }
.but {font-family:Tahoma;font-size:11px; }
.headclass {
	border-style: outset;
	border-width: 1px;
	border-color: black;
	font-family: Tahoma; 
	font-size: 11px; }

.normclass  { font-family : Tahoma; font-size:11px; border-color: black; border-style:1px;cursor:hand }
.exp {cursor:hand}

.MainTable		{ font-family:Tahoma; font-size: 11px; }
.RaisedRow		{ border-style: outset; border-width: 1px; background-color: Khaki; cursor: hand }
.Row			{ border-style: solid; border-width: 1px; background-color: PaleGoldenrod; cursor: default }
.RowDown		{ border-style: inset; border-width: 1px; background-color: Yellow; cursor: hand }

.navUp				{ border-style: none; border-width: 2px; background-color: none; padding-top: 0px; }
.navDown			{ border-style: none; border-width: 2px; background-color: none; padding-top: 0px; }
.navNone			{ border-style: none; border-width: 2px; background-color: White; padding-top: 0px; }
.pendingtask { color: #000000; font-size: 11px; font-family: Tahoma; text-decoration: none }

</style>

<HTML>
<HEAD>
<script src=CheckIfTaskExists.asp language=javascript></script>
<script language=javascript>
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


	function showText(strField)
	{
		var fParam = 'showTextActivate("'+strField+'",'+window.event.screenX+','+window.event.screenY+')';
		intDelay = window.setInterval(fParam,1000,"javascript");
	}
		
	function showTextActivate(strField) 
	{
		if(window.event.srcElement.type != 'checkbox')
		{
			parent.document.frames.frameToolTip.AppID = strField
			document.all("frameSTA").src = "TaskPadToolTip.asp?id=" + strField;
		}
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

	  code = eval(unescape(window.frames("frameSTA").document.body.innerHTML));
	  if(code == "EOF")
	  {
		alert ("The task you selected has been deleted by another user.  The list will now refresh.");
		parent.document.all("submit").click();
	  }
	  else
	  {
		  
		  var el = parent.document.all.tooltip ;
		  var frel = parent.document.frames.frameToolTip.document.all.tbdyToolTip
		  var frid = parent.document.frames.frameToolTip.document.all.txtAppointmentID
		  var booLastWasNote = false;
		  
		  
		  el.style.pixelTop = (screen.availHeight-400)/2;
		  el.style.pixelLeft = (screen.availWidth-400)/2;

		  var j = frel.rows.length;
 
		  for(i=0;i<j;i++)
			frel.deleteRow(0);
			
	    var a = aTT;		
		  
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
  
	if(strColor=="lightgreen" || strColor=="white")
		oRow.style.backgroundColor = "#C2F5C2";
	else
		oRow.style.backgroundColor = strColor;  	

  
		  
		  
		  for(i=4; i<a.length - 1; i++)
		  {
			var b = Array();
			
			b = a[i].split("|");
			b[1] = b[1].replace(/\&lt;<td>\&gt;/gi,"<<td>>").replace(/\&lt;<sq>\&gt;/gi,"<<sq>>")

			/* if(i==2)
				parent.document.frames.frameToolTip.CurDate = b[1]; */

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
				oCell.innerHTML = b[1].replace(/<<sq>>/g,"'").replace(/<<td>>/g,"");
				oCell.vAlign = "top"
				}
		  }
	
		  el.style.visibility = "visible"

		  booToolTipOpen = true;
		}
	}
</script>

</HEAD>
<BODY topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0>
<input id="txtSQL" type="hidden" value="<%=SQLNOORDER%>">
<iframe onload=sta() src=LoadingAppointment.asp id=frameSTA style=display:none;visibility:hidden></iframe>

<table cellpadding=2 cellspacing=0>

<%
strDetailHeight = "14"

Response.Write "<tr style=""height:" & strDetailHeight & "px;"" bgcolor=""menu"">"
Response.Write "<td class=""but"" align=right style=""height:" & strDetailHeight & "px;width:35px;border-style:outset;border-width:1px"">ID</td>"
Response.Write "<td class=""but"" align=right style=""height:" & strDetailHeight & "px;width:45px;border-style:outset;border-width:1px"">Date</td>"
Response.Write "<td class=""but"" style=""height:" & strDetailHeight & "px;width:52px;border-style:outset;border-width:1px""  align=right>Time&nbsp;</td>"
Response.Write "<td class=""but"" style=""height:" & strDetailHeight & "px;width:20px;border-style:outset;border-width:1px""  align=left>Room</td>"
Response.Write "<td class=""but"" style=""height:" & strDetailHeight & "px;width:16px;border-style:outset;border-width:1px""  align=left>Sal</td>"
Response.Write "<td class=""but"" style=""height:" & strDetailHeight & "px;width:45px;border-style:outset;border-width:1px""  align=left>First</td>"
Response.Write "<td class=""but"" style=""height:" & strDetailHeight & "px;width:65px;border-style:outset;border-width:1px""  align=left>Last</td>"
Response.Write "<td class=""but"" style=""height:" & strDetailHeight & "px;width:74px;border-style:outset;border-width:1px""  align=left>Task</td>"
Response.Write "<td class=""but"" style=""height:" & strDetailHeight & "px;width:120px;border-style:outset;border-width:1px"" align=left>Vendor</td>"
Response.Write "<td class=""but"" style=""height:" & strDetailHeight & "px;width:20px;border-style:outset;border-width:1px""  align=right>Amt</td>"
Response.Write "<td class=""but"" style=""height:" & strDetailHeight & "px;width:38px;border-style:outset;border-width:1px""  align=center>Status</td>"
Response.Write "<td class=""but"" style=""height:" & strDetailHeight & "px;width:89px;border-style:outset;border-width:1px""  align=left>Vendor Phone</td>"
'Response.Write "<td class=""but"" style=""height:" & strDetailHeight & "px;width:37px;border-style:outset;border-width:1px""  align=left>Status</td>"
Response.Write "<td class=""but"" style=""height:" & strDetailHeight & "px;width:32px;border-style:outset;border-width:1px""  align=center>Print</td>"
Response.Write "</tr>"

'Response.Write tHead
'Response.Write rs.RecordCount

c = 0


Do While not rs.EOF and c < REC_NUM

on error resume next

strDate = FormatDateTime(rs.Fields("ApptStartDate").Value,1)
strTime = ampm(hour(rs.Fields("ApptStartDate").Value),right("0" & minute(rs.Fields("ApptStartDate").Value),2))

vSTR = Split(rs("ApptStartDate")," ")
vDate = left(vSTR(0),instrrev(vSTR(0),"/")) & right(vSTR(0),2)
'vTime = vSTR(1) & " " & vSTR(2)
'vTime = left(vTime,instr(1,vTime," ")-4) & " " & right(vTime,2)
if rs("NoTime") then
	vTime = "<center title=""No time associated with this task""><font color=red>Note</font></center>"
else
	vTime = strTime
end if

'vAMPM = vSTR(2)
on error goto 0 

if strBColor = "EDFAE7" then
	strBColor = "FFFFFF"
else
	strBColor = "EDFAE7"
end if

'strStr2 = "javascript:showTextActivate('" & raid & "')"
raid = rs("AppointmentID")


'str = "<span >"

str = "<tr id=tr" & raid & " bgcolor=#" & strBColor & " class=normclass onmousedown=""mosel(" & raid & ")"" onmouseout=""mout(" & raid & ")"" onmouseover=""mo(" & raid & ", " & c & ")"">"

str = str & "<td id=td0" & raid & " class=bord align=right><div style=""overflow:hidden;height:" & strDetailHeight & "px;width:35px""> " &  raid & "</div></td>"
str = str & "<td id=td1" & raid & " class=bord align=right><div style=""overflow:hidden;height:" & strDetailHeight & "px;width:45px""> " &  vDate & "</div></td>"
str = str & "<td id=td2" & raid & " class=bord align=right><div style=""overflow:hidden;height:" & strDetailHeight & "px;width:52px""> " & vTime & "&nbsp;</div></td>"

str = str & "<td id=td3" & raid & " class=bord><div style=""overflow:hidden;height:" & strDetailHeight & "px;width:20px""> " & rs("Room") & "</div></td>"
str = str & "<td id=td4" & raid & " class=bord><div style=""overflow:hidden;height:" & strDetailHeight & "px;width:16px""> " & rs("Salutation") & "</div></td>"
str = str & "<td id=td5" & raid & " class=bord><div style=""overflow:hidden;height:" & strDetailHeight & "px;width:45px""> " & rs("GuestFirstName") & "</div></td>"
str = str & "<td id=td6" & raid & " class=bord><div style=""overflow:hidden;height:" & strDetailHeight & "px;width:65px""> " & rs("GuestLastName") & "</div></td>"
str = str & "<td id=td7" & raid & " class=bord><div style=""overflow:hidden;height:" & strDetailHeight & "px;width:74px""> " & rs("Action") & "</div></td>"

'str = str & "<td id=td8" & raid & " class=""bord"" nowrap><div style=""overflow:hidden;height:" & strDetailHeight & "px;width:120px""> " & rs("LocationText") & "</div></td>"
str = str & "<td id=td8" & raid & " class=""bord"" nowrap><div style=""overflow:hidden;height:" & strDetailHeight & "px;width:120px""> " & rs("Vendor") & "</div></td>"
str = str & "<td id=td9" & raid & " align=""right"" class=""bord"" nowrap><div style=""overflow:hidden;height:" & strDetailHeight & "px;width:20px""> " & rs("Amount") & "</div></td>"

select case rs.Fields("Status")
	case "o"
		strStatus = "Open"
	case "p"
		strStatus = "Pending"
	case "c"
		strStatus = "Closed"
	case "x"
		strStatus = "Canceled"
	case "n"
		strStatus = "N/A"
	case "w"
		strStatus = "Waiting List"
	case "r"
		strStatus = "Reconfirmed"
end select

if rs("Note") = false then
	if rs("Status") <> "c" Then
		str = str & "<td id=td10" & raid & " class=""bord"" align=center><div style=""color: red;overflow:hidden;height:" & strDetailHeight & "px;width:38px"">" & strStatus & "</div></td>"
	Else
		str = str & "<td id=td10" & raid & " class=""bord"" align=center><div style=""color: green;overflow:hidden;height:" & strDetailHeight & "px;width:38px"">" & strStatus & "</div></td>"
	End If
else
	if rs("Status") <> "c" Then
		str = str & "<td id=td10" & raid & " class=""bord"" align=center><div style=""color: red;overflow:hidden;height:" & strDetailHeight & "px;width:38px"">Note</div></td>"
	Else
		str = str & "<td id=td10" & raid & " class=""bord"" align=center><div style=""color: green;overflow:hidden;height:" & strDetailHeight & "px;width:38px"">Note</div></td>"
	End If
end if

str = str & "<td id=td11" & raid & " class=""bord""><div style=""overflow:hidden;height:" & strDetailHeight & "px;width:89px""> " & rs("Phone") & "</div></td>"
'str = str & "<td id=td12" & raid & " class=""bord""><div style=""overflow:hidden;height:" & strDetailHeight & "px;width:37px""> " & rs("ClosedBy") & "</div></td>"
str = str & "<td id=td12" & raid & " class=""bord"" style=""padding-top: 0px;padding-left: 0px;"" valign=top align=center><div style=""padding-top: 0px; overflow:hidden;height:16px;width:32px""><input onclick=""updateIDs(this)"" type=checkbox id=c" & rs("AppointmentID") & " ></div></td></tr>" & vbcrlf
							
'If Len(rs("ApptText")) > 0 Then
'	str = str & "<td class=""exp"" onClick=""ExpandNote(this,'d" & c & "')"">+</td>"
'	str = str & "</tr>"
'	str = str & "<tr class=""normclass""><td colspan=8><div id=d" & c & " style=""display:none;background-color:yellow"">" & rs("ApptText") & "</div></td></tr>"
'End IF

Response.Write str

'Figure out colors
'If rs.Fields("Closed") <> 0 Then
'	'Closed Task
'	strColor="closedtask"
'	strBackColor="lightgreen"
'Else
'	'Open Task
'	strColor="opentask"
'	strBackColor= "#FFB3B3" '"#FFD2D2" '"#FF6666" '"red" '"#C10000" '
'End If
select case rs.Fields("Status").Value
	case "c"
		'Closed Task
		strColor="closedtask"
		strBackColor="lightgreen"
	case "p"
		'Pending Task
		strColor="pendingtask"
		strBackColor= "#FFB353" '"#FFFFA8" '"#FFD2D2" '"#FF6666" '"red" '"#C10000" '
	case "o"
		'Open Task
		strColor="opentask"
		strBackColor= "#FFB3B3" '"#FFD2D2" '"#FF6666" '"red" '"#C10000" '
end select					
					
' However, if this is a note, then the backcolor is White no matter what the status of this note is :-)
if rs.Fields("Note") <> 0 then
	strColor="closedtask"
	strBackColor="white"
end if 
					

on Error Resume Next

rs.MoveNext
c = c + 1

Loop

prl = c-1	

Response.Write "<TR>"
n = ""
p = ""
nnu = "navNone"
pnd = ""

If not rs.EOF Then 
	'IF (rs.Bookmark-REC_NUM)>1 Then prevWrite = 1 
	if (rs.Bookmark <= rs.RecordCount) then	n = "<a style=""text-decoration:none"" href=""TaskSearchDetail.asp?p=n&prl=" & prl & """>Next&nbsp;>></a>"
	nnu = "navUp"
	'prevWrite = 1 or 
End If
'Response.Write rs.Bookmark-1 & " / "
'Response.Write REC_NUM
if not (rs.EOF and rs.BOF) then
	if rs.EOF then
		if Request.QueryString("p") <> "" Then 
			p = "<a style=""text-decoration:none"" href=""TaskSearchDetail.asp?p=p&prl=" & prl & """><<&nbsp;Prev</a>"
			pnu = "navUp"
		end if
	else
		if rs.Bookmark-1 > REC_NUM then
			if Request.QueryString("p") <> "" Then 
				p = "<a style=""text-decoration:none"" href=""TaskSearchDetail.asp?p=p&prl=" & prl & """><<&nbsp;Prev</a>"
				pnu = "navUp"
			end if
		end if
	end if
end if
'If rs.EOF Then Response.Write ("<tr>")

Response.Write "<TD colspan=2 id=tdp class=""" & pnu & """ width=""63px"">" & p & "</TD>"
Response.Write "<TD id=tdn class=""" & nnu & """ width=""52px"">" & n & "</TD>"
Response.Write "</TR>"
%>

<SCRIPT>
var lastColor = "";
var SelectedBGColor = "";
var SelectedID = 0;
var Red = "#ff5e5e"
if(parent.document.all("divTotal"))
{
<%
if not (rs.bof and rs.eof) then
	if rs.eof then
		rs.movelast
		strRecno = rs.bookmark-prl
		rs.movenext
	else
		strRecno = rs.bookmark-REC_NUM
	end if
	intRecCount = rs.recordcount

	lastpagerec = cint(strRecno)+prl

	Response.Write "parent.document.all(""divTotal"").innerText = ""Records " & strRecno & "-" & lastpagerec & " of " & intRecCount & """" & vbscrlf
else
	Response.Write "parent.document.all(""divTotal"").innerText = ''" & vbscrlf
end if
%>
}

<%
if Request.QueryString ("p") = "" then
	response.write "checkAll(true);" & vbcrlf
else
	response.write "fillChecks();"
end if
%>

function checkAll(o)
{
	var coll = document.all.tags("INPUT");
	for(var i=0; i<coll.length; i++)
		if(coll(i).id != "txtSQL" && coll(i).id != "txtRecCount")
		{
			if(coll(i).type = "checkbox")
				{
				coll(i).checked = o;
				//updateIDs(coll(i));
				}
		}
}

function fillChecks()
{
	var coll = document.all.tags("INPUT");
	for(var i=0; i<coll.length; i++)
		if(coll(i).id != "txtSQL")
		{
			if(coll(i).type = "checkbox")
				{
				//var o = (parent.document.all.txtUnSelectedIDs.value).indexOf((coll(i).id.substr(1))+", ") == -1;
				//alert(parent.strUnSelectedIDs)
				var o = (parent.strUnSelectedIDs).indexOf((coll(i).id.substr(1))+", ") == -1;
				coll(i).checked = o;
				}
		}
}
			
function mo( t, c )
{
	HoverOn(t);
	//parent.document.all("divTotal").innerText = 'Record# '+(<%'=rs.bookmark-REC_NUM%>+c) + ' of <%'=rs.recordcount%>';
}
function mout( t )	{ HoverOff(t, false); }
function mosel( t )
{ 
	Selected(t);
	showTextActivate(t);
}

function HoverOn( t )
{
	if(document.all("tr"+t).style.backgroundColor != Red)
	{
		lastColor = document.all("tr"+t).style.backgroundColor
		document.all("tr"+t).style.backgroundColor = "#CFCFF5"
}	}

function HoverOff( t, f )
{
	if(document.all("tr"+t).style.backgroundColor != Red)
		document.all("tr"+t).style.backgroundColor = lastColor;
}

function Selected( t )
{
	SelectedBGColor = lastColor;

	if(SelectedID != 0)
		document.all("tr"+SelectedID).style.backgroundColor = SelectedBGColor;
		
	document.all("tr"+t).style.backgroundColor = Red;
	SelectedID = t;
}

function updateIDs( o )
{
	var strID = o.id;
	var strVal = strID.substr(1)
	//var f = parent.document.all.txtUnSelectedIDs;
	if(!o.checked)
	{
		//f.value += strVal + ", ";
		parent.strUnSelectedIDs += strVal + ", ";
	}
	else
	{
		var pos = (parent.strUnSelectedIDs).indexOf(strVal+", ")
		var len = (strVal+", ").length
		//f.value = (f.value).substr(0,pos) + (f.value).substr(pos+len)
		parent.strUnSelectedIDs = parent.strUnSelectedIDs.substr(0,pos) + parent.strUnSelectedIDs.substr(pos+len)
	}
	//alert(parent.strUnSelectedIDs);
}
//alert("<%=remote.Session("SQLNOORDER")%>");
</script>
<%if not (rs.EOF and rs.BOF) then
	intRC = rs.RecordCount
else
	intRC = 0
end if
Response.Write "<input type=hidden id=txtRecCount value=" & intRC & ">"
%>
</BODY>

</HTML>
