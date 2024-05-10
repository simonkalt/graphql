<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))
%>

<html>
<head>
	<style>
		.histtable	{font-family:tahoma;font-size:11px}
		.big		{font-family:arial;font-size:18px}
	</style>
	<title>Guest History for <%=request.querystring("name")%></title>
</head>
<script language=javascript>
	var aid;
	
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

	function md( id )
	{
		aid = id;
		showTextActivate(aid);
		//alert(aid);
	}
	
	function mo( strRow )
	{
		//document.all('tbl_'+strRow).style.borderWidth = '2px';
		document.all('tbl_'+strRow).style.fontWeight = 'bold';
	}
	
	function mout( strRow )
	{
		//document.all('tbl_'+strRow).style.borderWidth = '1px';
		document.all('tbl_'+strRow).style.fontWeight = 'normal';
	}

	function showTextActivate( strField ) { 
		var str = "TaskPadToolTip.asp?id=" + strField +"&UserID=" + <%=remote.session("FloatingUser_UserID")%>
		document.all("frameSTA").src = str;
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
	  }
	  else
	  {
		eval(code)

		var el = document.all.tooltip //, a = new Array(), alen = 0;
		var frel = document.frames.frameToolTip.document.all.tbdyToolTip
		var frid = aid;
		var booLastWasNote = false;
	  
		//el.style.pixelTop = (screen.availHeight-400)/2;
		//el.style.pixelLeft = (screen.availWidth-580)/2;
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

		document.frames.frameToolTip.txtReminderStatus.value = a[3];
	  
		if(a[3] == " ")
	 		document.frames.frameToolTip.cmdReminder.value = "Add Reminder";
		else
			document.frames.frameToolTip.cmdReminder.value = "Edit Reminder";
	  
		if(strColor=="lightgreen" || strColor=="white")
			oRow.style.backgroundColor = "#C2F5C2";
		else
			oRow.style.backgroundColor = strColor;  	

		var cboStatus = document.frames.frameToolTip.document.all("cboStatus");
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
	  
	  var tooltip = document.frames.frameToolTip.document;
	  
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
			tooltip.all("cboStatus").disabled = false
	  else
			tooltip.all("cboStatus").disabled = true
		
	  el.style.visibility = "visible"
	  booToolTipOpen = true;
	}

</script>
<body topmargin=4px leftmargin=4px rightmargin=4px bgcolor="powderblue">
<%
dim rs, cn, strSQL, strWhere, cid

cid = Remote.Session("CompanyID")
gid = Request.QueryString("gid")

set cn = server.CreateObject("adodb.connection")
set rs = server.CreateObject("adodb.recordset")
cn.Open Application("sqlInnSight_ConnectionString")

set rs = cn.Execute("sp_GuestProfileHistory " & gid)

Response.Write "<table cellspacing=0 cellpadding=0><tr><td>"
Response.Write "<div style=border-style:solid;border-width:1px;overflow:auto;width:684px;height:387px>"

if rs.EOF then
	Response.Write "<table height=100% width=100% cellspacing=1 cellpadding=0>"
	Response.Write "<tr><td class=big align=center valign=middle>No tasks exist for this guest.</td></tr>"
else
	Response.Write "<table cellspacing=0 cellpadding=0>"
	do until rs.EOF
		factor = 0
		aid = rs.Fields("AppointmentID").Value
		select case rs.Fields("Status").Value
			case "c"
				'Closed Task
				strColor="closedtask"
				strBackColor="lightgreen"
				strStatusImage = "<img title=""Closed"" border=0 src=""images/Closed.gif"">"
			case "p"
				'Pending Task
				strColor="pendingtask"
				strBackColor= "#FFB353"
				strStatusImage = "<img title=""Pending"" border=0 src=""images/pending.gif"">"
			case "o"
				'Open Task
				strColor="opentask"
				strBackColor= "#FFB3B3"
				strStatusImage = "<img title=""Open"" border=0 src=""images/open.gif"">"
			case "x"
				'Canceled Task
				strColor="opentask"
				strBackColor= "lightblue"
				strStatusImage = "<img title=""Canceled"" border=0 src=""images/canceled.gif"">"
			case "r"
				'ReConfirmed Task
				strColor="opentask"
				strBackColor= "#82FFE0"
				strStatusImage = "<img title=""Re-Confirmed"" border=0 src=""images/Recon.gif"">"
			case "n"
				'Not Available Task
				strColor="pendingtask"
				strBackColor= "lavender"
				strStatusImage = "<img title=""Not Available"" border=0 src=""images/NA.gif"">"
				factor = 20
			case "w"
				'Not Available Task
				strColor="pendingtask"
				strBackColor= "#cc99ff"
				strStatusImage = "<img title=""Wait List"" border=0 src=""images/WL.gif"">"
				factor = 20
		end select					
		Response.Write "<tr onmouseout=mout('" & aid & "') onmouseover=mo('" & aid & "') onmousedown=md(" & aid & ") style=cursor:hand>"
		Response.Write "<td><table border=1 id=tbl_" & aid & " class=histtable style=""background-color:" & strBackColor & """ cellpadding=1 cellspacing=0><tr>"
		Response.Write "<td style=border-right-style:none><div nowrap style=overflow:hidden;width:630px;>&nbsp;" & FormatAppHist(true,rs.Fields("ApptStartDate").Value, rs.Fields("ApptEndDate").Value, rs.Fields("Room").Value, rs.Fields("Salutation").Value, rs.Fields("GuestName").Value, rs.Fields("GuestFirstName").Value, rs.Fields("GuestLastName").Value, rs.Fields("ActionTypeText").Value, rs.Fields("ActionText").Value, rs.Fields("LocationText").Value, rs.Fields("ApptText").Value, rs.Fields("Status").Value, false,"",false,"") & "</div></td>" ', rs.Fields("ReminderNote").Value,0
		Response.Write "<td style=border-left-style:none align=right width=24px>" & strStatusImage & "&nbsp;</td>"
		Response.Write "</tr></table></td>"
		Response.Write "</tr>"
		rs.MoveNext
	loop
end if

Response.Write "</table>"
Response.Write "</div>"

Response.Write "</td></tr><tr><td style=padding-top:6px align=right>"
Response.Write "<input onclick=window.close() style=width:100px class=histtable type=button id=cmdClose value=Close>&nbsp;&nbsp;"
Response.Write "</td></tr></table>"

rs.Close
set rs = nothing
cn.Close
set cn = nothing

Function FormatAppHist(booShowLocation, pstrStartTime, pstrEndTime, pstrRoom, pstrSalutation, pstrName,pstrFirstName,pstrLastName, pstrActionType, pstrAction, pstrLocation, pstrSubject, pstrStatus, pstrNoTime, pstrNotesDetails, pbooReminder, pstrReminderNote)
	
	if pbooReminder then
		if pstrNoTime then
			strAppt = "<b style=color:maroon>Reminder:</b> " & pstrReminderNote
			strDash = " - "
			strSpacer = "&nbsp;"
		else
			strAppt = trim(CustomTime( pstrStartTime, "" )) & " - <b style=color:maroon>Reminder:</b> " & pstrReminderNote
			strDash = " - "
			strSpacer = ""
		end if
	else
		d = formatdatetime(pstrStartTime,vbLongDate)
		strAppt = trim(mid(d,instr(1,d,",")+1))
		if pstrNoTime then
			strAppt = "Note: "
			strDash = ""
			strSpacer = "&nbsp;"
		else
			strAppt = strAppt & " - " & CustomTime( pstrStartTime, "" )
			if pstrStartTime <> pstrEndTime then
				strAppt = strAppt & "-" & CustomTime( pstrEndTime, "" )
			end if
			strDash = " - "
			strSpacer = ""
		end if
		
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
			pstrSubject = trim(pstrSubject)
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

		'hide for history
		'strAppt = strAppt & (strDash + ((pstrSalutation & " ") + (vName)))
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
	
	FormatAppHist = strSpacer & trim(strAppt)
End Function

function CustomTime( strTime, padding )
	on error resume next
	CustomTime = right("0" & left(strTime,instrrev(strTime,":")-1),5) & lcase(right(strTime,2))
	if err.number > 0 then
		CustomTime = "(no time)"
	end if
end function
%>
<iframe onload=sta() src="LoadingAppointment.asp" id=frameSTA style=display:none;visibility:hidden></iframe>
<div ID="tooltip" STYLE="top:10px;left:60px;font-family: Helvetica; font-size: 8pt; position: absolute; z-index: 200; visibility: hidden; width:250px;">
	<iframe height="404" width="580" frameborder="0" style="border-style: none; border-width: 1px;" src="tooltip.asp?from=gphistory" id="frameToolTip" scrolling="no"></iframe>
</div>

</body>
</html>
