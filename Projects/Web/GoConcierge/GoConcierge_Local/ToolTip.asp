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
<!--
	.Button_Normal		{ font-family: tahoma; font-size: 11px; width: 44px; }
	.Button_abNormal	{ font-family: tahoma; font-size: 11px; width: 90px; }
	.But				{ font-family: tahoma; font-size: 11px; }
-->
</style>
<script src=CheckIfTaskExists.asp language=javascript></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	<%if Request.QueryString("from") = "gphistory" then%>
		window.cmdCopy.style.visibility = "hidden";
		window.cboLetterhead.style.visibility = "hidden";
		window.cboStatus.style.visibility = "hidden";
		window.cmdDelete.style.visibility = "hidden";
		window.cmdEdit.style.visibility = "hidden";
		window.cmdGTL.style.visibility = "hidden";
		window.cmdPrint.style.visibility = "hidden";
		window.cmdReminder.style.visibility = "hidden";
	<%end if%>	
}

//-->
</SCRIPT>
</head>
<body id="bdy" bgcolor="#FFFFE1" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0" style="border-style:solid;border-color:black;border-width:1px" onload="return window_onload()">
	<table id="tbl1" width="100%" cellpadding="2">
		<tr>
			<td style="height: 22px;" align="right">
				<img onmousedown="this.src=imgDown.src;" onmouseup="this.src=imgUp.src;" onmouseout="this.src=imgUp.src;" onclick="parent.tooltip.style.visibility='hidden';" src="images/WindowClose.gif" WIDTH="16" HEIGHT="14">
			</td>
		</tr>
		<tr>
			<td>
				<div style="overflow: auto; height: 334px;">
					<table id="tblToolTip" cellpadding="2" cellspacing="0" style="font-family: Tahoma; font-size: 8pt;" width="100%">
						<tbody id="tbdyToolTip"></tbody>
					</table>
				</div>
			</td>
		</tr>
		<tr>
			<td align="center" style="border-top-style: solid; border-top-width: 1px; border-top-color: black;">
			<table cellpadding=0 cellspacing=3 border=0>
				<tr>
				<td>
					<select class="But" size="1" name="cboLetterhead">
						<option value="Yes">Letterhead</option>
						<option selected value="No">Plain Paper</option>
					</select>
				</td>
				<td>
				<table class=Label cellpadding=0 cellspacing=0><tr><td>
					<select style="width:90px" class="But" onChange="statusOnChange()" id="cboStatus" name="cboStatus">
					<option value="o">Open</option>
					<option value="p">Pending</option>
					<option value="c">Closed</option>
					<option value="r">Reconfirmed</option>
					<option value="x">Canceled</option>
					<option value="n">Not Available</option>
					<option value="w">Wait List</option>
					</select>
				</table>	
				</td>
				<td><input type="button" id="cmdReminder" value="Add Reminder" class="Button_abNormal" onClick="JavaScript:doCheckAID('cmdReminder')"></td>
				<td><input type="button" id="cmdPrint" value="Print" class="Button_Normal" onClick="JavaScript:doCheckAID('cmdPrint')"></td>
				<td><input type="button" id="cmdEdit" value="Edit" class="Button_Normal" onClick="JavaScript:doCheckAID('cmdEdit')"></td>
				<td><input type="button" id="cmdDelete" value="Delete" class="Button_Normal" onClick="JavaScript:doCheckAID('cmdDelete')"></td>
				<td><input type="button" id="cmdCopy" value="Copy" class="Button_Normal" onClick="JavaScript:doCheckAID('cmdCopy')"></td>
				<td><input type="button" id="cmdGTL" value="Letter" class="Button_Normal" title="Guest Letter" onClick="JavaScript:doCheckAID('cmdGTL')"></td>
				<td><input type="button" id="cmdHide" value="Exit" style="color:red" class="Button_Normal" onclick="parent.tooltip.style.visibility='hidden';"></td>
				<td>
					<table width="100%" cellspacing="1" cellpadding="0" topmargin="0">
					  <tr>
						<td width="20px" align="center" style="border-width:1px;border-color:lightyellow;border-style:solid" onmouseover="this.style.borderStyle='outset';this.borderColor='';" onmouseout="this.style.borderStyle='solid';this.borderColor='lightyellow';" onmousedown="this.style.borderStyle='inset';" onmouseup="this.style.borderStyle='solid';this.borderColor='lightyellow';">
							<img onclick=doCheckAID(this.id) src="images\prev.gif" id="cmdUP" title="Show Previous Task" WIDTH="16" HEIGHT="16">
						</td>
						<td width="20px" align="center" style="border-width:1px;border-color:lightyellow;border-style:solid" onmouseover="this.style.borderStyle='outset';this.borderColor='';" onmouseout="this.style.borderStyle='solid';this.borderColor='lightyellow';" onmousedown="this.style.borderStyle='inset';" onmouseup="this.style.borderStyle='solid';this.borderColor='lightyellow';">
							<img onclick=doCheckAID(this.id) src="images\next.gif" id="cmdDown" title="Show Next Task" WIDTH="16" HEIGHT="16">
						</td>
					  </tr>
					</table>
				</td>
				<!--td><input type="button" id="cmdEmail" value="E-Mail" style="color:Green" class="Button_Normal" onclick="processEmail();"></td-->
				
				</tr>
			</table>
			</td>
		</tr>
	</table>
	<input type="hidden" id="txtAppointmentID">
	<input type="hidden" id="txtReminderStatus">
</body>
</html>

<script language=javascript>
	var imgUp = new Image();
	var imgDown = new Image();
	imgUp.src = "images/WindowClose.gif";
	imgDown.src = "images/WindowCloseDown.gif";
	document.all("cboLetterHead").value = "<%=remote.Session("LetterHead")%>"
	
	function doCheckAID(s)
	{
		if(s=="cmdPrint" || s=="cmdEdit" || s=="cmdDelete" || s=="cmdCopy" || s=="cmdReminder" || s=="cmdGTL" || s=="cmdUP" || s=="cmdDown")
		{
				switch(s)
					{
					case "cmdUP":
						{
							var aid = document.all("txtAppointmentID").value;
							var ar = parent.idarr;
							for (i=0;i<ar.length;i++)
							{
								if (ar[i]==aid+'|0' || ar[i]==aid+'|1' ) break;
							}
							
							
							if (i > 0)
								parent.window.frames("frameTaskPad" + parent.intATF).showActivate(ar[i-1]);
							
							break;
						}
						
					case "cmdDown":
						{
							var aid = document.all("txtAppointmentID").value;
							var ar = parent.idarr;
							for (i=0;i<ar.length;i++)
							{
								if (ar[i]==aid+'|0' || ar[i]==aid+'|1' ) break;
							}
							
							if (i < ar.length-1)
								parent.window.frames("frameTaskPad" + parent.intATF).showActivate(ar[i+1]);
							
							break;
						}

					case "cmdReminder":
						{
						AddEditReminder();
						break;
						}
					case "cmdPrint":
						{
						PrintTask();
						break;
						}
					case "cmdEdit":
						{
							parent.tooltip.style.visibility='hidden';
							window.top('frameTaskPad'+window.top.intATF).document.all('divEditTask'+document.all('txtAppointmentID').value).click()
							break;
						}
					case "cmdDelete":
						{
						cmdDeleteonClick();
						break;
						}
					case "cmdCopy":
						{
						cmdCopyonClick();
						break;
						}
					case "cmdGTL":
						{
						cmdGTLonClick();
						break;
						}
					}

/*				}
				else
				{
					alert("The task you selected has been deleted by another user.  The calendar will now refresh.")
					window.top.TimerFunction(1);
					window.top.document.all("tooltip").style.visibility = "hidden";
				} */
		}
	}

	function updateTaskStatus (aid,sid)
	{
			try {
				closeReminder = "false";
				if(window.txtReminderStatus.value == 'o' && (window.cboStatus.value == 'c' || window.cboStatus.value == 'x'))
					if( calert("Would you like to close the reminder for this task?","Close Reminder") == 1 )
						closeReminder = "true";
				
				var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
				url = "AppointmentChangeStatus.asp?sid=" + sid + "&aid=" + aid + "&CloseReminder=" + closeReminder;
				xmlHttp.open( "POST" , url, false)
				xmlHttp.send()
				if (xmlHttp.responseText!='')
					alert (xmlHttp.responseText)
					
				xmlHttp = null
			} catch (e) { }
	}
	
	function statusOnChange()
	{
	
		var aid = document.all("txtAppointmentID").value;
		var sid = cboStatus.options (cboStatus.selectedIndex).value;
		
		<%if remote.Session("BPW") then%>
			updateTaskStatus (aid,sid)
		<%else%>
		
		var strSQL = "ValidatePassword.asp?v=2&aid=" + aid + "&Caption=Add/Edit Task"
		
		var	x = window.showModalDialog (strSQL,"","center:yes;status:no;scrollbars:no;dialogHeight:116px;dialogWidth:298px;")
			if (x != '')
			{
				switch (x)
				{
					case "Invalid Password" :
						alert ("Invalid password.")
						break
					case "Invalid Department":
						alert("You do not have rights to modify a task for this department.")
						break
					default:

						updateTaskStatus (aid,sid)
				}
			}	
			
		<%end if%>
		
		parent.tooltip.style.visibility='hidden';
		
		try {
			parent.Calobj.onDateChange (parent.Calobj.getVal(),1);
			}
		catch (e)
		 { }
		
	}
	
	function processEmail()
	{
	
		var ApptID = window.txtAppointmentID.value;
		var url = "EMailTaskDialog.asp?ID=" + ApptID;
		var xx = window.showModalDialog (url,null,"dialogheight: 120px; dialogwidth: 200px; status: no; center: yes; scroll: no");
		
		if (xx==1)
		{
			
			var strOptions = "center:yes;resizable:no;scroll:no;status:no;dialogHeight:340px;dialogWidth:340px";
			var strURL = "EMailVendorContactSelect.asp?ID=" + ApptID;
			var	xx = window.showModalDialog(strURL,window,strOptions)
			
			if (xx!='')
			{
			
				var h = 471
				var w = 516
				var strURL = xx
				var x = window.showModalDialog(strURL,window, "dialogheight: " + h + "px; dialogwidth: " + w + "px; center: yes; status: no; scroll: no")
			
			}
		
		}
		if (xx==2)
		{
				var str = '';
				
				var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
				xmlHttp.open("Get", "EmailTaskGuestValues.asp?ID=" + ApptID, false)
				xmlHttp.send()
				str = new String(xmlHttp.responseText)
				
				
				window.showModelessDialog (str, null, "center:yes;status:no;dialogHeight:518px;dialogWidth:666px;scroll:no")
		}
		
	}
	
	/*function test()
	{
		var x = window.bdy.createTextRange();
		x.expand("textarea")
		x.select();
		window.clipboardData.setData("Text",x.htmlText);
		
	}  */
</script>

<script language="vbscript">
	function calert(msg,title)
	    a = msgbox (msg, vbYesNo,title)
		if a = vbYes Then 
			calert = 1 
		else 
			calert = 0 
		End If
	End function

	function AddEditReminder()
		dim o
		set o = document.all("cmdReminder")
		<%if remote.Session("BPW") then%>
			if o.value = "Add Reminder" then
				str = "&NewFromToolTip=True"
			else
				str = ""
			end if
			x = window.showModalDialog("Reminder.asp?aid=" & document.all("txtAppointmentID").value & str, window, "dialogHeight:480px;dialogWidth:520px;status:no;scroll:no;center:yes")
			if x = "closeok" then
				parent.tooltip.style.visibility="hidden"
				parent.Calobj.onDateChange parent.Calobj.getVal(), 1
			end if
		<%else%>
			dim strSQL, strMode
			strSQL = "ValidatePassword.asp?v=1&Caption=" & o.Value
			x = showModalDialog(strSQL,"","center:yes;status:no;scrollbars:no;dialogHeight:116px;dialogWidth:298px;")
			if x <> "" then
				if x <> "Invalid Password" then
					if o.value = "Add Reminder" then
						str = "&NewFromToolTip=True"
					else
						str = ""
					end if
					x = window.showModalDialog("Reminder.asp?aid=" & document.all("txtAppointmentID").value & str, window, "dialogHeight:480px;dialogWidth:520px;status:no;scroll:no;center:yes")
					if x = "closeok" then
						parent.tooltip.style.visibility="hidden"
						parent.Calobj.onDateChange parent.Calobj.getVal(), 1
					end if
				else
					msgbox "Invalid password.",vbCritical,"Password"
				end if
			end if
		<%end if%>
	end function
	
	function cmdDeleteonClick()
		<%if remote.Session("BPW") then
			response.write "window.top.frames(""frameCalUpdate"").location = ""DeleteRights.asp?ID="" & window.txtAppointmentID.value & ""&Admin=" & remote.Session("FloatingUser_Admin") & """" & vbcrlf
		else%>
			dim strSQL
			strSQL = "ValidatePassword.asp?v=1&Caption=Delete Task"
			x = showModalDialog(strSQL,"","center:yes;status:no;scrollbars:no;dialogHeight:116px;dialogWidth:298px;")
			if x <> "" then
				if x <> "Invalid Password" then
					window.top.frames("frameCalUpdate").location = "DeleteRights.asp?ID=" & window.txtAppointmentID.value & "&Admin=" & left(mid(x,instr(1,x,"Admin=")+6),4)
				else
					msgbox "Invalid password.",vbCritical,"Password"
				end if
			end if
		<%end if%>
	end function
	
	function DeleteTask( booOK, RecID )
	'alert(window.txtAppointmentID.value)
		if booOK then
			If RecID > 0 Then
				xx = window.showModalDialog("RecurrenceDialog.asp?cap=Delete",,"dialogheight: 120px; dialogwidth: 200px; status: no; center: yes; scroll: no")
				If xx = 1 Then
						if msgbox("Are you absolutely sure you want to delete this occurence?",vbYesNo+vbQuestion,"Delete Task") = vbYes then
							parent.tooltip.style.visibility="hidden"
							window.top.frames("frameCalUpdate").location = "DeleteApptConfirm.asp?ID=" & window.txtAppointmentID.value & "&CalledFromTT=True&curdate=" & window.top.calObj.getVal()
						end if
				End IF
				If xx = 2 Then
						if msgbox("Are you absolutely sure you want to delete this series?",vbYesNo+vbQuestion,"Delete Task") = vbYes then
							parent.tooltip.style.visibility="hidden"
							window.top.frames("frameCalUpdate").location = "RecurrenceDelete.asp?ApptID=" & window.txtAppointmentID.value & "&RecID=" & RecID & "&CalledFromTT=True&curdate=" & window.top.calObj.getVal()
						end if
				End IF
			
			Else

				if msgbox("Are you absolutely sure you want to delete this task?",vbYesNo+vbQuestion,"Delete Task") = vbYes then
					parent.tooltip.style.visibility="hidden"
					window.top.frames("frameCalUpdate").location = "DeleteApptConfirm.asp?v=" & cstr((rnd(1)*10000)) & "&ID=" & window.txtAppointmentID.value & "&CalledFromTT=True&curdate=" & window.top.calObj.getVal()
				end if
			End IF
		else
			
			msgbox "Sorry, you do not have rights to delete this task.  See the administrator.",48,"Delete Task"
		end if	
	end function
	
	function cmdCopyonClick()
		<%if remote.Session("BPW") then
			response.write "CopyTask window.txtAppointmentID.value" & vbcrlf
		else%>
			dim strSQL
			strSQL = "ValidatePassword.asp?v=1&Caption=Copy Task"
			x = showModalDialog(strSQL,"","center:yes;status:no;scroll:no;dialogHeight:116px;dialogWidth:298px;")
			if x <> "" then
				if x <> "Invalid Password" then
					CopyTask window.txtAppointmentID.value
				else
					msgbox "Invalid password.",vbCritical,"Password"
				end if
			end if
		<%end if%>
	end function
	
	function cmdGTLonClick()
					pstr = "?ApptID=" & window.txtAppointmentID.value
					
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
						x = window.showModalDialog ("CustomReports/GuestLetterSelect2.asp" & pstr , null, (dOptions))
					Else
						x = retVal
					ENd If
					
					
					If Cstr(x) <> "" Then
							str = "ReportGuestTaskLetter.asp?ApptID=" & window.txtAppointmentID.value
							str = str & "&TemplateID=" & escape(x)  
					
							window.showModelessDialog str, null, "center:yes;status:no;dialogHeight:540px;dialogWidth:670px;scroll:no"
					End If
					
	end function
	
	'dim ctTimer ', winToClose
	'function CopyTaskAfterClose(win,id)
	'	'''''''' FIX THIS ''''''''''''''''
	'	'winToClose = win
	'	strFunc = "IsOpen(win," & id & ")"
	'	ctTimer = window.setInterval(strFunc,600)
	'end function
	'
	'function IsOpen(win,id)
	'	alert "In it"
	'	if not isobject(win) then
	'		'alert "did it"
	'		window.clearInterval ctTimer
	'		CopyTask id
	'	end if
	'end function
	
	function CopyTask(id)
		dim mode
		
		mode = window.top.showModalDialog("GetCopyMode.asp","","center:yes;status:no;scroll:no;dialogHeight:100px;dialogWidth:430px")
		if mode <> "" then
			parent.tooltip.style.visibility="hidden"
			window.top.frames("frameCalUpdate").location = "CopyTask.asp?ID=" & id & "&curdate=" & window.top.calObj.getVal() & "&mode=" & mode
		end if
	end function
	
	function PrintTask()
		parent.tooltip.style.visibility="hidden"
		booCCMask = not <%=remote.session("FloatingUser_VCCN")%>
		if not booCCMask then
			if instr(1,window.bdy.innerText,"CC Number:") > 0 then
				if msgbox("Display Credit Card Number Digits?",vbYesNo+vbQuestion,"Credit Card Information") = vbYes then
					booCCMask = false
				else
					booCCMask = true
				end if
			end if
		end if

		window.open "PrintTask.asp?Letterhead=" & document.all("cboLetterhead").value & "&amp;src=ToolTip&amp;Mode=v&amp;ID=" & document.all("txtAppointmentID").value & "&CCMask=" & booCCMask,"","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1"
	end function
</script>
