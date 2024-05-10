<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))

dim rs, cn

set cn = server.CreateObject("adodb.connection")
cn.Open Application("sqlInnSight_ConnectionString")

'set rs = server.CreateObject("adodb.recordset")

if Request.QueryString("mode") = "Appointment" then
	mode = "Appointment"
	strDisplay = "none"
	strBGColor = "powderblue"
else
	mode = "Switchboard"
	strDisplay = "inline"
	strBGColor = "#F0c568"
end if

strLastName = ""
strFirstName = ""
strPrimaryPhone = ""
strHotelGuestID = ""
%>

<HTML>
<HEAD>
<!--#INCLUDE file="PhoneMask.asp"-->
<title>Guest Profile Search</title>
<style>
	.mainFont		{ font-family:tahoma;font-size:11px }
	.mainTables		{ font-family:tahoma;font-size:11px;border-style:solid;border-width:1px;border-color:black }

	.id				{ font-family:tahoma;font-size:11px;width:50px }
	.detailBody		{ background-color:#eeebbb;font-family:tahoma;font-size:11px }

	.ext			{ font-family:tahoma;font-size:11px;width:48px }
	.phonenumber	{ font-family:tahoma;font-size:11px;width:100px }
	.phonenote		{ font-family:tahoma;font-size:11px;width:325px }
	.GuestHeader	{ font-family:tahoma;font-size:11px;color:black }
	.long			{ font-family:tahoma;font-size:11px;width:181px }
	.dh				{ color:white;background-color:black }
	.PhoneType		{ font-family:tahoma;font-size:11px;width:142px }
	.buttons		{ font-family:tahoma;font-size:11px;width:100px }
</style>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
var intLineCount = 10000000000;
var booFirstLoad = true;
var w;
var mode = "<%=mode%>"

function window_onload() {
	<%if mode = "Appointment" then%>
	w = window.dialogArguments;
	<%
	'response.write "document.all('txtLastName').value = '" & strLastName & "';" & vbcrlf
	'response.write "document.all('txtFirstName').value = '" & strFirstName & "';" & vbcrlf
	'response.write "document.all('txtHotelGuestID').value = '" & strHotelGuestID & "';" & vbcrlf
	'response.write "FillPhone('txtPrimaryPhone','" & strPrimaryPhone & "');" & vbcrlf
	%>
	var phone = w.document.all('txtGuestPhone').value
	var pos = phone.indexOf(" ");
	if(pos > 0)
		phone = phone.substr(0,pos);
	
	document.all('txtLastName').value = w.document.all('txtGuestLastName').value;
	document.all('txtFirstName').value = w.document.all('txtGuestFirstName').value;
	document.all('txtGuestID').value = w.document.all('txtRealGuestID').value;
	document.all('txtHotelGuestID').value = w.document.all('txtGuestID').value;
	FillPhone('txtPrimaryPhone',phone.replace(/[-()\s]/g,''));
	//if <%=Request.Querystring("load")%>
	doSubmit()
	<%end if%>
}

function enterSubmit()
{
	if(window.event.keyCode==13)
		{
		window.event.returnValue=false;
		doSubmit();
		}
}
function cmdEdit_onclick() {
	if(document.all("txtGuestID").value != "")
		if(window.showModalDialog("GuestProfileMain.asp?gid=" + document.all("txtGuestID").value,"","dialogHeight:520px;dialogWidth:800px;scroll:no;status:no;center:yes") == "refresh")
			doSubmit();
}

function cmdadd_onclick() {
	<%if mode = "Appointment" then%>
	var ln = window.frmSubmit.txtLastName.value;
	var fn = window.frmSubmit.txtFirstName.value;
	var ph = window.frmSubmit.txtPrimaryPhone.value;
	var sal = w.document.all("pvSalutation").value;
	var em = w.document.all("txtGuestEmail").value;
	var cct = w.document.all("cboChargeTo").value;
	var cn = w.document.all("txtNumber").value;
	var exp = w.document.all("CCExp").value;
	<%else%>
	var ln = '';
	var fn = '';
	var ph = '';
	var sal = '';
	var em = '';
	var cct = '';
	var cn = '';
	var exp = '';
	<%end if%>
	
	var url = "GuestProfileMain.asp?gid=0&ln="+ln+"&fn="+fn+"&ph="+ph+"&sal="+sal+"&em="+em+"&cct="+cct+"&cn="+cn+"&exp="+exp
	//alert(url)
	var result = window.showModalDialog(url,"","dialogHeight:520px;dialogWidth:800px;scroll:no;status:no;center:yes");

	
	<%if mode = "Appointment" then%>
	if( result == "refresh")
		doSubmit();
	<%end if%>
}

function cmdClose_onclick() {
	//parent.document.all("divGPLookup").style.display = 'none';
	window.returnValue = ","
	window.close();
}

function cmdSelect_onclick() {
	/*
	parent.booForceRefresh = true;
	parent.document.all("divGPLookup").style.display = 'none';
	parent.document.all("txtRealGuestID").value = document.all("txtGuestID").value;
	if(document.all("txtDisplayID").value.replace(/\s/g,'') == "")
		{
		parent.document.all("txtGuestID").value = document.all("txtGuestID").value;
		parent.lookupGuestID( 1 );
		parent.document.all("txtGuestID").value = "";
		}
	else
		{
		parent.document.all("txtGuestID").value = document.all("txtDisplayID").value;
		parent.lookupGuestID( <%=request.querystring("GPSearchID")%> );
		}
	*/
	window.returnValue = document.all("txtDisplayID").value + ',' + document.all("txtGuestID").value;
	window.close();
}

function clearVendor() {
	document.all("txtHotelGuestID").value = "";
	document.all("txtLastName").value = "";
	document.all("txtFirstName").value = "";
	FillPhone('txtPrimaryPhone','');
	document.all("txtGuestID").value = "";
	doSubmit();
}

function cmdMerge_onclick() {
	alert(window.frmSubmitFrame.strSelected);
}

function doSubmit()
{
	document.frmSubmitFrame.document.body.innerHTML = "<div style=color:silver;font-family:arial;font-size:24px;position:absolute;top:100px;left:300px valign=middle align=center>Loading...</div>";
	window.frmSubmit.submit();
}
</script>

<SCRIPT LANGUAGE=javascript FOR=cmdadd EVENT=onclick>
<!--
 cmdadd_onclick()
//-->
</SCRIPT>
</HEAD>
<BODY bgcolor="<%=strBGColor%>" LANGUAGE=javascript onload="return window_onload()" leftmargin=8 bottommargin=0 topmargin=0>
<input type=hidden id=txtGuestID name=txtGuestID>
<input type=hidden id=txtDisplayID name=txtDisplayID>

<table bbgcolor=#F0c568 sstyle="border-style:outset;border-width:1px" width="100%" cellpadding=0 cellspacing=2 class=GuestHeader>
  <tr>
  <td>
  
	  <table height=100% cellpadding=0 cellspacing=0>
		<tr>
		<td>
			<form id=frmSubmit name=frmSubmit target=frmSubmitFrame action="GuestProfileSetupDetail.asp?mode=<%=mode%>" method=post>
			<table onkeydown=enterSubmit() width=100% bbgcolor=fucia cellpadding=0 cellspacing=2 class=GuestHeader>
			<tr><td>
				<td align=right>Last Name:</td>
				<td width=100px><input style=width:90px class=mainFont type=text id=txtLastName name=txtLastName></td>
				<td align=right>First Name:</td>
				<td width=120px><input style=width:80px class=mainFont type=text id=txtFirstName name=txtFirstName></td>
				<td style=width:36px align=right>Phone:</td>
				<td width=120px>
					<script language=javascript>
							var x = CreatePhoneField( 'txtPrimaryPhone', 'font-family: Tahoma; font-size: 11', '13px', 100, null, null, 'white', true );
					</script>
				</td>
				<td>ID:</td>
				<td width=66px><input class=id type=text id=txtHotelGuestID name=txtHotelGuestID></td>
				<!--td>GCN ID:</td>
				<td><input class=mainFont style=padding-left:4px;width:70px type=text id=txtGID name=txtGID></td-->
				<td align="center" onclick="clearVendor()" onmousedown="this.style.borderStyle='inset'" onmouseout="this.style.borderStyle='outset'" onmouseup="this.style.borderStyle='outset'" style="width:20px;border-style:outset;border-width:1px;border-color:silver;background-color:menu"><img src="images/eraser.gif" WIDTH="16" HEIGHT="14" title="Clear Criteria"></td>
				<td width=60px align=right>
					<img title="Get listing based on criteria" style="cursor:hand" id="cmdRefresh" name="cmdRefresh" src="images/btn_go.gif" align="absMiddle" border="0" name="go" onclick="doSubmit()" WIDTH="35" HEIGHT="21">
				</td>
			</td></tr>
			</form>
			</table>
		</tr>
	</table>
</td>

</tr>
</table>
<table cellpadding=0 cellspacing=0 width="100%">
<tr>
	<td colspan=2>
		<table cellpadding=3px style=overflow:hidden class=mainFont width="100%">
			<tr height=20px bgcolor=black style=color:white>
				<td width=80px>Salutation</td>
				<td width=110px>Last Name</td>
				<td width=100px>First Name</td>
				<td width=129px>Company</td>
				<td width=95px>Phone</td>
				<td width=50px>ID</td>
				<td width=16px>&nbsp;</td>
				<td width=12px>&nbsp;</td>
				<td width=12px style=background-color:transparent>&nbsp;</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
<td colspan=2>
	<iframe AllowTransparency=yes frameborder=no scrolling=auto src="GuestProfileSetupDetail.asp?load=1" style="width:100%;height:292px;border-style:groove;border-width:1px" id=frmSubmitFrame name=frmSubmitFrame></iframe>
</td>
</tr>
<tr style=padding-top:10px;>
<td>
	<input class=buttons type=button value="Add" id=cmdAdd name=cmdAdd>&nbsp;
	<input class=buttons type=button value="Edit" id=cmdEdit name=cmdEdit LANGUAGE=javascript disabled onclick="return cmdEdit_onclick()">&nbsp;
	<input style=visibility:hidden class=buttons disabled type=button value="Merge" id=cmdMerge name=cmdMerge LANGUAGE=javascript disabled onclick="return cmdMerge_onclick()">
	<input style=visibility:hidden class=buttons disabled type=button value="Disable" id=cmdDelete name=cmdDelete>
</td>
<td align=right>
	<%if mode = "Appointment" then
		strSelectDisplay = "inline"
		strAddTask = "none"
	else
		strSelectDisplay = "none"
		strAddTask = "inline"
	end if%>
	<input style="display:<%=strSelectDisplay%>" class=buttons type=button value="Select" id=cmdSelect name=cmdSelect LANGUAGE=javascript disabled onclick="return cmdSelect_onclick()">&nbsp;
	<input style="display:<%=strAddTask%>" class=buttons type=button value="Add Task" id=cmdAddTask name=cmdAddTask LANGUAGE=javascript disabled onclick="return cmdSelect_onclick()">&nbsp;
	<input class=buttons type=button value="Close" id=cmdClose name=cmdClose LANGUAGE=javascript onclick="return cmdClose_onclick()">
</td>
</tr>
</table>
</BODY>
</HTML>
<%
'rs.Close        
'set rs = nothing

cn.Close
set cn = nothing

function sq(str)
	if isnull(str) then
		sq = ""
	else
		sq = replace(str,"'","''")
	end if
end function
%>
