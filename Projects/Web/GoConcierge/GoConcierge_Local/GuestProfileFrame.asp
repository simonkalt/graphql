<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))

dim rs, cn, strScript

set cn = server.CreateObject("adodb.connection")
set rsSal = server.CreateObject("adodb.recordset")
cn.Open Application("sqlInnSight_ConnectionString")
set rsSal = cn.Execute("select * from tblSalutations order by Salutation")

set rs = server.CreateObject("adodb.recordset")

' for demo
'booSU = false
booSU = (remote.session("SuperUser") = 1)

'If remote.Session("ScreenHeight") < 750 Then
'	bodyheight = 300
'else
	bodyheight = 300
'end if

gid = Request.QueryString("gid")
if gid = "" then
	gid = "0"
end if
set rsGuest = server.CreateObject("adodb.recordset")
set rsGuest = cn.Execute("select * from tblGuest where GuestID = " & gid)

if rsGuest.EOF then
	strGuestID = "(New)"
	strSalutation = iif(Request.QueryString("sal")="","Mr.",Request.QueryString("sal"))
	strLastName = Request.QueryString("ln")
	strMiddleName = ""
	strFirstName = Request.QueryString("fn")
	strCompany = ""
	strTitle = ""
	strPrimaryPhone = Request.QueryString("ph")
	strPhoneExt = ""
	strEMail = Request.QueryString("em")
	strEMail2 = ""
	'strSmoking = "false"
	strPMSGuestID = ""
	strHotelGuestID = ""
	strGuestNote = ""
else
	strGuestID = gid
	strSalutation = rsGuest.Fields("Salutation").Value
	strLastName = escapeJS(trim(rsGuest.Fields("LastName").Value))
	strMiddleName = escapeJS(trim(rsGuest.Fields("MiddleName").Value))
	strFirstName = escapeJS(trim(rsGuest.Fields("FirstName").Value))
	strCompany = escapeJS(trim(rsGuest.Fields("Company").Value))
	strTitle = escapeJS(rsGuest.Fields("Title").Value)
	strPrimaryPhone = trim(rsGuest.Fields("PrimaryPhone").Value)
	strPhoneExt = trim(rsGuest.Fields("PhoneExt").Value)
	strEMail = rsGuest.Fields("EMail1").Value
	strEMail2 = rsGuest.Fields("EMail2").Value
	'booSmoking = rsGuest.Fields("Smoking").Value
	'if booSmoking then
	'	strSmoking = "true"
	'else
	'	strSmoking = "false"
	'end if
	strPMSGuestID = trim(rsGuest.Fields("PMSGuestID").Value)
	strHotelGuestID = trim(rsGuest.Fields("HotelGuestID").Value)
	strGuestNote = replace(escapeJS(trim(rsGuest.Fields("GuestNotes").Value)),vbCrLf,"<<crlf>>")
end if

function iif(expr,retval,retvalelse)
	if expr then
		iif = retval
	else
		iif = retvalelse
	end if
end function
%>
<html>

<head>

<!--#INCLUDE file="PhoneMask.asp"-->

<title>Guest Profile</title>
<style>
	.mainFont		{ font-family:tahoma;font-size:11px }
	.mainTables		{ font-family:tahoma;font-size:11px;border-style:solid;border-width:1px;border-color:black }

	.id				{ font-family:tahoma;font-size:11px;width:70px }
	.detailBody		{ background-color:#eeebbb;font-family:tahoma;font-size:11px }
	.dbPhone		{ background-color:#C9DF86;font-family:tahoma;font-size:11px }
	.dbAddress		{ background-color:#C9DF86;font-family:tahoma;font-size:11px }
	.dbRewards		{ background-color:#C9DF86;font-family:tahoma;font-size:11px }
	.dbCC			{ background-color:#C9DF86;font-family:tahoma;font-size:11px }
	.dbFamily		{ background-color:#C9DF86;font-family:tahoma;font-size:11px }
	.dbDates		{ background-color:#C9DF86;font-family:tahoma;font-size:11px }
	.dbPrefs		{ background-color:#C9DF86;font-family:tahoma;font-size:11px }

	.ext			{ font-family:tahoma;font-size:11px;width:48px }
	.newExt			{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:48px }
	.phonenumber	{ font-family:tahoma;font-size:11px;width:100px }
	.newphonenumber	{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:100px }
	.phonenote		{ font-family:tahoma;font-size:11px;width:325px }
	.newphonenote	{ font-family:tahoma;font-size:11px;width:325px;background-color:#fbffa0 }
	.removebutton	{ font-family:tahoma;font-size:11px;width:60px }
	.newFont		{ font-family:tahoma;font-size:11px;background-color:#fbffa0 }
	.GuestHeader	{ font-family:tahoma;font-size:11px }
	.long			{ font-family:tahoma;font-size:11px;width:181px }
	.dh				{ color:white;background-color:black }
	.PhoneType		{ font-family:tahoma;font-size:11px;width:142px }
	.NewPhoneType	{ font-family:tahoma;font-size:11px;width:142px;background-color:#fbffa0 }
	
	.NewAddressType	{ font-family:tahoma;font-size:11px;width:100px;background-color:#fbffa0 }
	.AddressType	{ font-family:tahoma;font-size:11px;width:100px }
	.newstreet		{ font-family:tahoma;font-size:11px;background-color:#fbffa0 }
	.street			{ font-family:tahoma;font-size:11px }
	.newSuite		{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:48px }
	.Suite			{ font-family:tahoma;font-size:11px;width:48px }
	.newCity		{ font-family:tahoma;font-size:11px;background-color:#fbffa0 }
	.City			{ font-family:tahoma;font-size:11px }
	.newstate		{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:44px }
	.state			{ font-family:tahoma;font-size:11px;width:44px }
	.newZip			{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:48px }
	.Zip			{ font-family:tahoma;font-size:11px;width:48px }
	.newAddrNote	{ font-family:tahoma;font-size:11px;background-color:#fbffa0 }
	.AddrNote		{ font-family:tahoma;font-size:11px }

	.NewRewardsType	{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:142px }
	.RewardsType	{ font-family:tahoma;font-size:11px;width:142px }
	.NewProgName	{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:130px }
	.ProgName		{ font-family:tahoma;font-size:11px;width:130px }
	.NewProgNum		{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:100px }
	.ProgNum		{ font-family:tahoma;font-size:11px;width:100px }
	.NewProgLevel	{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:90px }
	.ProgLevel		{ font-family:tahoma;font-size:11px;width:90px }
	.NewGRNote		{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:180px }
	.GRNote			{ font-family:tahoma;font-size:11px;width:180px }

	.NewCCType		{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:100px }
	.CCType			{ font-family:tahoma;font-size:11px;width:100px }
	.NewCCNumber	{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:144px }
	.CCNumber		{ font-family:tahoma;font-size:11px;width:144px }
	.NewCCExp		{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:66px }
	.CCExp			{ font-family:tahoma;font-size:11px;width:66px }
	.NewGCNote		{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:220px }
	.GCNote			{ font-family:tahoma;font-size:11px;width:220px }

	.NewFMType		{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:100px }
	.FMType			{ font-family:tahoma;font-size:11px;width:100px }
	.NewFMSal		{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:100px }
	.FMSal			{ font-family:tahoma;font-size:11px;width:100px }
	.newFMFirstName	{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:80px }
	.FMFirstName	{ font-family:tahoma;font-size:11px;width:80px }
	.newFMLastName	{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:100px }
	.FMLastName		{ font-family:tahoma;font-size:11px;width:100px }
	.newFMAge		{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:40px }
	.FMAge			{ font-family:tahoma;font-size:11px;width:40px }
	.newFMDOB		{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:80px }
	.FMDOB			{ font-family:tahoma;font-size:11px;width:80px }
	.NewFMNote		{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:126px }
	.FMNote			{ font-family:tahoma;font-size:11px;width:126px }

	.newGINote		{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:174px }
	.GINote			{ font-family:tahoma;font-size:11px;width:174px }
	.newGPrefNote	{ font-family:tahoma;font-size:11px;background-color:#fbffa0;width:518px }
	.GPrefNote		{ font-family:tahoma;font-size:11px;width:518px }

	.abut			{ font-family:tahoma;font-size:11px;width:100px }
	A:hover		{color:red}
	A:active	{color:blue}
	A:visited	{color:blue}
</style>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
var intLineCount = 10000000000;
var booFirstLoad = true;
var nGPPrimID = 0;

function submitme()
{
	if(booFirstLoad)
		booFirstLoad = false;
	else
		window.close();
}

function window_onload() {


	<%
	' Salutation
	dim str
	do until rsSal.EOF
		str = str & "<option value=""" & sq(rsSal("Salutation").Value) & """>"
		str = str & sq(rsSal("Salutation").Value) & "</option>"
		rsSal.MoveNext
	loop
	response.write "document.all('tdSal').innerHTML = '<select class=mainFont id=ddSalutation name=ddSalutation>" & str & "</select>';" & vbcrlf
	'''''''''''''''''
	response.write "document.all('txtGID').value = '" & strGuestID & "';" & vbcrlf
	response.write "document.all('ddSalutation').value = '" & strSalutation & "';" & vbcrlf
	response.write "document.all('txtLastName').value = '" & strLastName & "';" & vbcrlf
	response.write "document.all('txtMiddleName').value = '" & strMiddleName & "';" & vbcrlf
	response.write "document.all('txtFirstName').value = '" & strFirstName & "';" & vbcrlf
	response.write "document.all('txtCompany').value = '" & strCompany & "';" & vbcrlf
	response.write "document.all('txtTitle').value = '" & strTitle & "';" & vbcrlf
	response.write "document.all('txtEMail').value = '" & strEMail & "';" & vbcrlf
	response.write "document.all('txtEMail2').value = '" & strEMail2 & "';" & vbcrlf
	'response.write "document.all('chkSmoking').checked = " & strSmoking & ";" & vbcrlf
	response.write "document.all('txtPMSGuestID').value = '" & strPMSGuestID & "';" & vbcrlf
	response.write "document.all('txtHotelGuestID').value = '" & strHotelGuestID & "';" & vbcrlf
	response.write "document.all('txtGuestNote').value = '" & strGuestNote & "'.replace(/<<crlf>>/gi,'\n');" & vbcrlf
	response.write "FillPhone('txtPrimaryPhone','" & strPrimaryPhone & "');" & vbcrlf
	response.write "document.all('txtPhoneExt').value = '" & strPhoneExt & "';" & vbcrlf
	response.write strScript
	%>
}

function doSave()
{

	window.frmSubmit.action = "GuestProfileProcess.asp"
	window.frmSubmit.target = "frmSubmitFrame"
	window.frmSubmit.method = "post"
	getHotelIDs();
	parent.returnValue = "refresh";
	window.frmSubmit.submit();
}

function getHotelIDs()
{
	var comma = ",", l = document.all("lstAssignedHotels").options.length;
	document.all("txtAssignedHotels").value = ""
	for(var i=0;i<l;i++)
		{
		if(i==(parseInt(l)-1))
			comma = "";
		document.all("txtAssignedHotels").value += document.all("lstAssignedHotels").options(i).value + comma
		}
}

function cmdAssign_onclick()
{
	var lu = document.all("lstUnAssignedHotels"), la = document.all("lstAssignedHotels")
	var si = lu.selectedIndex
	if(si > -1)
		{
		var o = document.createElement("option")
		o.text = lu.options(si).text;
		o.value = lu.options(si).value;
		la.add(o);
		lu.remove(si);
		si -= 1;
		if(si < 0)
			si = 0;
		if(lu.options.length > 0)
			lu.selectedIndex = si;
		}
}

function cmdUnAssign_onclick()
{
	var lu = document.all("lstUnAssignedHotels"), la = document.all("lstAssignedHotels")
	var si = la.selectedIndex
	if(si > -1)
		{
		var o = document.createElement("option")
		o.text = la.options(si).text;
		o.value = la.options(si).value;
		lu.add(o);
		la.remove(si);
		si -= 1;
		if(si < 0)
			si = 0;
		if(la.options.length > 0)
			la.selectedIndex = si;
		}
}

function doNothing()
{}

function removeLine()
{
	var se = window.event.srcElement, booPrimary = false;
	if(se.parentElement.parentElement.parentElement.parentElement.rows.length > 2)
	{
		var id = getGPID(se.id), bypass = false;

		// check for primary checkbox
		var group = se.parentElement.parentElement.id.substr(0,se.parentElement.parentElement.id.indexOf("_"))
		switch(group)
		{
			case "trGP":
				{
				newGroup = "chkGPPrimary"
				break;
				}
			case "trGA":
				{
				newGroup = "chkAddressPrimary"
				break;
				}
			case "trGR":
				{
				newGroup = "Dummy";
				bypass = true;
				break;
				}
			case "trGC":
				{
				newGroup = "chkGCPrimary"
				break;
				}
			case "trGF":
				{
				newGroup = "Dummy"
				bypass = true;
				break;
				}
			case "trGI":
				{
				newGroup = "Dummy"
				bypass = true;
				break;
				}
			case "trGPref":
				{
				newGroup = "Dummy"
				bypass = true;
				break;
				}
		}

		if(!bypass && document.all(newGroup+"_"+id).checked) // && !bypass)
		{
			tag = document.all.tags("INPUT")
			for(var i=0;i<tag.length;i++)
				{
					if(tag[i].id.indexOf(newGroup) > -1 && id != getGPID(tag[i].id))
					{
						tag[i].checked = true;
						break;
					}
				}
		}
				
		// delete the row
		se.parentElement.parentElement.parentElement.parentElement.deleteRow(se.parentElement.parentElement.rowIndex);
	}
	else
		alert("You must have at least one phone number.")
		
}

function getGPID( str )
{
	var pos1 = str.indexOf("_")
	var newstr = str.substr(pos1+1)
	return (newstr)
}

function addLine( str, defaults )
{
	var removeType = null;
	
	intLineCount += 1;
	switch(str)
	{
		case "PhoneType":
			{
			// Phone Type
			var tag = document.all.tags("select")
			for(var j=0;j<tag.length;j++)
			{
				if(tag[j].id.indexOf("selGPPhoneType") > -1)
					{
					o = tag[j]
					break;
					}
			}
			var t = document.all("tblGP");
			var myNewRow = t.insertRow();
			myNewRow.id = "trGP_"+intLineCount;
			var myNewCell = myNewRow.insertCell();
			var spc = "selGPPhoneType_"+intLineCount;
			var e = document.createElement("<select>");
			e.id = spc;
			e.name = spc;
			e.className = "NewPhoneType";
			for(j = 0;j < o.options.length;j++)
				{
				oOption = document.createElement("OPTION");
				e.options.add(oOption);
				e.options[j].value = o.options[j].value;
				e.options[j].text = o.options[j].text;
				if(e.options[j].text == "Business")
					e.options[j].selected = true;
				}
			myNewCell.insertBefore(e);
			
			// Phone Number
			myNewCell = myNewRow.insertCell();
			spc = "txtGPPhoneNumber"+intLineCount;
			CreatePhoneField( spc, 'font-family: Tahoma; font-size: 11', '13px', 100, null, myNewCell, "#fbffa0" );

			// Extention
			myNewCell = myNewRow.insertCell();
			spc = "txtGPExt_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newExt";
			myNewCell.insertBefore(e);

			// Phone Primary
			myNewCell = myNewRow.insertCell();
			myNewCell.align = "center";
			spc = "chkGPPrimary_"+intLineCount;
			e = document.createElement("<input type=checkbox>");
			e.id = spc;
			e.name = spc;
			e.className = "newFont";
			e.onclick = pponclick;
			myNewCell.insertBefore(e);
			
			// Phone Note
			myNewCell = myNewRow.insertCell();
			spc = "txtGuestNote_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newphonenote";
			myNewCell.insertBefore(e);
			
			break;
			}
		case 'AddressType':
			{
			// Address Type
			var tag = document.all.tags("select")
			for(var j=0;j<tag.length;j++)
			{
				if(tag[j].id.indexOf("selATAddressType") > -1)
					{
					o = tag[j]
					break;
					}
			}
			var t = document.all("tblGA");
			var myNewRow = t.insertRow();
			myNewRow.id = "trGA_"+intLineCount;
			var myNewCell = myNewRow.insertCell();
			var spc = "selATAddressType_"+intLineCount;
			var e = document.createElement("<select>");
			e.id = spc;
			e.name = spc;
			e.className = "NewAddressType";
			for(j = 0;j < o.options.length;j++)
				{
				oOption = document.createElement("OPTION");
				e.options.add(oOption);
				e.options[j].value = o.options[j].value;
				e.options[j].text = o.options[j].text;
				if(e.options[j].text == "Home")
					e.options[j].selected = true;
				}
			myNewCell.insertBefore(e);

			// Street
			myNewCell = myNewRow.insertCell();
			spc = "txtGAStreet_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newStreet";
			myNewCell.insertBefore(e);

			// Suite
			myNewCell = myNewRow.insertCell();
			spc = "txtGASuite_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newSuite";
			myNewCell.insertBefore(e);

			// City
			myNewCell = myNewRow.insertCell();
			spc = "txtGACity_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newCity";
			myNewCell.insertBefore(e);

			// State
			/*
			myNewCell = myNewRow.insertCell();
			spc = "txtGAState_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newState";
			myNewCell.insertBefore(e);
			*/

			for(var j=0;j<tag.length;j++)
			{
				if(tag[j].id.indexOf("selGAState") > -1)
					{
					o = tag[j]
					break;
					}
			}
			var myNewCell = myNewRow.insertCell();
			var spc = "selGAState_"+intLineCount;
			var e = document.createElement("<select>");
			e.id = spc;
			e.name = spc;
			e.className = "newState";
			for(j = 0;j < o.options.length;j++)
				{
				oOption = document.createElement("OPTION");
				e.options.add(oOption);
				e.options[j].value = o.options[j].value;
				e.options[j].text = o.options[j].text;
				if(e.options[j].text == "<%=remote.session("CompanyState")%>")
					e.options[j].selected = true;
				}
			myNewCell.insertBefore(e);
			
			
			// Zip
			myNewCell = myNewRow.insertCell();
			spc = "txtGAZip_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newZip";
			myNewCell.insertBefore(e);

			// Address Primary
			myNewCell = myNewRow.insertCell();
			myNewCell.align = "center";
			spc = "chkAddressPrimary_"+intLineCount;
			e = document.createElement("<input type=checkbox id=checkbox1 name=checkbox1>");
			e.id = spc;
			e.name = spc;
			e.className = "newFont";
			e.onclick = pponclick;
			myNewCell.insertBefore(e);
			
			// Note
			myNewCell = myNewRow.insertCell();
			spc = "txtAddrNote_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newAddrNote";
			myNewCell.insertBefore(e);
			
			break;
			}
		case 'RewardsProgram':
			{
			// Rewards Type
			var tag = document.all.tags("select")
			for(var j=0;j<tag.length;j++)
			{
				if(tag[j].id.indexOf("selGRRewardsType") > -1)
					{
					o = tag[j]
					break;
					}
			}
			var t = document.all("tblGR");
			var myNewRow = t.insertRow();
			myNewRow.id = "trGR_"+intLineCount;
			var myNewCell = myNewRow.insertCell();
			var spc = "selGRRewardsType_"+intLineCount;
			var e = document.createElement("<select>");
			e.id = spc;
			e.name = spc;
			e.className = "NewRewardsType";
			for(j = 0;j < o.options.length;j++)
				{
				oOption = document.createElement("OPTION");
				e.options.add(oOption);
				e.options[j].value = o.options[j].value;
				e.options[j].text = o.options[j].text;
				if(e.options[j].text == "Hotel")
					e.options[j].selected = true;
				}
			myNewCell.insertBefore(e);

			// Program Name
			for(var j=0;j<tag.length;j++)
			{
				if(tag[j].id.indexOf("selProgramName") > -1)
					{
					o = tag[j]
					break;
					}
			}
			myNewCell = myNewRow.insertCell();
			var spc = "selGRProgramName_"+intLineCount;
			var e = document.createElement("<select>");
			e.id = spc;
			e.name = spc;
			e.className = "NewProgName";
			for(j = 0;j < o.options.length;j++)
				{
				oOption = document.createElement("OPTION");
				e.options.add(oOption);
				e.options[j].value = o.options[j].value;
				e.options[j].text = o.options[j].text;
				if(e.options[j].text == "American")
					e.options[j].selected = true;
				}
			myNewCell.insertBefore(e);

			// Program Number
			myNewCell = myNewRow.insertCell();
			spc = "txtGRProgNum_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newProgNum";
			myNewCell.insertBefore(e);

			// Program Level
			myNewCell = myNewRow.insertCell();
			spc = "txtGRProgLevel_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newProgLevel";
			myNewCell.insertBefore(e);

			// Note
			myNewCell = myNewRow.insertCell();
			spc = "txtGRNote_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newGRNote";
			myNewCell.insertBefore(e);
			break;
			}

		case 'ChargeType':
			{
			// Charge Type
			if(defaults)
				var a = defaults.split("|");
			
			var tag = document.all.tags("select")
			for(var j=0;j<tag.length;j++)
			{
				if(tag[j].id.indexOf("selCCCCType") > -1)
					{
					o = tag[j]
					break;
					}
			}
			var t = document.all("tblGC");
			var myNewRow = t.insertRow();
			myNewRow.id = "trGC_"+intLineCount;
			var myNewCell = myNewRow.insertCell();
			var spc = "selCCCCType_"+intLineCount;
			var e = document.createElement("<select>");
			e.id = spc;
			e.name = spc;
			e.className = "NewCCType";
			for(j = 0;j < o.options.length;j++)
				{
				oOption = document.createElement("OPTION");
				e.options.add(oOption);
				e.options[j].value = o.options[j].value;
				e.options[j].text = o.options[j].text;
				if(e.options[j].text == "Visa")
					e.options[j].selected = true;
				}
			myNewCell.insertBefore(e);

			// CC Number
			myNewCell = myNewRow.insertCell();
			spc = "txtGCNumber_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newCCNumber";
			myNewCell.insertBefore(e);

			// CC Exp
			myNewCell = myNewRow.insertCell();
			spc = "txtGCExp_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newCCExp";
			myNewCell.insertBefore(e);

			// Zip
			myNewCell = myNewRow.insertCell();
			spc = "txtGCZip_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newZip";
			myNewCell.insertBefore(e);

			// Primary
			myNewCell = myNewRow.insertCell();
			myNewCell.align = "center";
			spc = "chkGCPrimary_"+intLineCount;
			e = document.createElement("<input type=checkbox>");
			e.id = spc;
			e.name = spc;
			e.onclick = pponclick;
			e.className = "newFont";
			myNewCell.insertBefore(e);
			
			// Note
			myNewCell = myNewRow.insertCell();
			spc = "txtGCNote_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newGCNote";
			myNewCell.insertBefore(e);

			break;
			}

		case 'FamilyMembers':
			{
			// Relationship
			var tag = document.all.tags("select")
			for(var j=0;j<tag.length;j++)
			{
				if(tag[j].id.indexOf("selFMType") > -1)
					{
					o = tag[j]
					break;
					}
			}
			var t = document.all("tblGF");
			var myNewRow = t.insertRow();
			myNewRow.id = "trGF_"+intLineCount;
			var myNewCell = myNewRow.insertCell();
			var spc = "selFMType_"+intLineCount;
			var e = document.createElement("<select>");
			e.id = spc;
			e.name = spc;
			e.className = "NewFMType";
			for(j = 0;j < o.options.length;j++)
				{
				oOption = document.createElement("OPTION");
				e.options.add(oOption);
				e.options[j].value = o.options[j].value;
				e.options[j].text = o.options[j].text;
				if(e.options[j].text == "Wife")
					e.options[j].selected = true;
				}
			myNewCell.insertBefore(e);

            // Salutation
			for(var j=0;j<tag.length;j++)
			{
				if(tag[j].id.indexOf("selFMSal") > -1)
					{
					o = tag[j]
					break;
					}
			}
			var myNewCell = myNewRow.insertCell();
			var spc = "selFMSal_"+intLineCount;
			var e = document.createElement("<select>");
			e.id = spc;
			e.name = spc;
			e.className = "NewFMSal";
			for(j = 0;j < o.options.length;j++)
				{
				oOption = document.createElement("OPTION");
				e.options.add(oOption);
				e.options[j].value = o.options[j].value;
				e.options[j].text = o.options[j].text;
				if(e.options[j].text == "Ms.")
					e.options[j].selected = true;
				}
			myNewCell.insertBefore(e);

			// First Name
			myNewCell = myNewRow.insertCell();
			spc = "txtFMFirstName_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newFMFirstName";
			myNewCell.insertBefore(e);

			// Last Name
			myNewCell = myNewRow.insertCell();
			spc = "txtFMLastName_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newFMLastName";
			myNewCell.insertBefore(e);

			// DOB
			myNewCell = myNewRow.insertCell();
			spc = "txtFMDOB_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newFMDOB";
			myNewCell.insertBefore(e);

			// Age
			myNewCell = myNewRow.insertCell();
			spc = "txtFMAge_"+intLineCount;
			e = document.createElement("<input type=text id=text1 name=text1>");
			e.id = spc;
			e.name = spc;
			e.className = "newFMAge";
			myNewCell.insertBefore(e);
			
			// Note
			myNewCell = myNewRow.insertCell();
			spc = "txtFMNote_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newFMNote";
			myNewCell.insertBefore(e);

			break;
			}

		case 'ImportantDates':
			{
			var tag = document.all.tags("select")
			var t = document.all("tblID");
			var myNewRow = t.insertRow();
			myNewRow.id = "trID_"+intLineCount;

            // Date Type
			for(var j=0;j<tag.length;j++)
			{
				if(tag[j].id.indexOf("selFMType") > -1)
					{
					o = tag[j]
					break;
					}
			}
			var myNewCell = myNewRow.insertCell();
			var spc = "selIDType_"+intLineCount;
			var e = document.createElement("<select>");
			e.id = spc;
			e.name = spc;
			e.className = "NewFMType";
			for(j = 0;j < o.options.length;j++)
				{
				oOption = document.createElement("OPTION");
				e.options.add(oOption);
				e.options[j].value = o.options[j].value;
				e.options[j].text = o.options[j].text;
				if(e.options[j].text == "Wife")
					e.options[j].selected = true;
				}
			myNewCell.insertBefore(e);

			// Date
			myNewCell = myNewRow.insertCell();
			spc = "txtIDDate_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newFMDOB";
			myNewCell.insertBefore(e);

			// First Name
			myNewCell = myNewRow.insertCell();
			spc = "txtIDFirstName_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newFMFirstName";
			myNewCell.insertBefore(e);

			// Last Name
			myNewCell = myNewRow.insertCell();
			spc = "txtIDLastName_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newFMLastName";
			myNewCell.insertBefore(e);

			// Relationship
			for(var j=0;j<tag.length;j++)
			{
				if(tag[j].id.indexOf("selGIRelationship") > -1)
					{
					o = tag[j]
					break;
					}
			}
			var myNewCell = myNewRow.insertCell();
			var spc = "selIDRelationship_"+intLineCount;
			var e = document.createElement("<select>");
			e.id = spc;
			e.name = spc;
			e.className = "NewFMType";
			for(j = 0;j < o.options.length;j++)
				{
				oOption = document.createElement("OPTION");
				e.options.add(oOption);
				e.options[j].value = o.options[j].value;
				e.options[j].text = o.options[j].text;
				if(e.options[j].text == "Wife")
					e.options[j].selected = true;
				}
			myNewCell.insertBefore(e);

			// Note
			myNewCell = myNewRow.insertCell();
			spc = "txtIDNote_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newGINote";
			myNewCell.insertBefore(e);

			break;
			}

		case 'Preferences':
			{
			var tag = document.all.tags("select")
			var t = document.all("tblGPref");
			var myNewRow = t.insertRow();
			myNewRow.id = "trGPref_"+intLineCount;

            // Date Type
			for(var j=0;j<tag.length;j++)
			{
				if(tag[j].id.indexOf("selPrefType") > -1)
					{
					o = tag[j]
					break;
					}
			}
			var myNewCell = myNewRow.insertCell();
			var spc = "selPrefType_"+intLineCount;
			var e = document.createElement("<select>");
			e.id = spc;
			e.name = spc;
			e.className = "NewPhoneType";
			for(j = 0;j < o.options.length;j++)
				{
				oOption = document.createElement("OPTION");
				e.options.add(oOption);
				e.options[j].value = o.options[j].value;
				e.options[j].text = o.options[j].text;
				}
			myNewCell.insertBefore(e);
			
			// Note
			myNewCell = myNewRow.insertCell();
			spc = "txtGPrefNote_"+intLineCount;
			e = document.createElement("<input type=text>");
			e.id = spc;
			e.name = spc;
			e.className = "newGPrefNote";
			myNewCell.insertBefore(e);

			break;
			}
	}

	// Remove hyperlink
	myNewCell = myNewRow.insertCell();
	spc = "aRemove_"+intLineCount;
	e = document.createElement("<a>")
	e.className = "mainFont"
	e.id = spc;
	e.name = spc;
	e.innerText = "Remove";
	e.onclick = removeLine
	e.href = "javascript:doNothing()"
	myNewCell.insertBefore(e);
}

function pponclick()
{
	optionCheck(this)
}

function optionCheck( o )
{
	var anychecked = false;
	var tag = document.all.tags("INPUT");
	var sepPos = o.id.indexOf("_");
	var group = o.id.substr(0,sepPos)
	
	for(var i = 0;i < tag.length; i++)
	{
		if(tag[i].id.indexOf(group) > -1)
			if(tag[i].id != o.id)
				tag[i].checked = false;
			else
				if(tag[i].checked)
					anychecked = true;
	}
	
	if(o.id.indexOf("chkGPPrimary_") > -1)
		{
		var id = o.id.substr(sepPos+1);
		FillPhone("txtPrimaryPhone",document.all("txtGPPhoneNumber"+id).value);
		nGPPrimID = id;
		//document.all("txtPrimaryPhone").value = document.all("txtGPPhoneNumber"+o.id.substr(sepPos+1)).value;
		}
		
	if(!anychecked)
		window.event.returnValue = false;
}

function kd()
{
	var se = window.event.srcElement
	var id = se.id.substr(se.id.lastIndexOf("_")+1)
	FillPhone("txtPrimaryPhone",document.all(id).value);
}
//-->
</SCRIPT>

</head>

<body topmargin=2 bottommargin=0 leftmargin=8 rightmargin=0 bgcolor=#F9D568 LANGUAGE=javascript onload="return window_onload()">
<iframe onload=submitme() style=visibility:hidden;display:none id=frmSubmitFrame name=frmSubmitFrame></iframe>
<form id=frmSubmit name=frmSubmit>
<table bgcolor=#F0c568 style=border-style:outset;border-width:2px width=100% cellpadding=0 cellspacing=2 class=GuestHeader>
  <tr>
  <td>
  
	  <table width=100% cellpadding=0 cellspacing=0>
		<tr>
		<td>
			<table style=border-style:outset;border-width:1px; width=100% bgcolor=lightyellow cellpadding=0 cellspacing=2 class=GuestHeader><tr><td>
				<td>GoConcierge ID:</td>
				<td><input tabindex=-1 unselectable=on class=mainFont style=padding-left:4px;background-color:silver;border:ridge;border-width:2px;width:70px type=text id=txtGID name=txtGID></td>
				<td align=right>Salutation:</td>
				<td style=width:70px id=tdSal></td>
				<td align=right>Last Name:</td>
				<td><input style=width:104px class=mainFont type=text id=txtLastName name=txtLastName></td>
				<td align=right>Middle Name:</td>
				<td><input style=width:70px class=mainFont type=text id=txtMiddleName name=txtMiddleName></td>
				<td align=right>First Name:</td>
				<td><input style=width:94px class=mainFont type=text id=txtFirstName name=txtFirstName></td>
			</td></tr></table>
		</tr>
		<tr>
		<td>
			<table cellpadding=0 cellspacing=2 width=100% class=GuestHeader><tr><td>
				<td>Hotel Guest ID:</td>
				<td><input class=id type=text id=txtHotelGuestID name=txtHotelGuestID></td>
				<td align=right>PMS Guest ID:</td>
				<td><input class=id type=text id=txtPMSGuestID name=txtPMSGuestID></td>
				<td style=width:52px align=right>Company:</td>
				<td><input class=long type=text id=txtCompany name=txtCompany></td>
				<td style=width:36px align=right>Title:</td>
				<td><input class=long type=text id=txtTitle name=txtTitle></td>
			</td></tr></table>
		</td>
		</tr>
		<tr>
			<td>
			<table cellpadding=0 cellspacing=2 width=100% class=GuestHeader><tr><td>
				<td>Phone:</td>
				<td>
					<script language=javascript>
							var x = CreatePhoneField( 'txtPrimaryPhone', 'font-family: Tahoma; font-size: 11', '13px', 100, null, null, 'white', true );
					</script>
					<!--input tabindex=-1 unselectable=on style=background-color:silver;width:100px class=mainFont type=text id=txtPrimaryPhone name=txtPrimaryPhone-->
				</td>
				<td align=right width=20px>Ext:</td>
				<td><input type=text id=txtPhoneExt name=txtPhoneExt class=mainFont style=width:60px></td>
				<td width=124px align=right>EMail:</td>
				<td><input class=long type=text id=txtEMail name=txtEMail></td>
				<td align=right>EMail 2:</td>
				<td><input class=long type=text id=txtEMail2 name=txtEMail2></td>
				<!--td align=left><input class=mainFont type=checkbox id=chkSmoking name=chkSmoking></td-->
			</td></tr></table>
			</td>
		</tr>
		<tr>
			<td>
			<table cellpadding=0 cellspacing=2 class=GuestHeader><tr><td>
				<td valign=top>Notes:</td>
	  			<td valign=top>
					<textarea class=mainFont style=width:730px;height:40px; id=txtGuestNote name=txtGuestNote></textarea>
				</td>
			</td></tr></table>
			</td>
		</tr>
	</table>
</td>

</tr>
</table>
<hr>
<div id=divBody style="width:100%;height:<%=bodyheight%>px;overflow:auto">
<table cellpadding=0 cellspacing=0 class=GuestHeader border=0 width="100%">
  <tr class=dbPhone>
    <td>
		<table class=mainTables cellpadding=0 cellspacing=0 width="100%">
			<tr>
				<td>
					<table id=tblGP class=mainFont cellspacing=0 cellpadding=0 width="100%">
						<tr>
							<td class=dh width=125px>&nbsp;Telephone Type</td>
							<td class=dh>&nbsp;Phone</td>
							<td class=dh>&nbsp;Ext.</td>
							<td class=dh width=50px align=center>Primary</td>
							<td class=dh>&nbsp;Notes</td>
							<td class=dh>&nbsp;</td>
						</tr>
						<%
						dim rsGuestPhone, line
						set rsGuestPhone = server.CreateObject("adodb.recordset")
						set rsGuestPhone = cn.Execute("sp_GuestPhone " & gid)
						set rsPT = server.CreateObject("adodb.recordset")
						set rsPT = cn.Execute("select * from tlkpPhoneType order by PhoneType")
						gpEOF = rsGuestPhone.EOF
						line = 1
						if gpEOF then
							gpid = 0
							Response.Write "<tr id=trGP_" & gpid & ">"

							'Phone type...
							Response.Write "<td><select class=NewPhoneType name=selGPPhoneType_" & line & " id=selGPPhoneType_" & line & ">"
							do until rsPT.EOF
								if rsPT.Fields("PhoneType").Value = "Business" then
									strSelected = "selected"
								else
									strSelected = ""
								end if
								Response.Write "<option " & strSelected & " value=" & rsPT.Fields("PhoneTypeID").Value & ">" & rsPT.Fields("PhoneType").Value & "</option>"
								rsPT.MoveNext
							loop
							Response.Write "</select></td>"
							'''''''''''''''
								
							Response.Write "<td><div>"
							Response.Write "<script language=javascript>"
							Response.Write "CreatePhoneField( 'txtGPPhoneNumber" & line & "', 'font-family: Tahoma; font-size: 11', '13px', 100, null, null, '#fbffa0' );" & vbcrlf
							Response.Write "document.all('pm_AreaCode_txtGPPhoneNumber" & line & "').onblur = kd" & vbcrlf
							Response.Write "document.all('pm_Prefix_txtGPPhoneNumber" & line & "').onblur = kd" & vbcrlf
							Response.Write "document.all('pm_Suffix_txtGPPhoneNumber" & line & "').onblur = kd" & vbcrlf
							Response.Write "nGPPrimID = " & line & ";" & vbcrlf
							Response.Write "</script>" & vbcrlf
							'<input class=phonenumber type=text id=txtGPPhoneNumber" & gpid & " value=""" & trim(rsGuestPhone.Fields("PhoneNumber").Value) & """>
							Response.Write "</div></td>"
								
							Response.Write "<td><input class=newExt name=txtGPExt_" & line & " id=txtGPExt_" & line & " type=text></td>"

							strChecked = "checked" 
							Response.Write "<td align=center><input onclick=""optionCheck(this)"" class=newFont name=chkGPPrimary_" & line & " id=chkGPPrimary_" & line & " type=checkbox " & strChecked & "></td>"
								
							Response.Write "<td><input class=newphonenote type=text name=txtGPGuestNote_" & line & " id=txtGPGuestNote_" & line & "></td>"
							Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & gpid & ">Remove</a></td>"
							Response.Write "</tr>" & vbcrlf
						else
							do until rsGuestPhone.EOF
								gpid = rsGuestPhone.Fields("PhoneTypeID").Value
								Response.Write "<tr id=trGP_" & line & ">"

								'Phone type...
								Response.Write "<td><select class=PhoneType name=selGPPhoneType_" & line & " id=selGPPhoneType_" & line & ">"
								rsPT.MoveFirst
								do until rsPT.EOF
									if rsGuestPhone.Fields("PhoneTypeID").Value = rsPT.Fields("PhoneTypeID").Value then
										strSelected = "selected"
									else
										strSelected = ""
									end if
									Response.Write "<option " & strSelected & " value=" & rsPT.Fields("PhoneTypeID").Value & ">" & rsPT.Fields("PhoneType").Value & "</option>"
									rsPT.MoveNext
								loop
								Response.Write "</select></td>"
								'''''''''''''''
								
								Response.Write "<td><div>"
								Response.Write "<script language=javascript>"
								Response.Write "CreatePhoneField( 'txtGPPhoneNumber" & line & "', 'font-family: Tahoma; font-size: 11', '13px', 100, null, null, 'white' );"
								Response.Write "FillPhone('txtGPPhoneNumber" & line & "','" & trim(rsGuestPhone.Fields("PhoneNumber").Value) & "');"
								Response.Write "document.all('pm_AreaCode_txtGPPhoneNumber" & line & "').onblur = kd" & vbcrlf
								Response.Write "document.all('pm_Prefix_txtGPPhoneNumber" & line & "').onblur = kd" & vbcrlf
								Response.Write "document.all('pm_Suffix_txtGPPhoneNumber" & line & "').onblur = kd" & vbcrlf
								if rsGuestPhone.Fields("PhonePrimary").Value then 
									strChecked = "checked"
									Response.Write "nGPPrimID = " & line & ";" & vbcrlf
								else 
									strChecked = ""
								end if
								Response.Write "</script>" & vbcrlf
								'<input class=phonenumber type=text id=txtGPPhoneNumber" & gpid & " value=""" & trim(rsGuestPhone.Fields("PhoneNumber").Value) & """>
								Response.Write "</div></td>"
								
								Response.Write "<td><input class=Ext name=txtGPExt_" & line & " id=txtGPExt_" & line & " type=text value=""" & rsGuestPhone.Fields("PhoneExt").Value & """></td>"

								Response.Write "<td align=center><input onclick=""optionCheck(this)"" class=mainFont name=chkGPPrimary_" & line & " id=chkGPPrimary_" & line & " type=checkbox " & strChecked & "></td>"
								
								Response.Write "<td><input class=phonenote type=text name=txtGuestNote_" & line & " id=txtGuestNote_" & line & " value=""" & trim(rsGuestPhone.Fields("PhoneNote").Value) & """></td>"
								
								Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & line & ">Remove</a></td>"
								Response.Write "</tr>" & vbcrlf
								rsGuestPhone.MoveNext
								line = line + 1
							loop
							line = 1
						end if
						rsPT.Close
						set rsPT = nothing
						rsGuestPhone.Close
						set rsGuestPhone = nothing%>
					</table>
				</td>
				<td width=46px height=100% valign=top align=center>
					<table width=100% height=100% cellpadding=0 cellspacing=0 class=mainFont valign=top>
						<tr>
							<td class=dh>&nbsp;</td>
						</tr>
						<tr>
							<td align=center height="100%"><a href=javascript:addLine('PhoneType')>Add</a></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
    </td>
  </tr>
  <tr><td>&nbsp;</td></tr>
  <tr class=dbAddress>
    <td>
		<table class=mainTables cellpadding=0 cellspacing=0 width="100%">
			<tr>
				<td>
					<table id=tblGA class=mainFont cellspacing=0 cellpadding=0 width="100%">
						<tr>
							<td class=dh>&nbsp;Address Type</td>
							<td class=dh>&nbsp;Street</td>
							<td class=dh>&nbsp;Suite</td>
							<td class=dh>&nbsp;City</td>
							<td class=dh>&nbsp;State</td>
							<td class=dh>&nbsp;Zip</td>
							<td class=dh>Primary</td>
							<td class=dh>&nbsp;Notes</td>
							<td class=dh>&nbsp;</td>
						</tr>
						<%
						set rs = cn.Execute("sp_GuestAddress " & gid)
						set rsLookup = server.CreateObject("adodb.recordset")
						set rsLookup = cn.Execute("select * from tlkpAddressType order by AddressType")
						set rsState = server.CreateObject("adodb.recordset")
						set rsState = cn.Execute("select * from tlkpState order by Abbreviation")
						booEOF = rs.EOF
						if booEOF then
							intID = 0
							Response.Write "<tr id=trGA_" & intID & ">"

							'Address type...
							Response.Write "<td><select class=NewAddressType name=selATAddressType_" & intID & " id=selATAddressType_" & intID & ">"
							do until rsLookup.EOF
								if trim(rsLookup.Fields("AddressType").Value) = "Home" then
									strSelected = "selected"
								else
									strSelected = ""
								end if
								Response.Write "<option " & strSelected & " value=" & rsLookup.Fields("AddressTypeID").Value & ">" & rsLookup.Fields("AddressType").Value & "</option>"
								rsLookup.MoveNext
							loop
							Response.Write "</select></td>"
							'''''''''''''''
								
							Response.Write "<td><input class=newstreet type=text name=txtGAStreet_" & intID & " id=txtGAStreet_" & intID & "></td>"
							Response.Write "<td><input class=newSuite type=text name=txtGASuite_" & intID & " id=txtGASuite_" & intID & "></td>"
							Response.Write "<td><input class=newCity type=text name=txtGACity_" & intID & " id=txtGACity_" & intID & "></td>"

							Response.Write "<td><select name=selGAState_" & intID & " id=selGAState_" & intID & " class=newState>"
							do until rsState.EOF
								if remote.session("CompanyState") = rsState.Fields("Abbreviation").Value then
									strSelected = "selected"
								else
									strSelected = ""
								end if
								Response.Write "<option " & strSelected & " value=" & rsState.Fields("Abbreviation").Value & ">" & rsState.Fields("Abbreviation").Value & "</option>"
								rsState.MoveNext
							loop
							Response.Write "</select></td>"
							'Response.Write "<td><input class=newState type=text name=txtGAState_" & intID & " id=txtGAState_" & intID & "></td>"

							Response.Write "<td><input class=newZip type=text name=txtGAZip_" & intID & " id=txtGAZip_" & intID & "></td>"
								
							strChecked = "checked" 
							Response.Write "<td align=center><input onclick=""optionCheck(this)"" class=newFont name=chkAddressPrimary_" & intID & " id=chkAddressPrimary_" & intID & " type=checkbox " & strChecked & "></td>"
								
							Response.Write "<td><input class=newAddrNote type=text name=txtAddrNote_" & intID & " id=txtAddrNote_" & intID & "></td>"
							Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & intID & ">Remove</a></td>"
							Response.Write "</tr>" & vbcrlf
						else
							do until rs.EOF
								intID = rs.Fields("GuestAddressID").Value
								Response.Write "<tr id=trGA_" & intID & ">"

								'Address type...
								Response.Write "<td><select class=AddressType name=selATAddressType_" & intID & " id=selATAddressType_" & intID & ">"
								rsLookup.MoveFirst
								do until rsLookup.EOF
									if rs.Fields("AddressTypeID").Value = rsLookup.Fields("AddressTypeID").Value then
										strSelected = "selected"
									else
										strSelected = ""
									end if
									Response.Write "<option " & strSelected & " value=" & rsLookup.Fields("AddressTypeID").Value & ">" & rsLookup.Fields("AddressType").Value & "</option>"
									rsLookup.MoveNext
								loop
								Response.Write "</select></td>"
								'''''''''''''''
								
								Response.Write "<td><input class=street type=text name=txtGAStreet_" & intID & " id=txtGAStreet_" & intID & " value=""" & rs.Fields("Address").Value & """></td>"
								Response.Write "<td><input class=Suite type=text name=txtGASuite_" & intID & " id=txtGASuite_" & intID & " value=""" & rs.Fields("Suite").Value & """></td>"
								Response.Write "<td><input class=City type=text name=txtGACity_" & intID & " id=txtGACity_" & intID & " value=""" & rs.Fields("City").Value & """></td>"
								
								Response.Write "<td><select name=selGAState_" & intID & " id=selGAState_" & intID & " class=State>"
								rsState.MoveFirst
								do until rsState.EOF
									if rs.Fields("State").Value = rsState.Fields("Abbreviation").Value then
										strSelected = "selected"
									else
										strSelected = ""
									end if
									Response.Write "<option " & strSelected & " value=" & rsState.Fields("Abbreviation").Value & ">" & rsState.Fields("Abbreviation").Value & "</option>"
									rsState.MoveNext
								loop
								Response.Write "</select></td>"
								' <input class=State type=text name=txtGAState_" & intID & " id=txtGAState_" & intID & " value=""" & rs.Fields("State").Value & """></td>"
								
								Response.Write "<td><input class=Zip type=text name=txtGAZip_" & intID & " id=txtGAZip_" & intID & " value=""" & rs.Fields("Zip").Value & """></td>"

								if rs.Fields("AddressPrimary").Value then 
									strChecked = "checked" 
								else 
									strChecked = ""
								end if
								Response.Write "<td align=center><input onclick=""optionCheck(this)"" class=mainFont name=chkAddressPrimary_" & intID & " id=chkAddressPrimary_" & intID & " type=checkbox " & strChecked & "></td>"
								
								Response.Write "<td><input class=addrnote type=text name=txtAddrNote_" & intID & " id=txtAddrNote_" & intID & " value=""" & trim(rs.Fields("Note").Value) & """></td>"
								
								Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & intID & ">Remove</a></td>"
								Response.Write "</tr>" & vbcrlf
								rs.MoveNext
							loop
						end if
						rsState.Close
						set rsState = nothing
						rsLookup.Close
						set rsLookup = nothing%>
					</table>
				</td>
				<td width=36px height=100% valign=top align=center>
					<table width=100% height=100% cellpadding=0 cellspacing=0 class=mainFont valign=top>
						<tr>
							<td class=dh>&nbsp;</td>
						</tr>
						<tr>
							<td align=center height="100%"><a href=javascript:addLine('AddressType')>Add</a></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
    </td>
  </tr>
  <tr><td>&nbsp;</td></tr>
  <tr class=dbRewards>
    <td>
		<table class=mainTables cellpadding=0 cellspacing=0 width="100%">
			<tr>
				<td>
					<table id=tblGR class=mainFont cellspacing=0 cellpadding=0 width="100%">
						<tr>
							<td class=dh>&nbsp;Rewards Program</td>
							<td class=dh>&nbsp;Program Name</td>
							<td class=dh>&nbsp;Number</td>
							<td class=dh>&nbsp;Level</td>
							<td class=dh>&nbsp;Notes</td>
							<td class=dh>&nbsp;</td>
						</tr>
						<%
						set rs = server.CreateObject("adodb.recordset")
						set rs = cn.Execute("sp_GuestRewards " & gid)
						set rsLookup = server.CreateObject("adodb.recordset")
						set rsLookup = cn.Execute("select * from tlkpRewardsType order by RewardsType")
						set rsLookup2 = server.CreateObject("adodb.recordset")
						set rsLookup2 = cn.Execute("select * from tlkpProgram order by Program")
						booEOF = rs.EOF
						if booEOF then
							intID = 0
							Response.Write "<tr id=trGR_" & intID & ">"

							'Rewards type...
							Response.Write "<td><select class=NewRewardsType name=selGRRewardsType_" & intID & " id=selGRRewardsType_" & intID & ">"
							do until rsLookup.EOF
								if trim(rsLookup.Fields("RewardsType").Value) = "Hotel" then
									strSelected = "selected"
								else
									strSelected = ""
								end if
								Response.Write "<option " & strSelected & " value=" & rsLookup.Fields("RewardsTypeID").Value & ">" & rsLookup.Fields("RewardsType").Value & "</option>"
								rsLookup.MoveNext
							loop
							Response.Write "</select></td>"
							'''''''''''''''
								
							'Program Name...
							Response.Write "<td><select class=NewProgName name=selGRProgramName_" & intID & " id=selGRProgramName_" & intID & ">"
							do until rsLookup2.EOF
								if trim(rsLookup2.Fields("ProgramID").Value) = "American" then
									strSelected = "selected"
								else
									strSelected = ""
								end if
								Response.Write "<option " & strSelected & " value=" & rsLookup2.Fields("ProgramID").Value & ">" & rsLookup2.Fields("Program").Value & "</option>"
								rsLookup2.MoveNext
							loop
							Response.Write "</select></td>"
							'''''''''''''''
							'Response.Write "<td><input class=newProgName type=text id=txtGRProg_" & intID & "></td>"
							
							Response.Write "<td><input class=newProgNum type=text name=txtGRProgNum_" & intID & " id=txtGRProgNum_" & intID & "></td>"
							Response.Write "<td><input class=newProgLevel type=text name=txtGRProgLevel_" & intID & " id=txtGRProgLevel_" & intID & "></td>"
							Response.Write "<td><input class=newGRNote type=text name=txtGRNote_" & intID & " id=txtGRNote_" & intID & "></td>"
							Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & intID & ">Remove</a></td>"
							Response.Write "</tr>" & vbcrlf
						else
							do until rs.EOF
								intID = rs.Fields("RewardsTypeID").Value
								Response.Write "<tr id=trGR_" & intID & ">"

								'Rewards type...
								Response.Write "<td><select class=RewardsType name=selGRRewardsType_" & intID & " id=selGRRewardsType_" & intID & ">"
								rsLookup.MoveFirst
								do until rsLookup.EOF
									if rs.Fields("RewardsTypeID").Value = rsLookup.Fields("RewardsTypeID").Value then
										strSelected = "selected"
									else
										strSelected = ""
									end if
									Response.Write "<option " & strSelected & " value=" & rsLookup.Fields("RewardsTypeID").Value & ">" & rsLookup.Fields("RewardsType").Value & "</option>"
									rsLookup.MoveNext
								loop
								Response.Write "</select></td>"
								'''''''''''''''
								
								'Program Name...
								Response.Write "<td><select class=ProgName name=selGRProgramName_" & intID & " id=selGRProgramName_" & intID & ">"
								rsLookup2.movefirst
								do until rsLookup2.EOF
									if trim(rsLookup2.Fields("ProgramID").Value) = rs.fields("ProgramID").value then
										strSelected = "selected"
									else
										strSelected = ""
									end if
									Response.Write "<option " & strSelected & " value=" & rsLookup2.Fields("ProgramID").Value & ">" & rsLookup2.Fields("Program").Value & "</option>"
									rsLookup2.MoveNext
								loop
								Response.Write "</select></td>"
								'''''''''''''''
								'Response.Write "<td><input class=ProgName type=text id=txtGRProg_" & intID & " value=""" & rs.fields("ProgramName").value & """></td>"

								Response.Write "<td><input class=ProgNum type=text name=txtGRProgNum_" & intID & " id=txtGRProgNum_" & intID & " value=""" & rs.fields("ProgramNumber").value & """></td>"
								Response.Write "<td><input class=ProgLevel type=text name=txtGRProgLevel_" & intID & " id=txtGRProgLevel_" & intID & " value=""" & rs.fields("ProgramLevel").value & """></td>"
								if rs.Fields("ShowOnSearch").Value then 
									strChecked = "checked" 
								else 
									strChecked = ""
								end if
								Response.Write "<td><input class=GRNote type=text name=txtGRNote_" & intID & " id=txtGRNote_" & intID & " value=""" & rs.fields("Note").value & """></td>"
								Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & intID & ">Remove</a></td>"
								Response.Write "</tr>" & vbcrlf
								rs.MoveNext
							loop
						end if
						rsLookup.Close
						set rsLookup = nothing%>
					</table>
				</td>
				<td width=36px height=100% valign=top align=center>
					<table width=100% height=100% cellpadding=0 cellspacing=0 class=mainFont valign=top>
						<tr>
							<td class=dh>&nbsp;</td>
						</tr>
						<tr>
							<td align=center height="100%"><a href=javascript:addLine('RewardsProgram')>Add</a></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
    </td>
  </tr>
  <tr><td>&nbsp;</td></tr>
  <tr class=dbCC>
    <td>
		<table class=mainTables cellpadding=0 cellspacing=0 width="100%">
			<tr>
				<td>
					<table id=tblGC class=mainFont cellspacing=0 cellpadding=0 width="100%">
						<tr>
							<td class=dh>&nbsp;Credit Card Type</td>
							<td class=dh>&nbsp;Number</td>
							<td class=dh>&nbsp;Expiration</td>
							<td class=dh>&nbsp;Zip Code</td>
							<td class=dh align=center width=60px>Primary</td>
							<td class=dh>&nbsp;Notes</td>
							<td class=dh>&nbsp;</td>
						</tr>
						<%
						set rs = server.CreateObject("adodb.recordset")
						set rs = cn.Execute("sp_GuestCharge " & gid)
						set rsLookup = server.CreateObject("adodb.recordset")
						set rsLookup = cn.Execute("select * from tlkpChargeType order by ChargeType")
						booEOF = rs.EOF
						if booEOF then
							intID = 0
							Response.Write "<tr id=trGC_" & intID & ">"

							Response.Write "<td><select class=NewCCType name=selCCCCType_" & intID & " id=selCCCCType_" & intID & ">"
							do until rsLookup.EOF
								if Request.QueryString("cct") <> "" then
									booEval = (rsLookup.Fields("ChargeTypeID").Value = cint(Request.QueryString("cct")))
								else
									booEval = (trim(rsLookup.Fields("ChargeType").Value) = "Visa")
								end if
								
								if booEval then
									strSelected = "selected"
								else
									strSelected = ""
								end if
								Response.Write "<option " & strSelected & " value=""" & rsLookup.Fields("ChargeTypeID").Value & """>" & rsLookup.Fields("ChargeType").Value & "</option>"
								rsLookup.MoveNext
							loop
							Response.Write "</select></td>"
							'''''''''''''''
								
							Response.Write "<td><input value=""" & Request.QueryString("cn") & """ class=newCCNumber type=text name=txtGCNumber_" & intID & " id=txtGCNumber_" & intID & "></td>"
							Response.Write "<td><input value=""" & Request.QueryString("exp") & """ class=newCCExp type=text name=txtGCExp_" & intID & " id=txtGCExp_" & intID & "></td>"
							Response.Write "<td><input class=newZip type=text name=txtGCZip_" & intID & " id=txtGCZip_" & intID & "></td>"
							Response.Write "<td align=center><input onclick=optionCheck(this) class=newFont name=chkGCPrimary_" & intID & " id=chkGCPrimary_" & intID & " type=checkbox checked></td>"
							Response.Write "<td><input class=newGCNote type=text name=txtGCNote_" & intID & " id=txtGCNote_" & intID & "></td>"
							Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & intID & ">Remove</a></td>"
							Response.Write "</tr>" & vbcrlf
						else
							do until rs.EOF
								intID = rs.Fields("ChargeTypeID").Value
								Response.Write "<tr id=trGC_" & intID & ">"

								Response.Write "<td><select class=CCType name=selCCCCType_" & intID & " id=selCCCCType_" & intID & ">"
								rsLookup.MoveFirst
								do until rsLookup.EOF
									if rs.Fields("ChargeTypeID").Value = rsLookup.Fields("ChargeTypeID").Value then
										strSelected = "selected"
									else
										strSelected = ""
									end if
									Response.Write "<option " & strSelected & " value=" & rsLookup.Fields("ChargeTypeID").Value & ">" & rsLookup.Fields("ChargeType").Value & "</option>"
									rsLookup.MoveNext
								loop
								Response.Write "</select></td>"
								'''''''''''''''
								
								Response.Write "<td><input class=CCNumber type=text name=txtGCNumber_" & intID & " id=txtGCNumber_" & intID & " value=""" & rs.fields("ChargeNumber").value & """></td>"
								Response.Write "<td><input class=CCExp type=text name=txtGCExp_" & intID & " id=txtGCExp_" & intID & " value=""" & rs.fields("Expiration").value & """></td>"
								Response.Write "<td><input class=Zip type=text name=txtGCZip_" & intID & " id=txtGCZip_" & intID & " value=""" & rs.fields("ZipCode").value & """></td>"
								if rs.Fields("ChargePrimary").Value then 
									strChecked = "checked" 
								else 
									strChecked = ""
								end if
								Response.Write "<td align=center><input onclick=optionCheck(this) class=mainFont name=chkGCPrimary_" & intID & " id=chkGCPrimary_" & intID & " type=checkbox " & strChecked & "></td>"
								Response.Write "<td><input class=GCNote type=text name=txtGCNote_" & intID & " id=txtGCNote_" & intID & " value=""" & rs.fields("Note").value & """></td>"
								Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & intID & ">Remove</a></td>"
								Response.Write "</tr>" & vbcrlf
								rs.MoveNext
							loop
						end if
						rsLookup.Close
						set rsLookup = nothing%>
					</table>
				</td>
				<td width=36px height=100% valign=top align=center>
					<table width=100% height=100% cellpadding=0 cellspacing=0 class=mainFont valign=top>
						<tr>
							<td class=dh>&nbsp;</td>
						</tr>
						<tr>
							<td align=center height="100%"><a href=javascript:addLine('ChargeType')>Add</a></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
    </td>
  </tr>
  <tr><td>&nbsp;</td></tr>
  <tr class=dbFamily>
    <td>
		<table class=mainTables cellpadding=0 cellspacing=0 width="100%">
			<tr>
				<td>
					<table id=tblGF class=mainFont cellspacing=0 cellpadding=0 width="100%">
						<tr>
							<td class=dh>&nbsp;Family Members</td>
							<td class=dh>&nbsp;Salutation</td>
							<td class=dh>&nbsp;First Name</td>
							<td class=dh>&nbsp;Last Name</td>
							<td class=dh>&nbsp;Birthday</td>
							<td class=dh>&nbsp;Age</td>
							<td class=dh>&nbsp;Notes</td>
							<td class=dh>&nbsp;</td>
						</tr>
						<%
						set rs = server.CreateObject("adodb.recordset")
						set rs = cn.Execute("sp_GuestFamily " & gid)
						set rsLookup = server.CreateObject("adodb.recordset")
						set rsLookup = cn.Execute("select * from tlkpRelationship order by Relationship")
						set rsLookup2 = server.CreateObject("adodb.recordset")
						set rsLookup2 = cn.Execute("select * from tblSalutations order by Salutation")
						booEOF = rs.EOF
						if booEOF then
							intID = 0
							Response.Write "<tr id=trGR_" & intID & ">"

							Response.Write "<td><select class=NewFMType name=selFMType_" & intID & " id=selFMType_" & intID & ">"
							do until rsLookup.EOF
								if trim(rsLookup.Fields("Relationship").Value) = "Wife" then
									strSelected = "selected"
								else
									strSelected = ""
								end if
								Response.Write "<option " & strSelected & " value=" & rsLookup.Fields("RelationshipID").Value & ">" & rsLookup.Fields("Relationship").Value & "</option>"
								rsLookup.MoveNext
							loop
							Response.Write "</select></td>"
							'''''''''''''''
								
							' salutation
							Response.Write "<td><select class=NewFMSal name=selFMSal_" & intID & " id=selFMSal_" & intID & ">"
							rsLookup2.movefirst
							do until rsLookup2.EOF
								if trim(rsLookup2.Fields("Salutation").Value) = "Mr." then
									strSelected = "selected"
								else
									strSelected = ""
								end if
								Response.Write "<option " & strSelected & " value=" & rsLookup2.Fields("Salutation").Value & ">" & rsLookup2.Fields("Salutation").Value & "</option>"
								rsLookup2.MoveNext
							loop
							Response.Write "</select></td>"
							'''''''''''''''

							Response.Write "<td><input class=newFMFirstName type=text name=txtFMFirstName_" & intID & " id=txtFMFirstName_" & intID & "></td>"
							Response.Write "<td><input class=newFMLastName type=text name=txtFMLastName_" & intID & " id=txtFMLastName_" & intID & "></td>"
							Response.Write "<td><input class=newFMDOB type=text name=txtFMDOB_" & intID & " id=txtFMDOB_" & intID & "></td>"
							Response.Write "<td><input class=newFMAge type=text name=txtFMAge_" & intID & " id=txtFMAge_" & intID & "></td>"
							Response.Write "<td><input class=newFMNote type=text name=txtFMNote_" & intID & " id=txtFMNote_" & intID & "></td>"
							Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & intID & ">Remove</a></td>"
							Response.Write "</tr>" & vbcrlf
						else
							do until rs.EOF
								intID = rs.Fields("GuestFamilyMemberID").Value
								Response.Write "<tr id=trGF_" & intID & ">"

								Response.Write "<td><select class=FMType name=selFMType_" & intID & " id=selFMType_" & intID & ">"
								rsLookup.MoveFirst
								do until rsLookup.EOF
									if rs.Fields("RelationshipID").Value = rsLookup.Fields("RelationshipID").Value then
										strSelected = "selected"
									else
										strSelected = ""
									end if
									Response.Write "<option " & strSelected & " value=" & rsLookup.Fields("RelationshipID").Value & ">" & rsLookup.Fields("Relationship").Value & "</option>"
									rsLookup.MoveNext
								loop
								Response.Write "</select></td>"
								'''''''''''''''
								
								' salutation
								Response.Write "<td><select class=FMSal name=selFMSal_" & intID & " id=selFMSal_" & intID & ">"
								rsLookup2.movefirst
								do until rsLookup2.EOF
									if trim(rsLookup2.Fields("Salutation").Value) = trim(rs.Fields("Salutation").value) then
										strSelected = "selected"
									else
										strSelected = ""
									end if
									Response.Write "<option " & strSelected & " value=" & rsLookup2.Fields("Salutation").Value & ">" & rsLookup2.Fields("Salutation").Value & "</option>"
									rsLookup2.MoveNext
								loop
								Response.Write "</select></td>"
								'''''''''''''''

								Response.Write "<td><input class=FMFirstName type=text name=txtFMFirstName_" & intID & " id=txtFMFirstName_" & intID & " value=""" & rs.fields("FirstName").value & """></td>"
								Response.Write "<td><input class=FMLastName type=text name=txtFMLastName_" & intID & " id=txtFMLastName_" & intID & " value=""" & rs.fields("LastName").value & """></td>"
								Response.Write "<td><input class=FMDOB type=text name=txtFMDOB_" & intID & " id=txtFMDOB_" & intID & " value=""" & rs.fields("Birthdate").value & """></td>"
								Response.Write "<td><input class=FMAge type=text name=txtFMAge_" & intID & " id=txtFMAge_" & intID & " value=""" & rs.fields("Age").value & """></td>"
								Response.Write "<td><input class=FMNote type=text name=txtFMNote_" & intID & " id=txtFMNote_" & intID & " value=""" & rs.fields("Note").value & """></td>"

								Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & intID & ">Remove</a></td>"
								Response.Write "</tr>" & vbcrlf
								rs.MoveNext
							loop
						end if
						rsLookup.Close
						set rsLookup = nothing%>
					</table>
				</td>
				<td width=36px height=100% valign=top align=center>
					<table width=100% height=100% cellpadding=0 cellspacing=0 class=mainFont valign=top>
						<tr>
							<td class=dh>&nbsp;</td>
						</tr>
						<tr>
							<td align=center height="100%"><a href=javascript:addLine('FamilyMembers')>Add</a></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
    </td>
  </tr>
  <tr><td>&nbsp;</td></tr>
  <tr class=dbDates>
    <td>
		<table class=mainTables cellpadding=0 cellspacing=0 width="100%">
			<tr>
				<td>
					<table id=tblID class=mainFont cellspacing=0 cellpadding=0 width="100%">
						<tr>
							<td class=dh>&nbsp;Important Dates</td>
							<td class=dh>&nbsp;Date</td>
							<td class=dh>&nbsp;First Name</td>
							<td class=dh>&nbsp;Last Name</td>
							<td class=dh>&nbsp;Relation</td>
							<td class=dh>&nbsp;Notes</td>
							<td class=dh>&nbsp;</td>
						</tr>
						<%
						set rs = server.CreateObject("adodb.recordset")
						set rs = cn.Execute("sp_GuestImportantDates " & gid)
						set rsLookup = server.CreateObject("adodb.recordset")
						set rsLookup = cn.Execute("select * from tlkpDateType order by DateType")
						set rsLookup2 = server.CreateObject("adodb.recordset")
						set rsLookup2 = cn.Execute("select * from tlkpRelationship order by Relationship")
						booEOF = rs.EOF
						if booEOF then
							intID = 0
							Response.Write "<tr id=trID_" & intID & ">"

							' Date type
							Response.Write "<td><select class=NewFMType name=selIDType_" & intID & " id=selIDType_" & intID & ">"
							do until rsLookup.EOF
								if trim(rsLookup.Fields("DateType").Value) = "Birthday" then
									strSelected = "selected"
								else
									strSelected = ""
								end if
								Response.Write "<option " & strSelected & " value=" & rsLookup.Fields("DateTypeID").Value & ">" & rsLookup.Fields("DateType").Value & "</option>"
								rsLookup.MoveNext
							loop
							Response.Write "</select></td>"
							'''''''''''''''
								
							Response.Write "<td><input class=newFMDOB type=text name=txtIDDate_" & intID & " id=txtIDDate_" & intID & "></td>"
							Response.Write "<td><input class=newFMFirstName type=text name=txtIDFirstName_" & intID & " id=txtIDFirstName_" & intID & "></td>"
							Response.Write "<td><input class=newFMLastName type=text name=txtIDLastName_" & intID & " id=txtIDLastName_" & intID & "></td>"

							' Relationship
							Response.Write "<td><select class=NewFMType name=selIDRelationship_" & intID & " id=selIDRelationship_" & intID & ">"
							do until rsLookup2.EOF
								if trim(rsLookup2.Fields("Relationship").Value) = "Wife" then
									strSelected = "selected"
								else
									strSelected = ""
								end if
								Response.Write "<option " & strSelected & " value=" & rsLookup2.Fields("RelationshipID").Value & ">" & rsLookup2.Fields("Relationship").Value & "</option>"
								rsLookup2.MoveNext
							loop
							Response.Write "</select></td>"
							'''''''''''''''

							Response.Write "<td><input class=newGINote type=text name=txtIDNote_" & intID & " id=txtIDNote_" & intID & "></td>"

							Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & intID & ">Remove</a></td>"
							Response.Write "</tr>" & vbcrlf
						else
							do until rs.EOF
								intID = rs.Fields("GuestImportantDateID").Value
								Response.Write "<tr id=trID_" & intID & ">"

								' Date type
								Response.Write "<td><select class=FMType name=selIDType_" & intID & " id=selIDType_" & intID & ">"
								rsLookup.movefirst
								do until rsLookup.EOF
									if rsLookup.Fields("DateTypeID").Value = rs.fields("DateTypeID").value then
										strSelected = "selected"
									else
										strSelected = ""
									end if
									Response.Write "<option " & strSelected & " value=" & rsLookup.Fields("DateTypeID").Value & ">" & rsLookup.Fields("DateType").Value & "</option>"
									rsLookup.MoveNext
								loop
								Response.Write "</select></td>"
								'''''''''''''''
									
								Response.Write "<td><input class=FMDOB type=text name=txtIDDate_" & intID & " id=txtIDDate_" & intID & " value=""" & rs.fields("DateTypeDate").value & """></td>"
								Response.Write "<td><input class=FMFirstName type=text name=txtIDFirstName_" & intID & " id=txtIDFirstName_" & intID & " value=""" & rs.fields("FirstName").value & """></td>"
								Response.Write "<td><input class=FMLastName type=text name=txtIDLastName_" & intID & " id=txtIDLastName_" & intID & " value=""" & rs.fields("LastName").value & """></td>"

								' Relationship
								Response.Write "<td><select class=FMType name=selIDRelationship_" & intID & " id=selIDRelationship_" & intID & ">"
								rsLookup2.movefirst
								do until rsLookup2.EOF
									if rsLookup2.Fields("RelationshipID").Value = rs.fields("RelationshipID").value then
										strSelected = "selected"
									else
										strSelected = ""
									end if
									Response.Write "<option " & strSelected & " value=" & rsLookup2.Fields("RelationshipID").Value & ">" & rsLookup2.Fields("Relationship").Value & "</option>"
									rsLookup2.MoveNext
								loop
								Response.Write "</select></td>"
								'''''''''''''''

								Response.Write "<td><input class=GINote type=text name=txtIDNote_" & intID & " id=txtIDNote_" & intID & "></td>"

								Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & intID & ">Remove</a></td>"
								Response.Write "</tr>" & vbcrlf
								rs.MoveNext
							loop
						end if
						rsLookup.Close
						set rsLookup = nothing%>
					</table>
				</td>
				<td width=36px height=100% valign=top align=center>
					<table width=100% height=100% cellpadding=0 cellspacing=0 class=mainFont valign=top>
						<tr>
							<td class=dh>&nbsp;</td>
						</tr>
						<tr>
							<td align=center height="100%"><a href=javascript:addLine('ImportantDates')>Add</a></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
    </td>
  </tr>
  <tr><td>&nbsp;</td></tr>
  <tr class=dbPrefs>
    <td>
		<table class=mainTables cellpadding=0 cellspacing=0 width="100%">
			<tr>
				<td>
					<table id=tblGPref class=mainFont cellspacing=0 cellpadding=0 width="100%">
						<tr>
							<td class=dh>&nbsp;Preferences</td>
							<td class=dh>&nbsp;Notes</td>
							<td class=dh>&nbsp;</td>
						</tr>
						<%
						set rs = server.CreateObject("adodb.recordset")
						set rs = cn.Execute("sp_GuestPreferences " & gid)
						set rsLookup = server.CreateObject("adodb.recordset")
						set rsLookup = cn.Execute("select * from tlkpPreference order by Preference")
						booEOF = rs.EOF
						if booEOF then
							intID = 0
							Response.Write "<tr id=trGPref_" & intID & ">"

							' Preference
							Response.Write "<td><select class=NewPhoneType name=selPrefType_" & intID & " id=selPrefType_" & intID & ">"
							do until rsLookup.EOF
								Response.Write "<option value=" & rsLookup.Fields("PreferenceID").Value & ">" & rsLookup.Fields("Preference").Value & "</option>"
								rsLookup.MoveNext
							loop
							Response.Write "</select></td>"
							'''''''''''''''
								
							Response.Write "<td><input class=newGPrefNote type=text name=txtGPrefNote_" & intID & " id=txtGPrefNote_" & intID & "></td>"

							Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & intID & ">Remove</a></td>"
							Response.Write "</tr>" & vbcrlf
						else
							do until rs.EOF
								intID = rs.Fields("GuestPreferenceID").Value
								Response.Write "<tr id=trGPref_" & intID & ">"

								' Preference
								Response.Write "<td><select class=PhoneType name=selPrefType_" & intID & " id=selPrefType_" & intID & ">"
								rsLookup.movefirst
								do until rsLookup.EOF
									if rsLookup.fields("PreferenceID").value = rs.fields("PreferenceID").value then
										strSelected = "selected"
									else
										strSelected = ""
									end if
									Response.Write "<option " & strSelected & " value=" & rsLookup.Fields("PreferenceID").Value & ">" & rsLookup.Fields("Preference").Value & "</option>"
									rsLookup.MoveNext
								loop
								Response.Write "</select></td>"
								'''''''''''''''
									
								Response.Write "<td><input class=GPrefNote type=text name=txtGPrefNote_" & intID & " id=txtGPrefNote_" & intID & " value=""" & rs.fields("Note").value & """></td>"

								Response.Write "<td><a href=javascript:doNothing() onclick=javascript:removeLine() class=mainFont id=aRemove_" & intID & ">Remove</a></td>"
								Response.Write "</tr>" & vbcrlf
								rs.MoveNext
							loop
						end if
						rsLookup.Close
						set rsLookup = nothing%>
					</table>
				</td>
				<td width=36px height=100% valign=top align=center>
					<table width=100% height=100% cellpadding=0 cellspacing=0 class=mainFont valign=top>
						<tr>
							<td class=dh>&nbsp;</td>
						</tr>
						<tr>
							<td align=center height="100%"><a href=javascript:addLine('Preferences')>Add</a></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
    </td>
  </tr>
  <%
  if booSU then
	strDisplay = "normal"
  else
	strDisplay = "none"
  end if
  %>
  <tr><td>&nbsp;</td></tr>
  <tr style="display:<%=strDisplay%>" class=dbPrefs>
    <td>
		<table class=mainTables cellpadding=0 cellspacing=0 width="100%">
			<tr>
				<td>
					<table id=tblGH class=mainFont cellspacing=0 cellpadding=0 width="100%">
						<tr>
							<td class=dh>Hotels (Super Users Only)</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td width=100% align=center>
					<table>
						<tr>
							<td>
								<select size=2 ondblclick=cmdAssign_onclick() style="height:400px;width:300px" name=lstUnAssignedHotels id=lstUnAssignedHotels>
								<%
								set rs = cn.Execute("sp_GuestHotel " & gid & ", 0, " & Remote.Session("FloatingUser_UserID"))
								do until rs.eof
									Response.Write "<option value=" & rs.fields("CompanyID").value & ">" & rs.fields("CompanyName").value & "</option>"
									rs.movenext
								loop
								%>
								</select>
							</td>
							<td valign=middle>
								<input type=hidden id=txtAssignedHotels name=txtAssignedHotels>
								<input class=abut id=cmdAssign onclick=cmdAssign_onclick() type=button value="Assign >>">
								<br>
								<input class=abut id=cmdUnAssign onclick=cmdUnAssign_onclick() type=button value="<< Unassign">
							</td>
							<td>
								<select size=2 ondblclick=cmdUnAssign_onclick() style="height:400px;width:300px" name=lstAssignedHotels id=lstAssignedHotels>
								<%
								set rs = cn.Execute("sp_GuestHotel " & gid & ", 1," & Remote.Session("FloatingUser_UserID"))
								do until rs.eof
									Response.Write "<option value=" & rs.fields("CompanyID").value & ">" & rs.fields("CompanyName").value & "</option>"
									rs.movenext
								loop
								%>
								</select>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
    </td>
  </tr>
</table>
</div>
</form>

</body>

</html>
<%
rs.Close        
set rs = nothing

rsSal.Close
set rsSal = nothing

cn.Close
set cn = nothing

function sq(str)
	if isnull(str) then
		sq = ""
	else
		sq = replace(str,"'","''")
	end if
end function


function escapeJS(str)
	if isnull(str) then
		escapeJS = ""
	else
		escapeJS = replace(str,"'","\'")
	end if
end function

%>
