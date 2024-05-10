<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))

cid = remote.session("CompanyID")
%>
<!--#INCLUDE file="checkuser.asp"-->
<!--#INCLUDE file="PhoneMask.asp"-->
<!--#INCLUDE file="include/vbFunc.asp"-->

<html>
<head>

<title>Hotel Setup</title>

<style>
	<!--	
		.MyFont			{ font-family: Tahoma; font-size: 11 }
	-->
</style>

<script LANGUAGE="JAVASCRIPT">
<!--
function ValidateInput() {
	var booRetVal = true;
	if (document.form1.txtHotelName.value.length == 0) {
		alert("Please enter a valid Hotel Name.");
		booRetVal = false;
	}

	if (booRetVal && document.form1.txtHotelAddress1.value.length == 0) {
		alert("Please enter a valid Hotel Address.");
		booRetVal = false;
	}

	if (booRetVal && document.form1.txtHotelCity.value.length == 0) {
		alert("Please enter a valid Hotel City.");
		booRetVal = false;
	}
	
	if (booRetVal && document.form1.chkSameAsHotelAddress.checked == false) {
	   if (document.form1.txtDirectionsAddress1.value.length == 0) {
		alert("Please enter a valid Directions Address.");
		booRetVal = false;
	   }
	   
	   if (booRetVal && document.form1.txtDirectionsCity.value.length == 0) {
		alert("Please enter a valid Directions City.");
		booRetVal = false;
	   }
	} 

	/*
	if (document.form1.txtEmail.value.length == 0) {
		alert("Please enter a valid E-mail address.");
		return false;
	}
	
	if (document.form1.txtWebPage.value.length == 0) {
		alert("Please enter a valid Web Page address.");
		return false;
	}
	*/

	if (booRetVal && document.form1.txtLateMinutes.value.length == 0) {
		alert("Please enter a valid time for late minutes.");
		booRetVal = false;
	}

	if (booRetVal && isNaN(document.form1.txtLateMinutes.value)) {
		alert("Please enter a valid time for late minutes.");
		booRetVal = false;
	}

	if (booRetVal && document.form1.txtTopMargin.value.length == 0) {
		alert("Please enter a valid value for top margin.");
		booRetVal = false;
	}

	if (booRetVal && isNaN(document.form1.txtTopMargin.value)) {
		alert("Please enter a valid value for top margin.");
		booRetVal = false;
	}

	if (booRetVal && document.form1.txtBottomMargin.value.length == 0) {
		alert("Please enter a valid value for bottom margin.");
		booRetVal = false;
	}

	if (booRetVal && isNaN(document.form1.txtBottomMargin.value)) {
		alert("Please enter a valid value for bottom margin.");
		booRetVal = false;
	}
	return (booRetVal);

}

function validateNumeric()
{
	var k = window.event.keyCode;
	if( !((k > 45 && k < 58) || (k > 95 && k < 106) || k == 37 || k == 39) )
		window.event.returnValue = false;
}

function validLen( obj, intLen )
	{
		if(obj.value.length > intLen)
			{
				obj.value = obj.value.substr(0,intLen)
				alert("This field takes a maximum of "+intLen+" characters.");
			}
	}
//-->
</script>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Sub window_onload()
	x = onload()
	SwitchPage(1)
	call refreshMasterDepartmentList()
	call refreshAssignedDepartmentList()
End sub

function SetRouteColor()
	if document.all("chkSameAsHotelAddress").checked = false then
		document.all("txtDirectionsAddress1").style.backgroundColor = "Yellow"
		document.all("txtDirectionsCity").style.backgroundColor = "Yellow"
		document.all("cboDirectionsState").style.backgroundColor = "Yellow"
		document.all("txtDirectionsPostalCode").style.backgroundColor = "Yellow"
		document.all("txtHotelAddress1").style.backgroundColor = "White"
		document.all("txtHotelCity").style.backgroundColor = "White"
		document.all("cboHotelState").style.backgroundColor = "White"
		document.all("txtHotelPostalCode").style.backgroundColor = "White"
	else
		document.all("txtHotelAddress1").style.backgroundColor = "Yellow"
		document.all("txtHotelCity").style.backgroundColor = "Yellow"
		document.all("cboHotelState").style.backgroundColor = "Yellow"
		document.all("txtHotelPostalCode").style.backgroundColor = "Yellow"
		document.all("txtDirectionsAddress1").style.backgroundColor = "White"
		document.all("txtDirectionsCity").style.backgroundColor = "White"
		document.all("cboDirectionsState").style.backgroundColor = "White"
		document.all("txtDirectionsPostalCode").style.backgroundColor = "White"
	end if
end function
function CheckSelectedAddress()
	if document.all("chkSameAsHotelAddress").checked = false then
		CheckAddress document.all("txtDirectionsAddress1").value, document.all("txtDirectionsCity").value, document.all("cboDirectionsState").value, document.all("txtDirectionsPostalCode").value
	else
		CheckAddress document.all("txtHotelAddress1").value, document.all("txtHotelCity").value, document.all("cboHotelState").value, document.all("txtHotelPostalCode").value
	end if
end function

function getText( prompt, title, defaultValue )
	getText = inputbox(prompt,title,defaultValue)
end function
</script>

</head>

<body id=bdy class="myFont" bgcolor="menu" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" link="black" vlink="black" alink="black">
<!--#include file = "Header.inc" ---> 
<%
	Set cnSQL = Server.CreateObject("ADODB.Connection")
	Set rsCompany = Server.CreateObject("ADODB.Recordset")
	Set rsGroups = Server.CreateObject("ADODB.Recordset")
	Set rsCalView = Server.CreateObject("ADODB.Recordset")
	
	cnSQL.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")

	'Set rsCompany = cnSQL.Execute("SELECT * from tblCompany Where CompanyID=" & cid)
	
	'Response.Write Request.QueryString("CompanyID") 
	'Response.End
	if len(trim(Request.QueryString("CompanyID"))) > 0 then
		if Request.QueryString("CompanyID") = "0" then
			dim cmd, fso, rssu
			set fso = CreateObject("Scripting.FileSystemObject")
			set cmd = server.CreateObject("ADODB.command")
			set rssu = server.CreateObject("ADODB.recordset")
			cmd.ActiveConnection = cnSQL
			cmd.CommandText = "sp_AddCompany"
			cmd.CommandType = adCmdStoredProc
			set pCompanyID = cmd.CreateParameter("@CompanyID",adInteger,adParamOutput)
			cmd.Parameters.Append pCompanyID
			cmd.Execute
			strCompanyID = cmd.Parameters("@CompanyID").Value
			'Response.Write "CompanyID: " & strCompanyID
			cnSQL.Execute "update tblCompany set LogoLocation = '" & strCompanyID & "_Letterhead.jpg', ScreenLogoLocation = '" & strCompanyID & "_Logo.jpg' where CompanyID=" & strCompanyID
			'set rssu = cnSQL.Execute("SELECT UserID FROM tblUser WHERE SuperUser = 1")
			' This is taken care of in the trigger...
			'cnSQL.Execute "insert tblCompanyUser (CompanyID, UserID, Admin) values (" & strCompanyID & ", " & rssu("UserID") & ", 1)"
			'
			'rssu.Close
			'set rssu = nothing
			set cmd = nothing
			
			'Response.Write Application("ENV_PATH") & "ClientUploads\GoConciergeNet.jpg" & "<br>"
			'Response.Write Application("ENV_PATH") & "ClientUploads\" & strCompanyID & "_LetterHead.jpg"
						
			fso.CopyFile Application("ENV_PATH") & "ClientUploads\DefaultLetterhead.jpg", Application("ENV_PATH") & "ClientUploads\" & strCompanyID & "_LetterHead.jpg"
			fso.CopyFile Application("ENV_PATH") & "ClientUploads\DefaultLogo.jpg", Application("ENV_PATH") & "ClientUploads\" & strCompanyID & "_Logo.jpg"
			set fso = nothing
		else
			strCompanyID = trim(Request.QueryString("CompanyID"))
		end if
	else
		strCompanyID = cid
	end if
	
	'Response.Write "CompanyID: " & strCompanyID

	Set rsCompany = cnSQL.Execute("sp_HotelSetup " & strCompanyID)
	Set rsGroups = cnSQL.Execute("sp_tblGroups")
	
	set rsHL = Server.CreateObject ("Adodb.recordset")
	rsHL.Open "Select * from tblLocation where LocationID=" & rsCompany("LocationID"), cnSQL
	
	set rsCalView = cnsql.Execute("select * from vwCalView where CompanyID = " & strCompanyID & " or CompanyID = 0 order by Name")
	strCalView = "window.form1.cmbCalView.length++;window.form1.cmbCalView(window.form1.cmbCalView.length-1).value='';window.form1.cmbCalView(window.form1.cmbCalView.length-1).text = '(None)';"
	do until rsCalView.EOF
		if rsCalView.Fields("CalViewID").Value = rsCompany.Fields("DefaultCalView").Value then
			selected = "window.form1.cmbCalView.value = " & rsCalView.Fields("CalViewID").Value & ";"
		else
			selected = ""
		end if
		strCalView = strCalView & "window.form1.cmbCalView.length++;window.form1.cmbCalView(window.form1.cmbCalView.length-1).value = " & rsCalView.Fields("CalViewID").Value & ";window.form1.cmbCalView(window.form1.cmbCalView.length-1).text = '" & trim(rsCalView.Fields("Name").Value) & " (" & trim(rsCalView.Fields("MenuName").Value) & ")';" & selected
		rsCalView.MoveNext
	loop
	'Response.Write strCalView

    ' Second check: Make sure that the new record was inserted well
    If (rsCompany.EOF And rsCompany.BOF) Then
		Response.Write "Database Error:  Inserting Company Record Query failed.  Please contact System Administrator." & "<BR>"
	else
		if Request.QueryString("Mode") = "A" then
			strQS = "?CompanyID=" & strCompanyID
		else
			strQS = ""
		end if
		'Response.Write Request.QueryString("Mode") & ":" & strQS
		'Response.End

		if isnull(rsCompany.fields("ItinFontName").value) then
			strItinFontName = "Arial"
			strItinFontSize = 8
		else
			strItinFontName = rsCompany.fields("ItinFontName").value
			strItinFontSize = rsCompany.fields("ItinFontSize").value
		end if

		strSearchIDDisabled = ""
		if rsCompany.Fields("UseGuestProfile").Value then
			strUGPChecked = "checked"
		else
			strUGPChecked = ""
			'strSearchIDDisabled = "disabled"
		end if
%>		

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function updateItinFont(fontName, fontSize)
{
	document.all("selItinFontName").style.fontFamily = fontName;
	document.all("selItinFontName").style.fontSize = parseInt(fontSize)+2;
	document.all("selItinFontSize").style.fontName = fontName;
	document.all("selItinFontSize").style.fontSize = parseInt(fontSize)+2;
}

function setItinFontVals(fontName, fontSize)
{
	for(var i=0;i<document.all("selItinFontName").length;i++)
		if(document.all("selItinFontName").options(i).text == fontName)
			{
			document.all("selItinFontName").selectedIndex = i;
			break;
			}
	for(i=0;i<document.all("selItinFontSize").length;i++)
		if(document.all("selItinFontSize").options(i).text == fontSize)
			{
			document.all("selItinFontSize").selectedIndex = i;
			break;
			}
}

function fileLetterhead_onchange() {
	document.images("imgLetterheadBitmap",0).src = window.form1.fileLetterhead.value;
}

function fileScreen_onchange() {
	document.images("imgScreenLogoBitmap",0).src = window.form1.fileScreen.value;
}

//-->

function SwitchPage(x)
{
	switch(x)
		{
		case 1: 
			document.all("divMain").style.display = "block";
			document.all("divLogo").style.display = "none";
			document.all("divDetail").style.display = "none";
			document.all("divDisclaimer").style.display = "none";
			document.all("divDepartments").style.display = "none";
			document.all("divPreferences").style.display = "none";
			break;
		case 3:
			document.all("divMain").style.display = "none";
			document.all("divLogo").style.display = "block";
			document.all("divDetail").style.display = "none";
			document.all("divDisclaimer").style.display = "none";
			document.all("divDepartments").style.display = "none";
			document.all("divPreferences").style.display = "none";
			break;
		case 2:	
			document.all("divMain").style.display = "none";
			document.all("divLogo").style.display = "none";
			document.all("divDetail").style.display = "block";
			document.all("divDisclaimer").style.display = "none";
			document.all("divDepartments").style.display = "none";
			document.all("divPreferences").style.display = "none";
			break;
		case 4:
			document.all("divMain").style.display = "none";
			document.all("divLogo").style.display = "none";
			document.all("divDetail").style.display = "none";
			document.all("divDisclaimer").style.display = "block";
			document.all("divDepartments").style.display = "none";
			document.all("divPreferences").style.display = "none";
			break;
		case 5:
			document.all("divMain").style.display = "none";
			document.all("divLogo").style.display = "none";
			document.all("divDetail").style.display = "none";
			document.all("divDisclaimer").style.display = "none";
			document.all("divDepartments").style.display = "block";
			document.all("divPreferences").style.display = "none";
			break;
		case 6:
			document.all("divMain").style.display = "none";
			document.all("divLogo").style.display = "none";
			document.all("divDetail").style.display = "none";
			document.all("divDisclaimer").style.display = "none";
			document.all("divDepartments").style.display = "none";
			document.all("divPreferences").style.display = "block";
			break;
			
		}
}

function colorize(o)
{
try {
	document.all.tdLogoBGColor.style.backgroundColor=o.value
} catch(e){}
}
function cmdAdd_onclick() {
	x = getText("Please enter new department name:","Add Master Department","")
	if(x)
	{
		var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
		xmlHttp.open("POST","AddDepartment.asp?t="+x,false)
		xmlHttp.send();
		xmlHttp = null;
		refreshMasterDepartmentList();
	}
}

var curDeptID = '';
var curDeptSelect = '';

function SelectDepartment(d)
{
	curDeptSelect = d;
	curDeptID = d.options(d.selectedIndex).value;
}
function cmdEdit_onclick() {

	if(curDeptID!='')
		{
			x = window.showModalDialog("EditDescription.asp?Mode=Edit&Table=tblDepartment&IDFieldName=DepartmentID&ID="+ curDeptID +"&DescriptionFieldName=DepartmentName","","center: yes; dialogheight: 175px; dialogwidth: 440px; status: no; scroll: no;")
			
			if(x != curDeptSelect.options(curDeptSelect.selectedIndex).text)
			{
				
				refreshMasterDepartmentList();
			}
		}
		else
			alert ('You must select a department first.');
}

function cmdDelete_onclick() {
	var v = document.all("lstDepartment").value;
	if(v)
		{
		var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
		xmlHttp.open("POST","DeleteDepartment.asp?did="+v,false)
		xmlHttp.send();
		var x = xmlHttp.responseText;
		xmlHttp = null;
		if(x == 'OK')
			refreshMasterDepartmentList();
		else
			alert("Can't delete this department.  A task exists under this department.  Delete it, then try again.")
		}
}

function refreshMasterDepartmentList()
{
	var o = document.all("lstDepartment")
	o.innerHTML = "";
	var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
	xmlHttp.open("POST","GetDepartments.asp?cid=<%=strCompanyID%>&Mode=All",false)
	xmlHttp.send();
	var x = xmlHttp.responseText;
	xmlHttp = null;
	if(x)
	{
		var a = x.split("|")
		var b;
		for(var i=0;i<a.length;i++)
			{
			b = a[i].split("~")
			o.length++
			o[i].value = b[0];
			o[i].text = b[1];
			}
	}
}
function refreshAssignedDepartmentList()
{
	var o = document.all("lstDepartmentAssigned")
	o.innerHTML = "";
	var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
	xmlHttp.open("POST","GetDepartments.asp?cid=<%=strCompanyID%>&Mode=Assigned",false)
	xmlHttp.send();
	var x = xmlHttp.responseText;
	xmlHttp = null;
	if(x)
	{
		var a = x.split("|")
		var b;
		for(var i=0;i<a.length;i++)
			{
			b = a[i].split("~")
			o.length++
			o[i].value = b[0];
			o[i].text = b[1];
			}
	}
}

function refreshBoth()
{
	refreshMasterDepartmentList();
	refreshAssignedDepartmentList();
}

function cmdAssignDepartment_onclick()
{
	var o = document.all("lstDepartment")
	if(o.value)
	{
		var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
		xmlHttp.open("POST","AssignDepartment.asp?mode=a&cid=<%=strCompanyID%>&did="+o.value,false)
		xmlHttp.send();
		var x = xmlHttp.responseText;
		xmlHttp = null;
		refreshBoth();
	}
}

function cmdUnAssignDepartment_onclick()
{
	var o = document.all("lstDepartmentAssigned")
	if(o.value)
	{
		var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
		xmlHttp.open("POST","AssignDepartment.asp?mode=d&cid=<%=strCompanyID%>&did="+o.value,false)
		xmlHttp.send();
		var x = xmlHttp.responseText;
		xmlHttp = null;
		refreshBoth();
	}
}
</SCRIPT>
<br>

<p style="font-family:tahoma; font-size: 11px">&nbsp;&nbsp;&nbsp;&nbsp;CompanyID:&nbsp;<input style="width: 60px; font-family:tahoma; font-size: 11px; background-color: silver; border-style: outset; padding-left: 4px" type=text disabled value="<%=strCompanyID%>" id=text1 name=text1>
<table border=2 cellpadding="0" cellspacing="0" width="100%" class="myFont">
<tr>
	<td colspan=5>
		<input value="Main" type=button style="width:150px" id=cmdMainDiv onclick="SwitchPage(1)">
		<input value="Detail" type=button style="width:150px" id=cmdLogoDiv onclick="SwitchPage(2)">
		<input value="Logo" type=button style="width:150px" id=cmdDetailDiv onclick="SwitchPage(3)">
		<input value="Disclaimer" type=button style="width:150px" id=cmdDisclaimer onclick="SwitchPage(4)">
		<input value="Departments" type=button style="width:150px" id=cmdDepartments onclick="SwitchPage(5)">
		<input value="Preferences" type=button style="width:150px" id=cmdPrefs onclick="SwitchPage(6)">
	</td>
</tr>

<tr>
<td>

<!-- Logo STart -->

<div id="divLogo"  style="display:none;height:500px;">
<table width="780" border=1 bordercolordark=Gray class="myFont">

	
	<form METHOD="post" ENCTYPE="multipart/form-data" ACTION="HotelSetupConfirm.asp<%=strQS%>" id="form1" name="form1" ONSUBMIT="return ValidateInput()">

	<tr>
		<td valign="top" align="center" bgcolor="black" width="50%">
			<font face="Tahoma" size="1" color="white">Letterhead Logo
		</td>
		<td valign="top" align="center" bgcolor="black"width="50%">
			<font face="Tahoma" size="1" color="white">Screen Logo
		</td>
	</tr>
	<tr>
		<td  height="80px" valign="baseline" align="center" width="50%">
			
				<img align="center" valign="middle" name="imgLetterheadBitmap" ID="imgLetterheadBitmap" src="ClientUploads/<%=rsCompany("LogoLocation")%>">&nbsp;&nbsp;
				<input id="fileLetterhead" TYPE="FILE" SIZE="0" NAME="FILE1" style="FONT-SIZE: xx-small; HEIGHT: 22px; WIDTH: 100px;" onchange="return fileLetterhead_onchange()">
		</td>
		<td style="backGround-color:<%=rsCompany("LogoBGColor")%>" id=tdLogoBGColor name=tdLogoBGColor valign="baseline" align="center" width="50%">
			
				<img align="center"  valign="middle" name="imgScreenLogoBitmap" ID="imgScreenLogoBitmap" src="ClientUploads/<%=rsCompany("ScreenLogoLocation")%>">&nbsp;&nbsp;
				<input id="fileScreen" SIZE="0" TYPE="FILE" NAME="FILE2" style="FONT-SIZE: xx-small; HEIGHT: 22px; WIDTH: 100px" LANGUAGE=javascript onchange="return fileScreen_onchange()"><br>
				Logo Background Color:&nbsp;<input onkeyup="colorize(this);" type=text id=txtLogoBGColor name=txtLogoBGColor value="<%=trim(rsCompany("LogoBGColor"))%>">
		</td>
	</tr>
</table>

</div>

<!-- Logo End -->

</td>
</tr>
<tr>
<td>
<%if rsCompany.Fields("SameAsHotel").Value then
	strChecked = "CHECKED"
  else
	strChecked = ""
  end if
%>

<table cellpadding="10" class="myFont">




<tr>
<td align="center">
<div id=divMain   style="display:none;height:500px">

<table class="myFont" border="0" width="780" cellpadding="0" cellspacing="0" align="center">
<tr>
<td><strong>Hotel Name:</strong></td>
<td colspan="3&quot;"><input id="txtHotelName" class="myFont" name="txtHotelName" style="WIDTH: 591px" value="<%=rsCompany.Fields("CompanyName")%>"></td>
</tr>
<tr>
<td>Contact:</td>
<td class="myFont">First Name:<input value="<%=rsCompany("ContactFirstName")%>" style="width:50px" name="txtContactFirst" id="txtContactFirst" class="myFont">&nbsp;&nbsp;Last Name:<input style="width:70px" name="txtContactLast" id="txtContactLast" value="<%=rsCompany("ContactLastName")%>" class="myFont"></td>
<td><u>Directions Address</u><input language="vbscript" onclick="SetRouteColor()" class="myFont" id="chkSameAsHotelAddress" name="chkSameAsHotelAddress" type="checkbox" <%=strChecked%>>Same as Hotel Address</td>
</tr>

<tr>
<td>Address1:</td>
<td><input class="myFont" id="txtHotelAddress1" name="txtHotelAddress1" style="WIDTH: 193px" value="<%=rsCompany.Fields("Address1")%>"></td>
<td>Address1:</td> 
<td><input class="myFont" id="txtDirectionsAddress1" name="txtDirectionsAddress1" style="WIDTH: 193px" value="<%=rsCompany.Fields("DirectionsAddress1")%>"></td>
</tr>

<tr>
<td>Address2:</td> 
<td><input class="myFont" id="txtHotelAddress2" name="txtHotelAddress2" style="WIDTH: 192px" value="<%=rsCompany.Fields("Address2")%>"></td>
<td>Address2:</td>
<td><input class="myFont" id="txtDirectionsAddress2" name="txtDirectionsAddress2" style="WIDTH: 192px" value="<%=rsCompany.Fields("DirectionsAddress2")%>"></td>
</tr>


<tr>
<td>City:</td>
<td><input class="myFont" id="txtHotelCity" name="txtHotelCity" style="WIDTH: 192px" value="<%=rsCompany.Fields("City")%>"></td>
<td>City:</td>
<td><input class="myFont" id="txtDirectionsCity" name="txtDirectionsCity" style="WIDTH: 192px" value="<%=rsCompany.Fields("DirectionsCity")%>"></td>
</tr>

<tr>
<td>State:</td>
<td>
<select class="myFont" id="cboHotelState" name="cboHotelState" style="WIDTH: 45px"> 
  <%
	Set rsStates = Server.CreateObject("ADODB.Recordset")
	Set rsStates = cnSQL.Execute("SELECT Abbreviation from tlkpState")

	Do While Not rsStates.EOF
	    if (rsStates.Fields("Abbreviation") = rsCompany.Fields("State")) then
			Response.Write "<OPTION selected value=" & rsStates.Fields("Abbreviation") & ">" & rsStates.Fields("Abbreviation") & "</Option>"
	    else
			Response.Write "<OPTION value=" & rsStates.Fields("Abbreviation") & ">" & rsStates.Fields("Abbreviation") & "</Option>"
		end if

		rsStates.MoveNext
	Loop
  %>
</select>
&nbsp;&nbsp;Country:&nbsp;
<select class="myFont" id="cboHotelCountry" name="cboHotelCountry" style="WIDTH: 90px"> 
<% 
	
	Select Case UCASE(TRIM(rsCompany("Country")))
	
		Case "USA" : strC1 = " selected "
		Case "MEXICO" : strC2 = " selected "
		Case "CANADA" : strC3 = " selected "
		Case else 
			strC1 = " selected "
	End Select
		
		
%>

<option <%=strC1%>>USA</option>
<option <%=strC2%>>Mexico</option>
<option <%=strC3%>>Canada</option>
</select>

</td>

<td>State:</td>
<td>
<select class="myFont" id="cboDirectionsState" name="cboDirectionsState" style="WIDTH: 45px"> 
  <%
	Set rsDirectionsStates = Server.CreateObject("ADODB.Recordset")
	Set rsDirectionsStates = cnSQL.Execute("SELECT Abbreviation from tlkpState")
  
	Do While Not rsDirectionsStates.EOF
	    if (rsDirectionsStates.Fields("Abbreviation") = rsCompany.Fields("State")) then
			Response.Write "<OPTION selected value=" & rsDirectionsStates.Fields("Abbreviation") & ">" & rsDirectionsStates.Fields("Abbreviation") & "</Option>"
	    else
			Response.Write "<OPTION value=" & rsDirectionsStates.Fields("Abbreviation") & ">" & rsDirectionsStates.Fields("Abbreviation") & "</Option>"
		end if

		rsDirectionsStates.MoveNext
	Loop
  %>
</select>
&nbsp;&nbsp;Country:&nbsp;
<select class="myFont" id="cboDirectionsCountry" name="cboDirectionsCountry" style="WIDTH: 90px"> 
<% 
	
	Select Case UCASE(TRIM(rsCompany("DirectionsCountry")))
	
		Case "USA" : strC1 = " selected "
		Case "MEXICO" : strC2 = " selected "
		Case "CANADA" : strC3 = " selected "
		Case else 
			strC1 = " selected "
	End Select
		
		
%>

<option <%=strC1%>>USA</option>
<option <%=strC2%>>Mexico</option>
<option <%=strC3%>>Canada</option>
</select>
</td>
</tr>

<tr>
<td>Postal Code:</td>
<td><input class="myFont" id="txtHotelPostalCode" name="txtHotelPostalCode" style="WIDTH: 88px" value="<%=rsCompany.Fields("PostalCode")%>"></td>
<td>Postal Code:</td>
<td><input class="myFont" id="txtDirectionsPostalCode" name="txtDirectionsPostalCode" style="WIDTH: 88px" value="<%=rsCompany.Fields("DirectionsPostalCode")%>"></td>
</tr>

<tr>
<td>Phone:</td> 
<td>
	<!-- <script language="JavaScript1.2">
		CreatePhoneField ( "txtPhone", "font-family: Tahoma; font-size: 11", "13px", 192 );
		FillPhone ( "txtPhone","<%'=rsCompany.Fields("Phone")%>" );
	</script> -->
	<input class="myFont" id="txtPhone" name="txtPhone" value="<%=rsCompany.Fields("Phone")%>">
</td>
<td>Fax:</td>
<td>
	<!--<script language="JavaScript1.2">
		CreatePhoneField ( "txtFax", "font-family: Tahoma; font-size: 11", "13px", 192 );
		FillPhone ( "txtFax","<%=rsCompany.Fields("Fax")%>" );
	</script> -->
	<input class="myFont" id="txtFax" name="txtFax" value="<%=rsCompany.Fields("Fax")%>">
</td>
</tr>

<tr>

<td>Email:</td> 
<td><input class="myFont" style="WIDTH: 192px" id="txtEmail" name="txtEmail" value="<%=rsCompany.Fields("EMail")%>"></td>
<td>Web Page:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;http://</td> 
<td><input class="myFont" style="WIDTH: 192px" id="txtWebPage" name="txtWebPage" value="<%=rsCompany.Fields("WebPage")%>"></td> 
</tr>

<tr>
<td>Letterhead Default:</td> 
<td><select class="myFont" id="cboLetterheadDefault" name="cboLetterheadDefault" style="WIDTH: 47px"> 
  <% if (rsCompany.Fields("UseCompanyLetterHead") = true) then %>
		<option selected value="1">Yes</option>
		<option value="0">No</option>
  <% else %>
		<option value="1">Yes</option>
		<option selected value="0">No</option>
  <% end if %>
</select>
</td>


<td>Task is late after:</td>
<td><input class="myFont" id="txtLateMinutes" name="txtLateMinutes" style="WIDTH: 43px" value="<%=rsCompany.Fields("LateTaskTime")%>"> minutes</td>
</tr>
<tr>
	<td>Guest Services Ext:</td> 
	<td><input onkeydown=validateNumeric() class="myFont" id="txtGuestServicesExtension" name="txtGuestServicesExtension" value="<%=rsCompany.Fields("GuestServicesExtension")%>"></td>
	<td>Backup Printing Interval:</td>
	<td><input class="myFont" id="txtBackupInterval" name="txtBackupInterval" style="WIDTH: 43px" value="<%=rsCompany.Fields("BackupInterval")%>"> minutes</td>
</tr>
<tr>
<td>Top Letterhead Margin: &nbsp;</td>
<td><input class="myFont" id="txtTopMargin" name="txtTopMargin" style="WIDTH: 55px" value="<%=rsCompany.Fields("LetterheadMargin")%>"> inches </td>
<td>Bottom Letterhead Margin: </td>
<td><input class="myFont" id="txtBottomMargin" name="txtBottomMargin" style="WIDTH: 55px" value="<%=rsCompany.Fields("LetterfootMargin")%>"> inches</td>
<td>
</tr>
<tr>
	<td>Use Guest Profile:</td>
	<td>
		<input oonclick="javascript:document.all('selSearchType').disabled=!this.checked" type=checkbox id=chkUseGuestProfile name=chkUseGuestProfile <%=strUGPChecked%>>
		&nbsp;&nbsp;&nbsp;
		Search ID:&nbsp;
		<select <%=strSearchIDDisabled%> id=selSearchType name=selSearchType class=myfont>
			<option <%if rsCompany.Fields("GPSearchID").Value = 0 then Response.Write "selected" end if%> value=0>(None)</option>
			<option <%if rsCompany.Fields("GPSearchID").Value = 1 then Response.Write "selected" end if%> value=1>Guest ID (GCN ID)</option>
			<option <%if rsCompany.Fields("GPSearchID").Value = 2 then Response.Write "selected" end if%> value=2>PMS ID</option>
			<option <%if rsCompany.Fields("GPSearchID").Value = 3 then Response.Write "selected" end if%> value=3>Hotel ID (Custom ID)</option>
		</select>
	</td>
	<td></td>
	<td></td>
</tr>
<tr>
	<td>InfoUSA Import:</td>
	<td>
		<select id=cboInfoUSA name=cboInfoUSA class=myfont>
			<% if (rsCompany.Fields("InfoUSAImport") = true) then %>
					<option selected value="1">Yes</option>
					<option value="0">No</option>
			<% else %>
					<option value="1">Yes</option>
					<option selected value="0">No</option>
			<% end if %>
		</select>
	</td>
	<td></td>
	<td></td>
</tr>
<tr>
	<td>Fax Name:</td>
	<td><input class="myFont" id="txtFaxName" name="txtFaxName" style="WIDTH: 193px" value="<%=rsCompany.Fields("FaxName")%>" onKeyUp="validLen (this,33)"></td>
	<td></td>
	<td></td>
</tr>
<tr valign=top>
	<td>Fax Setup:</td>
	<td><textarea name=txtFaxSetup style="font-family:Tahoma;font-size:11px" rows=10 cols=34 onKeyUp="validLen (this,255)"><%=rsCompany("FaxSetup")%></textarea></td>
	<td></td>
	<td></td>	
</tr>
</table>
</div>

<div id=divDetail   style="display:none;height:500px">
<table class="myFont" border="0" width="780" cellpadding="0" cellspacing="0" align="center">

<tr>
	<td align=right>Backup&nbsp;&nbsp;</td> 
	<td><input class="myFont" id="txtBackupStart" name="txtBackupStart" style="WIDTH: 43px" value="<%=rsCompany.Fields("BackupStart")%>">&nbsp;&nbsp;day(s) before current date</td>
	<td align=right>Backup&nbsp;&nbsp;</td>
	<td><input class="myFont" id="txtBackupEnd" name="txtBackupEnd" style="WIDTH: 43px" value="<%=rsCompany.Fields("BackupEnd")%>">&nbsp;&nbsp;day(s) after current date</td>
</tr>
<!--tr>
<td> Mapping Highway Preference:</td>
<td><input class="myFont" id="txtHighwayPref" name="txtHighwayPref" style="WIDTH: 40px" value="<%if rsCompany.Fields("HighWayPref")&"" ="" then Response.Write 45 else Response.write rsCompany.Fields("HighWayPref") end if%>"> %</td>
<td> Favor Shortest Route:</td>
<td><input type="checkbox" class="myFont" id="chkRoutePref" name="chkRoutePref" <% if rsCompany.Fields("RoutePref") Then Response.write " Checked " %> style="WIDTH: 40px"></td>
</tr-->
<tr>
<td>Comp. Gen. Latitude:</td><td>
<input disabled class="myFont" id="CGLat" name="CGLat" style="WIDTH: 100px" value="<%=rsHL.Fields("Latitude")%>">
<input id="txtCGLat" name="txtCGLat" type=hidden value="<%=rsHL.Fields("Latitude")%>"></td>
<td>Comp. Gen. Longitude:</td>
<td><input disabled class="myFont" id="CGLon" name="CGLon" style="WIDTH: 100px" value="<%=rsHL.Fields("Longitude")%>">
<input id="txtCGLon" name="txtCGLon" type=hidden value="<%=rsHL.Fields("Longitude")%>">
<!--&nbsp;&nbsp;&nbsp;Use Lat/Long
<input type="checkbox" class="myFont" id="chkLatLong" name="chkLatLong" <% if rsCompany.Fields("useLatLong") Then Response.write " Checked " %> style="WIDTH: 40px"-->
</td>
<%
if rsCompany.Fields("UseCustomLatLong").Value = 0 then
	strChecked = ""
else
	strChecked = "checked"
end if
%>
</tr>
<tr>
<td>Custom Latitude:</td><td>
<input class="myFont" id="txtLat" name="txtLat" style="WIDTH: 100px" value="<%=rsHL.Fields("CGLatitude")%>"></td>
<td>Custom Longitude:</td>
<td><input class="myFont" id="txtLon" name="txtLon" style="WIDTH: 100px" value="<%=rsHL.Fields("CGLongitude")%>">
&nbsp;&nbsp;&nbsp;<input type=checkbox id=chkusecustomlatlong name=chkusecustomlatlong value="on" <%=strChecked%>>&nbsp;Use Custom Lat/Long
</td>

</tr>
<tr>
<td>Assign Distance:</td><td>
<input class="myFont" id="txtAssignDistance" name="txtAssignDistance" style="WIDTH: 100px" value="<%=rsCompany.Fields("AssignDistance")%>"></td>
<td>Type:</td>
<td><input class="myFont" id="txtCompanyType" name="txtCompanyType" style="WIDTH: 100px" value="<%=rsCompany.Fields("CompanyType")%>"></td>
</tr>

	<td>Users Copied on Requests:</td>
	<td colspan=4><input class="myFont" type=text id=txtUsersCopied name=txtUsersCopied value="<%=rsCompany.Fields("EMailCopy")%>" style="width: 591px"></td>
</tr>
<tr>
	<td>Hotel Group:</td>
	<td colspan=3>
		<select class="myFont" id=cboGroups>
			<%do until rsGroups.EOF
				if rsGroups("ID") = rsCompany("ID") then
					strSelected = " selected"
				else
					strSelected = ""
				end if
				Response.Write "<option" & strSelected & " value=" & rsGroups("ID") & ">" & rsGroups("Descr") & "</option>"
				rsGroups.MoveNext
			loop%>
		</select>
	</td>
	<%if isnull(rsCompany.Fields("QuickLinkDefault").Value) then
		strAdd = "checked"
		strView = ""
		strWeb = ""
	else
		select case rsCompany.Fields("QuickLinkDefault").Value
			case 1
				strAdd = "checked"
				strView = ""
				strWeb = ""
			case 2
				strAdd = ""
				strView = "checked"
				strWeb = ""
			case 3
				strAdd = ""
				strView = ""
				strWeb = "checked"
		end select
	end if
	if rsCompany.Fields("QuickLinkForce").Value then
		strChecked = "checked"
	else
		strChecked = ""
	end if
	
	strDCat = rsCompany.Fields("DefaultCategory").Value
	strBCK = rsCompany.Fields("DefaultBCK").Value
	strSortBy = rsCompany.Fields("DefaultSortBy").Value
	strDefaultState = rsCompany.Fields("DefaultState").Value
	
	select case strBCK
		case 1
			strBCK1 = " checked"
			strBCK2 = ""
			strBCK3 = ""
		case 2
			strBCK1 = ""
			strBCK2 = " checked"
			strBCK3 = ""
		case 3
			strBCK1 = ""
			strBCK2 = ""
			strBCK3 = " checked"
	end select
	%>
</tr>
<tr>
	<td>Quick Link Default:</td><td  colspan=3><input <%=strAdd%> type="radio" id="grpLinkType" name="grpLinkType" value="1">&nbsp;Add New Task&nbsp;&nbsp;<input <%=strView%> type="radio" id="grpLinkType" name="grpLinkType" value="2">&nbsp;View Location&nbsp;&nbsp;<input <%=strWeb%> type="radio" id="grpLinkType" name="grpLinkType" value="3">&nbsp;Go to Website&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input <%=strChecked%> type="checkbox" id="chkForceQL" name="chkForceQL" value="on">&nbsp;Force?</td>
	</td>
</tr>
<tr>
	<td>Search Defaults:</td><td colspan=3>
	<table width=100% border=1 style="border-style:outset;border-width:2px" cellspacing=0 cellpadding=2 class=myFont><tr><td>Category:&nbsp;
	<select class=myFont id=DefaultCategory name=DefaultCategory>
	<%set rscat = server.CreateObject("adodb.recordset")
	rscat.Open "select * from tblCategory order by Category",cnsql,adOpenStatic
	do until rscat.EOF
		if strDCat = rscat.Fields("CategoryID").Value then
			strSelected = "selected "
		else
			strSelected = ""
		end if
		Response.Write "<option " & strSelected & "value=" & rscat.Fields("CategoryID").Value & ">" & rscat.Fields("Category").Value & "</option>"
		rscat.MoveNext
	loop
	
	function sbSelected( n )
		if strSortBy = n then
			retval = "selected"
		else
			retval = ""
		end if
		sbSelected = retval
	end function
	%>
	</select>
	</td>
	<td><input <%=strBCK1%> type=radio id=grpBCK name=grpBCK value=1>Begins&nbsp;<input <%=strBCK2%> type=radio id=grpBCK name=grpBCK value=2>Contains&nbsp;<input disabled <%=strBCK3%> type=radio id=grpBCK name=grpBCK value=3>Keyword</td>
	<td>Sort by:
	<select class=myFont id=DefaultSortBy name=DefaultSortBy>
		<option <%=sbSelected(1)%> value=1>Name</option>
		<option <%=sbSelected(2)%> value=2>City</option>
		<option <%=sbSelected(3)%> value=3>Phone</option>
		<option <%=sbSelected(4)%> value=4>Miles</option>
		<option <%=sbSelected(5)%> value=5>Stars</option>
		<option <%=sbSelected(6)%> value=6>Price</option>
	</select>
	</td>
	<td>State:
	<select class=myFont id=DefaultState name=DefaultState>
	<%	Set rsStates = Server.CreateObject("ADODB.Recordset")
	Set rsStates = cnSQL.Execute("SELECT Abbreviation from tlkpState")

	Do While Not rsStates.EOF
	    if (rsStates.Fields("Abbreviation") = strDefaultState) then
			Response.Write "<OPTION selected value=" & rsStates.Fields("Abbreviation") & ">" & rsStates.Fields("Abbreviation") & "</Option>"
	    else
			Response.Write "<OPTION value=" & rsStates.Fields("Abbreviation") & ">" & rsStates.Fields("Abbreviation") & "</Option>"
		end if

		rsStates.MoveNext
	Loop

	%>
	</select>
	</td>
	</tr></table></td>
	</td>
</tr>
<tr>
<td>Time Zone:</td>
<td colspan=3>
	<select class="myFont" id=cboTimeZone name=cboTimeZone>
		<option <%If rsCompany("TimeZone") = -2 Then Response.write " selected "%>value=-2>(GMT -10:00) Hawaii</option>
		<option <%If rsCompany("TimeZone") = -1 Then Response.write " selected "%> value=-1>(GMT -09:00) Alaska</option>
		<option <%If rsCompany("TimeZone") = 0 Then Response.write " selected "%> value=0>(GMT -08:00) Pacific Time (US & Canada); Tijuana</option>
		<option <%If rsCompany("TimeZone") = 1 Then Response.write " selected "%> value=1>(GMT -07:00) Mountain Time (US & Canada); Arizona</option>
		<option <%If rsCompany("TimeZone") = 2 Then Response.write " selected "%> value=2>(GMT -06:00) Central Time (US & Canada); Chicago</option>
		<option <%If rsCompany("TimeZone") = 3 Then Response.write " selected "%> value=3>(GMT -05:00) Eastern Time (US & Canada); Boston</option>
    </select>
</td>
</tr>
<tr>
	<td>
		<table class="myFont" border=0 cellpadding=0 cellspacing=0 width=100%><tr>
		<td>Weather URL:</td><td align=right>http://</td>
		</tr></table>
	</td>
	<td colspan=3><input value="<%=trim(rsCompany("WeatherURL").value)%>" type="text" name="txtWeatherURL" id="txtWeatherURL" style="width:460px" class="myFont"></td>
<tr>
	<td>
		<table class="myFont" border=0 cellpadding=0 cellspacing=0 width=100%><tr>
		<td>Movies URL:</td><td align=right>http://</td>
		</tr></table>
	</td>
	<td colspan=3><input value="<%=trim(rsCompany("MoviesURL").value)%>" type=text id=txtMoviesURL name=txtMoviesURL style=width:460px class=myFont></td>
</tr>
<tr>
	<td>
		<table class="myFont" border=0 cellpadding=0 cellspacing=0 width=100%><tr>
		<td>Ticket Agency URL:</td><td align=right>http://</td>
		</tr></table>
	</td>		
	<td colspan=3><input value="<%=trim(rsCompany("TicketsURL").value)%>" type=text id=txtTicketsURL name=txtTicketsURL style=width:460px class=myFont></td>
</tr>
<tr>
	<td>
		<table class="myFont" border=0 cellpadding=0 cellspacing=0 width=100%><tr>
		<td>Zagat URL:</td><td align=right>http://</td>
		</tr></table>
	</td>		
	<td colspan=3><input value="<%=trim(rsCompany("ZagatURL").value)%>" type=text id=txtZagatURL name=txtZagatURL style=width:460px class=myFont></td>
</tr>
<tr>
	<td>
		<table class="myFont" border=0 cellpadding=0 cellspacing=0 width=100%><tr>
		<td>Flights URL:</td><td align=right>http://</td>
		</tr></table>
	</td>			
	<td colspan=3><input value="<%=trim(rsCompany("FlightsURL").value)%>" type=text id=txtFlightsURL name=txtFlightsURL style=width:460px class=myFont></td>
</tr>
<tr>
<!--td>
Map UNC:</td>
<%'If trim(rsCompany("MapUNC").value) <> "" Then 
'	strMapUNC = trim(rsCompany("MapUNC").value) 
'else 
'	strMapUNC= "\\Shaq\LiveMaps\" & strCompanyID
'End IF
%>

<td><input value="<%=strMapUNC%>" type="text" name="txtMapUNC" id="txtMapUNC" style="width:190px" class="myFont"></td-->
<td>OT Hotel ID:</td><td><input name=txtOTHotelID style="width:40px" class="myFont" type=text value="<%=rsCompany("OTHotelID")%>"></td>
<td valign=top align=right colspan=2>
	<div align=left valign=top>OT Special Message</div>
	<textarea name=txtOTMessage style="font-family:Tahoma;font-size:11px" rows=4 cols=64><%=rsCompany("OTMessage")%></textarea>
	</td>
	
</tr>
<tr>
	<td>GCN Location ID:</td><td><input name=txtLocationID id=txtLocationID style="width:60px" class="myFont" type=text value="<%=rsCompany("LocationID")%>"></td>
	<td>Default Calendar View:</td><td><select class=myFont style=width:200px id=cmbCalView name=cmbCalView></select></td>
</tr>


<tr>
	<td>Itinerary Font:</td>
	<td colspan=1>
		<select onchange="updateItinFont(this.value,document.all('selItinFontSize').value)" id=selItinFontName name=selItinFontName>
			<option selected value=Arial>Arial</option>
			<option value="Comic Sans MS">Comic Sans MS</option>
			<option value="Gill Sans">Gill Sans</option>
			<option value="Helvetica">Helvetica</option>
			<option value="Script">Script</option>
			<option value="System">System</option>
			<option value="Tahoma">Tahoma</option>
			<option value="Times New Roman">Times New Roman</option>
			<option value="Verdana">Verdana</option>
		</select>
		&nbsp;
		<select onchange="updateItinFont(document.all('selItinFontName').value,this.value)" id=selItinFontSize name=selItinFontSize>
			<option value="8">8</option>
			<option selected value="10">10</option>
			<option value="11">11</option>
			<option value="12">12</option>
			<option value="14">14</option>
			<option value="16">16</option>
			<option value="18">18</option>
			<option value="20">20</option>
			<option value="22">22</option>
			<option value="24">24</option>
			<option value="26">26</option>
		</select>
	</td>
	<td colspan=2>
	Default Action: &nbsp;
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	&nbsp;&nbsp;
		<input type="hidden" id=txtActionType name="txtActionType" value="<%=rsCompany("ActionType")%>">
		<select onChange="txtActionType.value=this.value" id=cmdActionType class="myFont" >
		<%
		Set rsActionType = Server.CreateObject("ADODB.Recordset")
		'' fill in the ActionTypes!!		
		    
		'  remarked and modified to force Arrange to top (hard coded ID 2)
		'Set rsActionType = cnSQL.Execute("SELECT 0 as ActionTypeID, '' as ActionType UNION SELECT ActionTypeID, ActionType FROM tlkpActionType ORDER BY ActionType")
		Set rsActionType = cnSQL.Execute("SELECT 0 as ActionTypeID, '' as ActionType, 'a' as myIndex UNION SELECT at.ActionTypeID, at.ActionType, at.ActionType as myIndex FROM tlkpActionType at join tblCompanyActionType  cat on cat.ActionTypeID=at.ActionTypeID  where cat.CompanyID=" & cid & "ORDER BY myIndex")
		        
		Do While Not rsActionType.EOF
			
				If rsActionType.Fields("ActionTypeID") = rsCompany("ActionType") then 
					Response.Write "<OPTION selected "
				Else
					Response.Write "<OPTION "
				End If

			Response.Write "value=" & rsActionType.Fields("ActionTypeID") & ">" & rsActionType.Fields("ActionType") & "</Option>"
			rsActionType.MoveNext
		Loop
		rsActionType.Close 
		set rsActionType = nothing
		%>
		
		</select>
	</td>
</tr>
<tr>
	<td>Default New Tasks to Rollover</td>
	<td colspan=3><input type="checkbox" class="myFont" id="chkRollover" name="chkRollover" <% if rsCompany.Fields("Rollover") Then Response.write " Checked " %> style="WIDTH: 40px"></td>
</tr>
<tr>
	<td> Concierge Phone </td>
	<td>
		<input type=text class="myFont" name=txtConciergePhone id=txtConciergePhone value="<%=rsCompany.Fields("ConciergePhone")%>">
		<input type="checkbox" class="myFont" id="chkConciergePhone" name="chkConciergePhone" <% if rsCompany.Fields("showConciergePhone") Then Response.write " Checked " %> style="WIDTH: 40px">
		Show on Stationary
	</td>
	<td>SuperShuttle ID:</td>
	<td><input type=text class="myFont" name=txtSSID id=txtSSID value="<%=rsCompany.Fields("SSID")%>">
</table>
</div>
<div id=divDisclaimer style="display:none;height:500px;">
<table>
<tr><td><b>Property Level Disclaimer:<b></td></tr>
<tr><td>
	<textarea id=txtDisclaimer name=txtDisclaimer style="height:400px;width:500px"><%=rsCompany.Fields("Disclaimer").Value%></textarea>
</td></tr>
</table>
</div>
<div id=divDepartments style="display:none;height:500px;">
<table style="font-family:arial;font-size:13px;">
<tr><td><b>All Available Departments</b></td><td>&nbsp;</td><td><b>Assigned to this Hotel</b></td></tr>
<tr>
	<td>
		<select onclick="SelectDepartment(this)" ondblclick=cmdAssignDepartment_onclick() id=lstDepartment size=2 style="height:300px;width:250px"></select></td><td valign=middle><input type=button id=cmdAssignDepartment value=" > " onclick="cmdAssignDepartment_onclick()"><br><br><input type=button id=cmdUnAssignDepartment value = " < " onclick="cmdUnAssignDepartment_onclick()"></td><td>
		<select onclick="SelectDepartment(this)" ondblclick=cmdUnAssignDepartment_onclick() id=lstDepartmentAssigned size=2 style="height:300px;width:250px">
	</td>
</tr>
<tr>
	<td colspan=3>
		<input value=Add type=button id=cmdAdd LANGUAGE=javascript onclick="return cmdAdd_onclick()">&nbsp;<input value=Edit type=button id=cmdEdit onclick="return cmdEdit_onclick()">&nbsp;<input value=Delete type=button disabled id=cmdDelete onclick="return cmdDelete_onclick()"></td></tr>
</table>

</div>
<%



set rsPref = Server.CreateObject ("Adodb.recordset")

rsPref.Open "Select top 1 * from tblCompanyPrefs where CompanyID=" & strCompanyID, cnSQL





Function WCFF (fName,fValue)

	
	On Error Resume Next 
	
	str = fName & "&nbsp;&nbsp;"
	str = str & "<input type=checkbox value=1 name=boo" & fValue & " id=boo" & fValue & " "
	If rsPref(fValue)=true Then
		str = str & " checked "
	End If
	str = str & ">"
	
	WCFF = str
	

End Function

%>

<div id=divPreferences style="display:none;height:500px;">
<table border=1 cellspacing=2 cellpadding=0 style="font-family:arial;font-size:13px;">
<tr>
	<td colspan=2>
		<b>Print Fields	</b>
	</td>
</tr>

<tr>
	<td>
		<%=WCFF("Company Name","CompanyName")%>
	</td>
	<td>
		<%=WCFF("Contact","Contact")%>
	</td>
	
</tr>

<tr>
	<td>
		<%=WCFF("Address","Street")%>
	</td>
	<td>
		<%=WCFF("City, State, Zip","City")%>
	</td>
	
</tr>

<tr>
	<td>
		<%=WCFF("Phone","Phone")%>
	</td>
	<td>
		<%=WCFF("Alternate Phone","PhoneAlternate")%>
	</td>
	
</tr>

<tr>
	<td>
		<%=WCFF("Fax Number","FaxNumber")%>
	</td>
	<td>
		<%=WCFF("Pager Number","PagerNumber")%>
	</td>
</tr>


<tr>
	<td>
		<%=WCFF("E-Mail Address","Email")%>
	</td>
	<td>
		<%=WCFF("Hotel Rating","HotelRating")%>
	</td>
	
</tr>

<tr>
	<td>
		<%=WCFF("Cost Rating","CostRating")%>
	</td>
	<td>
		<%=WCFF("Recommended","Recommended")%>
	</td>
	
</tr>

<tr>
	<td>
		<%=WCFF("Live Music","live_music")%>
	</td>
	<td>
		<%=WCFF("Cross Streets","CrossStreets")%>
	</td>
	
</tr>



<tr>
	<td>
		<%=WCFF("Website","lWebsite")%>
	</td>
	<td>
		<%=WCFF("Private Notes","PrivateNotes")%>
	</td>
	
</tr>

<tr>
	<td>
		<%=WCFF("Notes","Notes")%>
	</td>
	<td>
		<%=WCFF("Hours","Hours")%>
	</td>
	
</tr>

<tr>
	<td>
		<%=WCFF("Price","Price")%>
	</td>
	<td>
		<%=WCFF("Parking","Parking")%>
	</td>
</tr>

<tr>
	<td>
		<%=WCFF("Teaser","Teaser")%>
	</td>
	<td>
		<%=WCFF("Synopsis","Synopsis")%>
	</td>
	
</tr>
<tr>
	<td>
		<%=WCFF("Meal","Meal")%>
	</td>
	<td>
		<%=WCFF("Amenity","Amenity")%>
	</td>
</tr>
<tr>
	<td>
		<%=WCFF("Atmosphere","Atmosphere")%>
	</td>
	<td>
		<%=WCFF("Payment","Payment")%>
	</td>
	
</tr>

<tr>
	<td>
		<%=WCFF("Neighborhood","Neighborhood")%>
	</td>
	<td>
		<%=WCFF("General Directions","Directions")%>
	</td>
	
</tr>
<tr>
	<td>
		<%=WCFF("Directions From Hotel","DirectionsToLocation")%>
	</td>
	<td>
		<%=WCFF("Directions To Hotel","DirectionsToHotel")%>
	</td>
	
</tr>

<tr>
	<td>
		<%=WCFF("Hotel Notes","HotelNotes")%>
	</td>
	<td>
		<%=WCFF("Transportation","Transportation")%>
	</td>
</tr>

<tr>
	<td>
		<%=WCFF("Display Maps","MainMap")%>
	</td>
</tr>




</table>
</div>

<table class="myFont" border="0" width="780" cellpadding="0" cellspacing="0" align="center">
<tr>
	<td align="center" width="750" colspan="4">
		<input style="FONT-SIZE: xx-small; COLOR: fucia; WIDTH: 100px;height: 22px" type=button value="Itinerary Intro..." onclick="EditItinIntro()" id=cmdItinIntro name=cmdItinIntro>&nbsp;&nbsp;&nbsp;&nbsp;
		<input style="FONT-SIZE: xx-small; COLOR: #0075AA; WIDTH: 100px;height: 22px" type=button value="Check Address" onclick="CheckSelectedAddress()" id=button1 name=button1>&nbsp;&nbsp;&nbsp;&nbsp;
		<input id="submit1" name="submit1" style="FONT-SIZE: xx-small;HEIGHT: 22px; LEFT: 150px; TOP: 831px; WIDTH: 100px" type="submit" value="Submit">&nbsp;&nbsp;&nbsp;&nbsp;
		<input id="cmdMainMenu" name="cmdMainMenu" style="FONT-SIZE: xx-small; HEIGHT: 22px; LEFT: 150px; TOP: 831px; WIDTH: 100px" type="button" value="Exit">
		<input id="cmdPrinter" name="cmdPrinter" style="FONT-SIZE: xx-small; HEIGHT: 22px; LEFT: 150px; TOP: 831px; WIDTH: 100px" type="button" value="Priner List" oonclick"ShowPrinterList()">
	</td>
</tr>
</table>
</td>
</tr>
</table>

</td>
</tr>
</table>
<span style="font-family: tahoma; font-size: 10">&nbsp;&nbsp;&nbsp;&nbsp;* note: You may include multiple Users to be Copied by seperating the e-mail addresses by a semi-colon (i.e. pat@wof.com; bob@pir.com)</span>
</p>

<!--<a href="Switchboard3.asp"> Back to Home Page</a>-->

<%end if%>
<script language=javascript>
function onload()
{
<%
response.write strCalView & vbcrlf
response.write "setItinFontVals('" & strItinFontName & "'," & strItinFontSize & ");"
response.write "updateItinFont('" & strItinFontName & "'," & strItinFontSize & ");"
%>
}
</script>

<input type=hidden id=txtItinIntro name=txtItinIntro value="<%=rsCompany("ItinIntroTemplate").Value%>">
</form>
</body>
</html>

<script language=vbscript>
sub EditItinIntro()
	window.form1.txtItinIntro.value =  window.showModalDialog("EditItinIntro.asp",document.all("txtItinIntro").value,"dialogwidth:700px;dialogheight:500px;center:yes;status:no;scroll:no")
end sub

sub cmdPrinter_onClick

		 x = window.showModelessDialog("PrinterList.asp?ID=<%=remote.Session("CompanyID")%>" ,"","dialogheight: 340px; dialogwidth: 800px; status: no; center: yes; scroll: no")

end sub

SetRouteColor

Sub cmdMainMenu_onclick
	'window.location.href = "Switchboard3.asp"
	<%
	if request.querystring("CompanyID") = "0" then
		response.write "if msgbox(""Do you want to lose all your changes to this new Hotel?"",vbYesNo,""Lose Changes"") = vbYes then" & vbcrlf
		response.write "	location.href = ""DeleteCompany.asp?Mode=A&CompanyID=" & strCompanyID & """" & vbcrlf
		response.write "end if" & vbcrlf
	else
		if request.querystring("Mode") = "A" then
			response.write("location.href = ""HotelSetupNewMain.asp""")
		else
			response.write("location.href = ""Administration.asp""")
		end if
	end if
	%>
End Sub
</script>
