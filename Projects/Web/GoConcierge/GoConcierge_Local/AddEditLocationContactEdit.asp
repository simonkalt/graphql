<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Function FormatPhone (phone)
		phone2 = Replace(phone,"(","")
		phone2 = Replace(phone2,")","")
		phone2 = Replace(phone2,"-","")
		phone2 = Replace(phone2," ","")
		phone2 = Replace(phone2,".","")
		phone2 = Replace(phone2,"/","")
		phone2 = Replace(phone2,"\","")
		 FormatPhone = "(" & Left(phone2,3) & ") " & Mid(phone2,4,3) & "-" & Right(phone2,4)
End Function


Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))

txtLocationID = Request.QueryString("LID")


Set cnSQL = Server.CreateObject("ADODB.Connection")
cnSQL.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")

set rs = cnSQL.Execute ("Select * from tblLocation where LocationID=" & txtLocationID)
If not rs.BOF and not rs.EOF Then
	txtLocationName = rs("CompanyName")
	txtLocationAddress = rs("Street")
	txtLocationCSZ = rs("City") & ", " & rs("State") & " " & rs("ZIP")
	txtLocationPhone = FormatPhone(rs("Phone"))
	rs.Close()
End If

strSQL = "Select * from tblLocationContact where LocationID=" & txtLocationID & " and (GlobalContact=1 or CompanyID=" & remote.Session("CompanyID") & ") and ContactID not in (select ContactID from tblCompanyContactExceptions where CompanyID=" & remote.Session("CompanyID") & ")"


rs.Open strSQL, cnSQL

%>

<HTML>
<HEAD>

<TITLE>E-Mail Vendor</TITLE>

<style type="text/css">
<!--
TABLE	 { font-family: tahoma; font-size: 11px; }
.search	 { font-family: tahoma; font-size: 11px; width: 100px; }
.norm	 { font-family: tahoma; font-size: 11px; height:19px; width:175px}
.tdnorm	 { font-family: tahoma; font-size: 11px; height:19px; width:180px}
.lbl	 { font-family: tahoma; font-size: 11px; width:81px}
.buttons { font-family: tahoma; font-size: 11px; width: 70px; height: 22px; }
.refreshbutton { font-family: tahoma; font-size: 11px; width: 50px; height: 19px; }

-->
</style>

<!--#INCLUDE file="PhoneMask.asp"-->


</HEAD>

<BODY bgcolor="#FCE8AB" LANGUAGE=javascript onload="return window_onload()">
<form method=post id=form1 name=form1>
<TABLE border=0 cellpadding=6 BORDER=0>
<TR colspan=2>
<TD>
<table border=1>
	<TR style="background-color:#FAD666;height:25px;">
		<td style="width:100px">Vendor Name:</td><td style="color:blue"><strong><%=txtLocationName%></strong></td>
	</tr>
	<tr style="background-color:#FAD666;">	
		<td></td><td style="color:blue"><%=txtLocationAddress%></td>
	</tr>
	<tr style="background-color:#FAD666;">
		<td></td><td style="color:blue"><%=txtLocationCSZ%></td>
	</tr>
	<tr style="background-color:#FAD666;">
		<td></td><td style="color:blue"><%=txtLocationPhone%></td>
	</tr>
</table>	
</TD>
<TD></TD>
</TR> 
	
	<TR>
		<TD colspan=2>
			<table "display:block;height:130px" width="100%">
				<tr>
					<td>
						<div id=divDetail style="display:block;height:130px">
						<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD class="lbl" ALIGN=right>Contact Name:</TD>	<TD class="tdnorm" ALIGN=left>
								<input class="norm" id=txtContactName name=txtContactName></TD>
							</TR>
							<TR>
								<TD class="lbl" ALIGN=right>Phone:</TD>
								<TD class="tdnorm" ALIGN=left>
									<!--input class=norm style="width:100px" id=txtPhone name=txtPhone-->
									<table cellpadding=0 cellspacing=0>
									<tr><td>
									<script language="JavaScript1.2">
										CreatePhoneField ( "Phone", "font-family: Tahoma; font-size: 11", "13px", 100 );
									</script>
									<input type=hidden id="txtPhone" name="txtPhone">
									</td><td>
									&nbsp;&nbsp;Ext.&nbsp;<input type=text id=txtPhoneExt name=txtPhoneExt class=norm style="width:46px">
									</td></tr>
									</table>
								</TD>
							</TR>
							<TR>
								<TD class="lbl" ALIGN=right>Fax:</TD>	<TD class="tdnorm" ALIGN=left>
									<script language="JavaScript1.2">
										CreatePhoneField ( "Fax", "font-family: Tahoma; font-size: 11", "13px", 100 );
									</script>

									<input class="norm" type="hidden" id=txtFax name=txtFax>
								</TD>
							</TR>
							<TR>
								<TD class="lbl" ALIGN=right>E-Mail:</TD>	<TD class="tdnorm" ALIGN=left><input class="norm" onKeyUp="validateEmail()" id=txtEmail name=txtEmail></TD>
								
							</TR>
							<TR>
								<TD colspan=2 class="tdnorm" ALIGN=right>
							Global Contact:<input type="checkbox" id=txtShare name=txtShare></TD>
							</TR> 
							
						</TABLE>
						</div>
						
			</table>
		</TD>
	</TR>
	<TR>
		<TD colspan=2>
			<table cellpadding="5" cellspacing="0" width="100%">
				<tr>
					<td align="right">
						<input type="button" class="buttons" id="cmdSubmit" value="Save" style="color:green" onclick="addContact()">
					</td>
					<td align="left">
						<input type="button" class="buttons" id="cmdCancel" value="Cancel" style="color:red"  onclick="returnCancel()">
					</td>
				</TR>
			</table>
		</TD>
	</TR>
</TABLE>
</form>

<div id="divLoading" name="divLoading" style="position: absolute; z-index: 1; visibility: hidden">
	<iframe height="69" width="169" frameborder="0" style="border-style: none; border-width: 1px;" src="LoadingDiv.asp?v=2&String=Saving, please wait..." id="frameLoadingDiv" allowTransperancy="true" scrolling="no"></iframe>
</div>
	
</script>

<script>

var strAction = 'a'; // Default action is to add 
var curContact = '<%=request.querystring("cid")%>';

function editContact()
{

	var cid = '<%=Request.querystring("cid")%>'
	var str='';
	str += 'ContactId=' + cid;
	str += '&Action=g'

	var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
	xmlHttp.open("Get", "EmailVendorContactAddConfirm.asp?" + str, false)
	xmlHttp.send()
	
	var s = new String(xmlHttp.responseText);
	var tmparr = s.split ('|')
	
	document.all("txtContactName").value = tmparr[0];
	FillPhone("Phone",tmparr[1]);
	FillPhone("Fax",tmparr[2]);
	document.all("txtEmail").value = tmparr[3];	
	if (tmparr[4]==1)
		document.all('txtShare').checked = true
		
	validateEmail()	

	strAction = 'e';
	
	
}

function validateEmail()
{

var e = document.all("txtEmail").value.toString();
var v1 = e.indexOf('@')
var v2 = e.indexOf('.')
var l = e.length;

	if ((v1 > 0) && (v2 > 0) && (v2 > v1) && (l > (v2+3)))
	{
		document.all("cmdSubmit").disabled = false;
	}
	else
	{
		document.all("cmdSubmit").disabled = true;
	}
	
}

function addContact()
{
		save();
}

function returnCancel()
{
		window.close();
}

var timer
function save()
{
	window.divLoading.style.visibility = "visible";
	//timer = window.setInterval("save_process()",10)
	save_process()
}

function save_process()
{
window.clearInterval(timer);

var str = '';

document.all("txtPhone").value = document.all("Phone").value;
document.all("txtFax").value = document.all("Fax").value;
str += 'txtContactName=' + escape(document.all("txtContactName").value);
str += '&txtPhone='  + escape(document.all("txtPhone").value);
str += '&txtFax='  + escape(document.all("txtFax").value);
str += '&txtEMail='  + escape(document.all("txtEMail").value);
str += '&txtPhoneExt=' + document.all("txtPhoneExt").value;
str += '&ContactId=' + curContact;
str +='&Action=' + strAction ;

if (document.all("txtShare").checked)
	str += '&Global=' + '1';
else
	str += '&Global=' + '0';
	
str += '&LID=' + '<%=txtLocationID%>';

var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
xmlHttp.open("Get", "EmailVendorContactAddConfirm.asp?" + str, false)
xmlHttp.send()

if (xmlHttp.responseText=='0') 
	{
	 alert('There was an error saving your contact.');
	}
	else
	{
	
		curContact = xmlHttp.responseText.toString();
		document.all("cmdSubmit").disabled = true;
	}

xmlHttp = null
window.divLoading.style.visibility = "hidden";
returnCancel();
}

function window_onload() {
	
		FillPhone("Phone","");
		window.divLoading.style.top = 80
		window.divLoading.style.left = 100
		
		var cid = '<%=request.querystring("cid")%>';
		if (cid > 0)
		{
		
			editContact()
			strAction = 'e';
		}
		else
			strAction = 'a';
		
			
}

</script>
<%
	rs.Close
	cnSql.Close
	set cnSQL = Nothing
%>
