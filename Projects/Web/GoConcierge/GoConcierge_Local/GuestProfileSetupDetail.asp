<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))

dim rs, cn, strSQL, strWhere, cid

cid = Remote.Session("CompanyID")

strWhere = ""
strField = ""

select case remote.Session("GPSearchID")
	case 0:
		' should never happen
	case 1:
		strField = "GuestID"
	case 2:
		strField = "PMSGuestID"
	case 3:
		strField = "HotelGuestID"
end select

set cn = server.CreateObject("adodb.connection")
set rsSal = server.CreateObject("adodb.recordset")
set rs = server.CreateObject("adodb.recordset")
cn.Open Application("sqlInnSight_ConnectionString")
set rsSal = cn.Execute("select * from tblSalutations order by Salutation")

' see if top 100 is functionally ok
strSQL = "select distinct top 100 case when a.GuestID = gh.GuestID then 1 else 0 end as History, gh.* from vw_GuestHotel gh left join tblAppointment a on gh.GuestID = a.GuestID "

strLastName = trim(Request.Form("txtLastName"))
strFirstName = trim(Request.Form("txtFirstName"))
strPrimaryPhone = trim(Request.Form("txtPrimaryPhone"))
strHotelID = trim(Request.Form("txtHotelGuestID"))
'strGID = trim(Request.Form("txtGID"))

if strLastName <> "" then
	'strLastName = "((soundex(LastName) = soundex('" & strLastName & "')) or (LastName like '" & left(strLastName,4) & "%')) and "
	strLastName = "(gh.LastName like '" & left(strLastName,4) & "%') and "
else
	strLastName = null
end if
if strFirstName <> "" then
	'strFirstName = "((soundex(FirstName) = soundex('" & strFirstName & "')) or (FirstName like '" & left(strFirstName,1) & "%')) and "
	strFirstName = "(gh.FirstName like '" & left(strFirstName,1) & "%') and "
else
	strFirstName = null
end if
if strPrimaryPhone <> "" then
	strPrimaryPhone = "(gh.PrimaryPhone like '" & strPrimaryPhone & "%') and "
else
	strPrimaryPhone = null
end if
if strHotelID <> "" then
	if strField <> "" then
		strHotelID = "gh." & strField & " = " & strHotelID & " and "
	else
		strHotelID = null
	end if
else
	strHotelID = null
end if

'if strGID <> "" then
'	strGID = "GuestID = " & strGID & " and "
'else
'	strGID = null
'end if

if Request.QueryString("load") = "1" then
	strSQL = strSQL & " where gh.GuestID = 0"
else
	if (not isnull(strLastName)) or (not isnull(strFirstName)) or (not isnull(strPrimaryPhone)) or (not isnull(strHotelID))  then  'or (not isnull(strGID)) then 
		strWhere = trim(strLastName & strFirstName & strPrimaryPhone & strHotelID) '& strGID
		strSQL = strSQL & " where " & left(strWhere,len(strWhere)-4)
	end if
	if strWhere = "" then
		strAnd = " where "
	else
		strAnd = " and "
	end if
	strSQL = strSQL & strAnd & "gh.HotelID = " & cid
end if

'Response.Write "<textarea>" & strSQL & "</textarea>"
'Response.End

set rs = cn.Execute(strSQL)

strStandardBGC = "beige"
%>
<HTML>
<HEAD>
<style>
	.myFont		{ font-family:tahoma; font-size: 11px; }
	.history	{ font-family:tahoma; font-size: 9px; }
</style>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	parent.document.all("txtGuestID").value = "";	
	parent.document.all("txtDisplayID").value = "";	
	parent.document.all("cmdSelect").disabled = true;
	parent.document.all("cmdEdit").disabled = true;
	parent.document.all("cmdMerge").disabled = true;
}

//-->
</SCRIPT>
</HEAD>
<BODY leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0 bgcolor=transparent LANGUAGE=javascript onload="return window_onload()">
<script>
var strRow = "", strSelected = "";
var strStandardBGC = "<%=strStandardBGC%>";

function mo( gid )
{
	if("tr_"+gid != strRow)
	{
		document.all("tr_"+gid).style.backgroundColor = "#488488";
		document.all("tr_"+gid).style.color = "white";
		document.all("tr_"+gid).style.fontWeight = "normal";
		document.all("td8_"+gid).style.color = "white";
	}
}
function mout( gid )
{
	if("tr_"+gid != strRow)
	{
		document.all("tr_"+gid).style.color = "black";
		document.all("tr_"+gid).style.backgroundColor = strStandardBGC;
		document.all("tr_"+gid).style.fontWeight = "normal";
		document.all("td8_"+gid).style.color = "purple";
	}
}
function mdn( gid, displayID )
{
	parent.document.all("txtGuestID").value = gid;
	parent.document.all("txtDisplayID").value = displayID;
	parent.document.all("cmdSelect").disabled = false;
	parent.document.all("cmdAddTask").disabled = false;
	parent.document.all("cmdEdit").disabled = false;

	if(strRow != "")
	{
		document.all(strRow).style.backgroundColor = strStandardBGC;
		document.all(strRow).style.color = "black";
		document.all(strRow).style.fontWeight = "normal";
		document.all("td8_"+strRow.substr(strRow.indexOf("_")+1)).style.color = "purple";
	}
	
	strRow = "tr_"+gid;

	document.all("td8_"+strRow.substr(strRow.indexOf("_")+1)).style.color = "white";
	document.all(strRow).style.backgroundColor = "sienna";
	document.all(strRow).style.color = "white";
	document.all(strRow).style.fontWeight = "bold";
		
	document.all("td8_"+gid).style.fontWeight = "normal";
}
function mdclk( gid, displayID )
{
	parent.document.all("txtGuestID").value = gid;
	parent.document.all("txtDisplayID").value = displayID;
	parent.document.all("cmdSelect").disabled = false;
	parent.document.all("cmdAddTask").disabled = false;
	parent.document.all("cmdEdit").disabled = false;
	parent.cmdSelect_onclick();
}

function mergeStatus()
{
	strSelected = '';
	var e = window.document.all.tags("input")
	for(var i=0; i < e.length; i++)
		if(e[i].id.indexOf('chk') > -1)
			if(e[i].checked)
				strSelected += e[i].id.substr(e[i].id.indexOf('_')+1) + ',';

	if(strSelected.length > 0)
		strSelected = strSelected.substr(0,strSelected.length-1);
	
	if(strSelected.indexOf(',') > 0)
		parent.document.all("cmdMerge").disabled = false;
	else
		parent.document.all("cmdMerge").disabled = true;
}

function history( gid ) {
	var name = document.all("td2_"+gid).innerText+' '+document.all("td3_"+gid).innerText+' '+document.all("td4_"+gid).innerText;
	var x = window.showModalDialog("GuestProfileHistory.asp?name="+name+"&gid="+gid,"","dialogWidth:700px;dialogHeight:450px;center:yes;scroll:no;status:no")
}
</script>
<table cellpadding=3px style=overflow:hidden class=myFont>
	<%
	do until rs.EOF
		gid = trim(rs.Fields("GuestID").Value)
		
		if strField = "" then
			displayID = ""
		else
			if isnull(rs.Fields(strField).Value) then
				displayID = ""
			else
				displayID = trim(rs.Fields(strField).Value)
			end if
		end if
				
		if displayID = "" then
			dID = "''"
		else
			dID = displayID
		end if
		Response.Write "<tr id=tr_" & gid & " style=cursor:hand;background-color:" & strStandardBGC & " onmouseover=mo(" & gid & ") onmousedown=mdn(" & gid & "," & dID & ") onmouseout=mout(" & gid & ") ondblclick=mdclk(" & gid & "," & dID & ")>"
		Response.Write "<td id=td2_" & gid & "><div nowrap style=overflow:hidden;padding-top:0px;width:79px>" & rs.Fields("Salutation").Value & "</div></td>"
		Response.Write "<td id=td3_" & gid & "><div nowrap style=overflow:hidden;padding-top:0px;width:110px>" & rs.Fields("LastName").Value & "</div></td>"
		Response.Write "<td id=td4_" & gid & "><div nowrap style=overflow:hidden;padding-top:0px;width:100px>" & rs.Fields("FirstName").Value & "</div></td>"
		Response.Write "<td id=td5_" & gid & "><div nowrap style=overflow:hidden;padding-top:0px;width:129px>" & rs.Fields("Company").Value & "</div></td>"
		Response.Write "<td id=td6_" & gid & "><div nowrap style=overflow:hidden;padding-top:0px;width:95px>" & formatPhone(rs.Fields("PrimaryPhone").Value) & "</div></td>"
		Response.Write "<td id=td1_" & gid & "><div nowrap style=overflow:hidden;padding-top:0px;width:50px>" & displayID & "</div></td>"
		Response.Write "<td id=td7_" & gid & " style=padding-left:0px;padding-right:4px align=center valign=top><div nowrap style=padding-top:0px;vertical-align:top;overflow:hidden;height:16px;width:18px;clip:auto><input onclick=mergeStatus() type=checkbox id=chk_" & gid & " name=chk_" & gid & "></div></td>"
		if rs.Fields("History").Value = 1 then
			Response.Write "<td  title=""View History"" unselectable=on onclick=history(" & gid & ") onmousedown=""document.all('td8_" & gid & "').style.borderStyle = 'inset'"" onmouseout=this.style.borderStyle='outset' onmouseup=this.style.borderStyle='outset' id=td8_" & gid & " style=padding-left:1px;padding-right:1px;color:purple;text-align:center;border-style:outset;border-width:2px><div unselectable=on class=history nowrap style=overflow:hidden;padding-top:0px;width:12px>"
			Response.Write "<b>H</b>"
		else
			Response.Write "<td title=""No History"" unselectable=on id=td8_" & gid & " style=padding-left:1px;padding-right:1px;color:purple;text-align:center;border-style:outset;border-width:2px><div unselectable=on class=history nowrap style=overflow:hidden;padding-top:0px;width:12px>"
		end if
		Response.Write "</div></td>"
		Response.Write "</tr>"
		rs.MoveNext
	loop
	%>
</table>

</BODY>
</HTML>
<%
'rs.Close        
'set rs = nothing

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

function formatPhone( strPhone )
	str = trim(strPhone)
	if len(trim(str)) = 10 then
		strRetVal = "(" & Left(str,3) & ") " & Mid(str,4,3) & "-" & Right(str,4)
	else
		strRetVal = str
	end if
	formatPhone = strRetVal
end function
%>
