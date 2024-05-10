<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>

<HTML>
<HEAD>
<script>
	function changeTable()
	{
		//window.frameTables.src = "SetupLookupsFrame.asp?table="+window.selLookups.value;
		window.frameTables.document.location = "SetupLookupsFrame.asp?table="+window.selLookups.value;
	}
</script>
</HEAD>
<BODY>

<select onchange=changeTable() id=selLookups>
	<option value="tlkpAddressType">Address Type</option>
	<option value="tlkpChargeType">Charge Type</option>
	<option value="tlkpDateType">Date Type</option>
	<option value="tlkpPhoneType">Phone Type</option>
	<option value="tlkpPreference">Preferences</option>
	<option value="tlkpProgram">Programs</option>
	<option value="tlkpRelationship">Relationships</option>
	<option value="tlkpRewardsType">Rewards Type</option>
</select>
<br>
<iframe width=100% id=frameTables src="SetupLookupsFrame.asp?table=tlkpAddressType"></iframe>
</BODY>
</HTML>
