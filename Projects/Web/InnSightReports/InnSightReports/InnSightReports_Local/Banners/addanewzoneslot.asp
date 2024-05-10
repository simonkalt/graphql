<% 	If strTask="Edit" Then
		strTask2="Update"
		strButtonText="Update Zone"
		'strSQL2="SELECT * FROM ZoneStatsSum WHERE ZoneID=" & strZoneID  & " AND UserID= " & CLng(Session("BanManProSiteID"))
		'Set rsZoneStatsSum=connBanManPro.Execute(strSQL2)
	Else
		strTask2="Insert"
		strButtonText="Submit New Zone"
	End If
	

	'EVENLY Distributed Campaigns
	'create array of data
	intCounter=0
	blnNoEvenCampaigns=True
	Do While Not rsEvenCampaigns.EOF
		If rsEvenCampaigns("CampaignSiteDefault") <> True Then
			blnNoEvenCampaigns=False
			intCounter=intCounter+1
			ReDim Preserve arrEvenCampaignID(intCounter)
			ReDim Preserve strEvenCampaignName(intCounter)
			ReDim Preserve blnEvenSelected(intCounter)
			ReDim Preserve strEvenWeighting(intCounter)
			ReDim Preserve blnValidCampaign(intCounter)
			arrEvenCampaignID(intCounter)=rsEvenCampaigns("CampaignID")
			strEvenCampaignName(intCounter)=rsEvenCampaigns("CompanyName") & ": " & rsEvenCampaigns("CampaignName")
			blnEvenSelected(intCounter)=False
			blnValidCampaign(intCounter)=False
		End If
		rsEvenCampaigns.MoveNext
	Loop

	'7/16/00 NEW CODE FOR NOTING EXPIRED CAMPAIGNS ****************************************
	'Determine if campaigns are valid or not
	strSQL="SELECT validcampaigns_type.CampaignID,"
    	strSQL=strSQL & " Advertisers.CompanyNamE FROM validcampaigns_type INNER JOIN Advertisers ON validcampaigns_type.AdvertiserID = Advertisers.AdvertiserID INNER "
    	strSQL=strSQL & " JOIN Campaigns ON validcampaigns_type.CampaignID = Campaigns.CampaignID "
	strSQL=strSQL & " Where (Campaigns.UserID=" & CLng(Session("BanManProSiteID")) & " Or Campaigns.UserID=0)  AND (Campaigns.CampaignDistribution='Normal' OR Campaigns.CampaignDistribution='Weighted')  AND Campaigns.CampaignDistribution<>'Keyword' ORDER BY Advertisers.CompanyName ASC,Campaigns.CampaignName ASC"	
	Set rsTemp=connBanManPro.Execute(strSQL)
	Do While Not rsTemp.EOF
		intCounter=1
		If IsArray(arrEvenCampaignID) Then
			Do While intCounter <= Ubound(arrEvenCampaignID)
				If rsTemp("CampaignID")=arrEvenCampaignID(intCounter) Then
					blnValidCampaign(intCounter)=True
					Exit Do
				End If
				intCounter=intCounter+1
			Loop
		End if
		rsTemp.MoveNext
	Loop
	Set rsTemp=Nothing
	'7/16/00 END NEW CODE FOR NOTING EXPIRED CAMPAIGNS *************************************

	If strTask="Edit" Then
		'find matching selected campaigns
		On Error Resume Next
		rsZoneCampaigns.MoveFirst
		On Error GoTo 0
		Do While Not rsZoneCampaigns.EOF
			intCounter=1
			If IsArray(arrEvenCampaignID) Then
			Do While intCounter <= Ubound(arrEvenCampaignID)
				If rsZoneCampaigns("CampaignID")=arrEvenCampaignID(intCounter) And rsZoneCampaigns("Even")=True Then
					blnEvenSelected(intCounter)=True
					strEvenWeighting(intCounter)=rsZoneCampaigns("ZoneCampaignWeighting")
					Exit Do
				End If
				intCounter=intCounter+1
			Loop
			End if
			rsZoneCampaigns.MoveNext
		Loop
	End If


	'DEFAULT Campaigns
	'create array of data
	intCounter=0
	blnNoDefaultCampaigns=True
	blnNoDefaultsSelected=True
	Do While Not rsAllDefaults.EOF
			blnNoDefaultCampaigns=False
			intCounter=intCounter+1
			ReDim Preserve arrDefaultCampaignID(intCounter)
			ReDim Preserve strDefaultCampaignName(intCounter)
			ReDim Preserve blnDefaultSelected(intCounter)
			arrDefaultCampaignID(intCounter)= rsAllDefaults("CampaignID")
			strDefaultCampaignName(intCounter)= rsAllDefaults("CompanyName") & ": " & rsAllDefaults("CampaignName")
			blnDefaultSelected(intCounter)=False
			rsAllDefaults.MoveNext
	Loop
	If strTask="Edit" Then
		'find matching selected banners
		On Error Resume Next
		rsSelectedDefaults.MoveFirst
		On Error GoTo 0
		Do While Not rsSelectedDefaults.EOF
			intCounter=1
			If IsArray(arrDefaultCampaignID) Then
			Do While intCounter <= Ubound(arrDefaultCampaignID)
				If rsSelectedDefaults("CampaignID")=arrDefaultCampaignID(intCounter) Then
					blnDefaultSelected(intCounter)=True
					blnNoDefaultsSelected=False
					Exit Do
				End If
				intCounter=intCounter+1
			Loop
			End if
			rsSelectedDefaults.MoveNext
		Loop
	End If
%>
<html>

<head>
<title></title>
</head>

<body>

<script LANGUAGE="JavaScript">
<!-- Begin
// This script handles adding and removing campaigns to the zone
sortitems = 1;  // Automatically sort items within lists? (1 or 0)

function move(fbox,tbox) {
for(var i=0; i<fbox.options.length; i++) {
if(fbox.options[i].selected && fbox.options[i].value != "") {
var no = new Option();
no.value = fbox.options[i].value;
no.text = fbox.options[i].text;
tbox.options[tbox.options.length] = no;
fbox.options[i].value = "";
fbox.options[i].text = "";
   }
}
BumpUp(fbox);
if (sortitems) SortD(tbox);
}
function BumpUp(box)  {
for(var i=0; i<box.options.length; i++) {
if(box.options[i].value == "")  {
for(var j=i; j<box.options.length-1; j++)  {
box.options[j].value = box.options[j+1].value;
box.options[j].text = box.options[j+1].text;
}
var ln = i;
break;
   }
}
if(ln < box.options.length)  {
box.options.length -= 1;
BumpUp(box);
   }
}

function SortD(box)  {
var temp_opts = new Array();
var temptext = new Object();
var tempvalue = new Object();
for(var i=0; i<box.options.length; i++)  {
temp_opts[i] = box.options[i];
}
for(var x=0; x<temp_opts.length-1; x++)  {
for(var y=(x+1); y<temp_opts.length; y++)  {
if(temp_opts[x].text > temp_opts[y].text)  {
temptext = temp_opts[x].text;
tempvalue = temp_opts[x].value;
temp_opts[x].text = temp_opts[y].text;
temp_opts[x].value = temp_opts[y].value;
temp_opts[y].text = temptext;
temp_opts[y].value = tempvalue;
      }
   }
}
for(var i=0; i<box.options.length; i++)  {
box.options[i].value = temp_opts[i].value;
box.options[i].text = temp_opts[i].text;
   }
}

function submitform(FormName,ListToSubmit){
   for(var x = 0; x < ListToSubmit.length; x++){
          ListToSubmit.options[x].selected = true;
   }
   return false;	
}

// End -->
</script>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.ZoneDescription.value == "")
  {
    alert("Please enter a value for the \"Zone Name\" field.");
    theForm.ZoneDescription.focus();
    return (false);
  }

  if (theForm.ZoneWidth.value == "")
  {
    alert("Please enter a value for the \"Zone Width\" field.");
    theForm.ZoneWidth.focus();
    return (false);
  }

  var checkOK = "0123456789-.,";
  var checkStr = theForm.ZoneWidth.value;
  var allValid = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch == ".")
    {
      allNum += ".";
      decPoints++;
    }
    else if (ch != ",")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Zone Width\" field.");
    theForm.ZoneWidth.focus();
    return (false);
  }

  if (decPoints > 1)
  {
    alert("Please enter a valid number in the \"ZoneWidth\" field.");
    theForm.ZoneWidth.focus();
    return (false);
  }

  var chkVal = allNum;
  var prsVal = parseFloat(allNum);
  if (chkVal != "" && !(prsVal > "0"))
  {
    alert("Please enter a value greater than \"0\" in the \"Zone Width\" field.");
    theForm.ZoneWidth.focus();
    return (false);
  }

  if (theForm.ZoneHeight.value == "")
  {
    alert("Please enter a value for the \"Zone Width\" field.");
    theForm.ZoneHeight.focus();
    return (false);
  }

  var checkOK = "0123456789-.,";
  var checkStr = theForm.ZoneHeight.value;
  var allValid = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch == ".")
    {
      allNum += ".";
      decPoints++;
    }
    else if (ch != ",")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Zone Width\" field.");
    theForm.ZoneHeight.focus();
    return (false);
  }

  if (decPoints > 1)
  {
    alert("Please enter a valid number in the \"ZoneHeight\" field.");
    theForm.ZoneHeight.focus();
    return (false);
  }

  var chkVal = allNum;
  var prsVal = parseFloat(allNum);
  if (chkVal != "" && !(prsVal > "0"))
  {
    alert("Please enter a value greater than \"0\" in the \"Zone Width\" field.");
    theForm.ZoneHeight.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="zones.asp?Task=<%=strTask2%>&amp;ZoneID=<%=strZoneID%>" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1">
  <div align="center">
    <center>
    <table border="0" cellpadding="0" cellspacing="0" width="590">
      <tr>
        <td><a href="help/zones.htm" target="_new"><img border="0" src="images/ListingofAllZones.gif" WIDTH="590" HEIGHT="30"></a></td>
      </tr>
    </table>
    </center>
  </div>
  <div align="center"><center><table border="0" cellpadding="2" cellspacing="0" width="590" background="images/tableback.gif">
<% If strTask="Edit" Then %>
    <tr>
      <td width="578" align="right" colspan="2" bgcolor="#7A74FA"><div align="center"><center><p><font face="Arial" size="2" color="#000000">This Zone Averages <strong><%If IsNull(rsz("ZoneAverage")) Then%>0<%Else%><%=rsz("ZoneAverage")%><%End If%></strong> Impressions/Day
      based on <strong><%=Application("ZoneAverageDays")%></strong>  Days.</font>
          </div>
        </center></td>
    </tr>
<%  
End If%>
    <tr align="center">
      <td width="175" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Zone
        Name:</strong></font></td>
      <td width="391" align="left"><!--webbot bot="Validation" S-Display-Name="Zone Name" B-Value-Required="TRUE" --><input type="text" name="ZoneDescription" size="40" <% If strTask="Edit" Then %>value="<%=rsz("ZoneDescription")%>" <%End If%>></td>
    </tr>
    <tr align="center">
      <td width="175" align="right"><font face="Arial" size="2"><strong>Optional
        Description/URL:</strong></font></td>
      <td width="391" align="left"><input type="text" name="ZonePageURL" size="40" <% If strTask="Edit" Then %>value="<%=rsz("ZonePageURL")%>" <%End If%>></td>
    </tr>
<tr>
      <td width="175" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Mode:</strong></font></td>
      <td width="391" align="left"><div align="left"><p><select name="ZoneMode" size="1">
        <option <% If strTask="Edit" Then %><%If rsz("ZoneMode")="SSI" Then%>selected<%End If%><%End If%> value="SSI">SSI</option>
        <option <% If strTask="Edit" Then %><%If rsz("ZoneMode")="HTML" Then%>selected<%End If%><%End If%> value="HTML">HTML</option>
      </select>
        </div>
      </td>
</tr>
    <tr align="center">
      <td width="175" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Zone
        Width:</strong></font></td>
      <td width="391" align="left"><!--webbot bot="Validation" S-Display-Name="Zone Width" S-Data-Type="Number" S-Number-Separators=",." B-Value-Required="TRUE" S-Validation-Constraint="Greater than" S-Validation-Value="0" --><input type="text" name="ZoneWidth" size="5" <% If strTask="Edit" Then %>value="<%=rsz("ZoneWidth")%>" <%Else%>value="468" <%End If%>></td>
    </tr>
    <tr align="center">
      <td width="175" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Zone
        Height:</strong></font></td>
      <td width="391" align="left"><!--webbot bot="Validation" S-Display-Name="Zone Width" S-Data-Type="Number" S-Number-Separators=",." B-Value-Required="TRUE" S-Validation-Constraint="Greater than" S-Validation-Value="0" --><input type="text" name="ZoneHeight" size="5" <% If strTask="Edit" Then %>value="<%=rsz("ZoneHeight")%>" <%Else%>value="60" <%End If%>></td>
    </tr>
  </center>
 <% If IsArray(arrEvenCampaignID) Then %>
 <center>
<tr>
      <td align="right" colspan="2">
        <p align="left"><img border="0" src="images/slotoptionbar.gif" WIDTH="586" HEIGHT="25">
      </td>
</tr>
   
   

<tr>
      <td align="right" colspan="2">
        <p align="center"><font face="Arial" size="2">Select all slot campaigns
        to include in this zone.</font></td>
</tr><tr align="center">
<td align="right" colspan="2">

        <p align="center"><font face="Arial" size="2">(Formatted as &quot;Company: Campaign&quot;)</font>
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="500" background="images/tableback.gif">
            <tr>
              <td align="center"><br>
<b>Available Campaigns:<b><br>          
<select name="AvailableCampaigns">
<%intCounter=1
If blnNoEvenCampaigns=False Then
Do While intCounter <= Ubound(arrEvenCampaignID)
If blnEvenSelected(intCounter)=False And blnValidCampaign(intCounter)=True Then %>
    <option value="<%=arrEvenCampaignID(intCounter)%>"><%=strEvenCampaignName(intCounter)%><%If blnValidCampaign(intCounter)<>True Then%> [EXPIRED]<%End IF%></option>
<%End If
intCounter=intCounter+1
Loop
End If %>
</select><input type="button" value="Include" onclick="move(this.form.AvailableCampaigns,this.form.EvenCampaignsClipboard);" name="includebtn">
<br><br>
<b>Campaigns Running In This Zone:</b><br>  
<select size="6" name="EvenCampaignsClipboard" multiple>
<%
intCounter=1
If blnNoEvenCampaigns=False Then
Do While intCounter <= Ubound(arrEvenCampaignID)
If blnEvenSelected(intCounter)=True And blnValidCampaign(intCounter)=True Then %>
          <option <%If strTask="Edit" Then%><%End If%> value="<%=arrEvenCampaignID(intCounter)%>"><%=strEvenCampaignName(intCounter)%><%If blnValidCampaign(intCounter)<>True Then%> [EXPIRED]<%End IF%></option>
<%End If
intCounter=intCounter+1
Loop
End If
%></select><br>
<input type="button" value="Remove Campaign from Zone" onclick="move(this.form.EvenCampaignsClipboard,this.form.AvailableCampaigns);" name="removebtn">
<%
intCounter=1
If blnNoEvenCampaigns=False Then
Do While intCounter <= Ubound(arrEvenCampaignID)
If blnValidCampaign(intCounter)<>True And blnEvenSelected(intCounter)=True Then %>
<input type="hidden" name="EvenCampaigns" value="<%=arrEvenCampaignID(intCounter)%>">
<%
End If
intCounter=intCounter+1
Loop
End If
%>
        <br>&nbsp;</td></tr></table>
        </div>
<%End If%>
    <center>
<% ' campaign site default *****************************************************************  %>
        </center>
    <tr align="center">
      <td width="566" align="left" colspan="2"><a href="help/zones.htm#Default" target="_blank"><img border="0" src="images/includedefaultcampaigns.gif" WIDTH="586" HEIGHT="25"></a></td>
    </tr>

    <tr align="center">
      <td width="566" align="center" colspan="2">
      <div align="center">
  <center>
  <table border="1" cellpadding="2" cellspacing="0" width="400" background="images/tableback.gif" bordercolor="#000000">
    <tr>
      <td align="center" width="326"><font face="Arial" size="2"><b>Company: Campaign</b></font></td>
      <td align="center" width="70"><font face="Arial" size="2"><b>Selected</b></font></td>
    </tr>
    <tr>
      <td align="center" width="326"><font face="Arial" size="2">Don't include
        any defaults</font></td>
      <td align="center" width="70">
          <input type="radio" value="0" <%If blnNoDefaultsSelected=True Then%> checked<%End If%> name="DefaultCampaign">
      </td>
    </tr>
<% 
intCounter=1
If blnNoDefaultCampaigns=False Then
Do While intCounter <= Ubound(arrDefaultCampaignID)
	  %>

    <tr>
      <td align="center" width="326"><font face="Arial" size="2"><%=strDefaultCampaignName(intCounter)%></font></td>
      <td align="center" width="70">
          <p><input type="radio" value="<%=arrDefaultCampaignID(intCounter)%>" <%If strTask="Edit" Then%><% If blnDefaultSelected(intCounter)=True Then%> checked<%End If%><%End If%> name="DefaultCampaign"></p>
      </td>
    </tr>
<%  intCounter=intCounter+1
Loop
End If %>
  </table>
  </center>
</div>
      </td>
    </tr>

<% 'End Campaign Site default *************************************************************** %>
    <tr align="center">
      <td width="175" align="right">&nbsp;</td>
      <td width="391" align="left"><input type="submit" value="<%=strButtonText%>" onClick="submitform(this.form,this.form.EvenCampaignsClipboard);this.form.EvenCampaignsClipboard.name = 'EvenCampaigns';"></td>
    </tr>
    <tr align="center">
      <td width="175" align="right">&nbsp;</td>
      <td width="391" align="left"><font face="Arial" size="2"><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Indicates Required Fields</font></td>
    </tr>
  </table>
    </div>
  <div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="590">
    <tr>
      <td><img border="0" src="images/bottomblue.gif" WIDTH="590" HEIGHT="30"></td>
    </tr>
  </table>
  </center>
</div>
</form>

</body>
</html>


