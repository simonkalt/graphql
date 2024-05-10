<% 	If strTask="Edit" Then
		strTask2="Update"
		strButtonText="Update Zone"
		'strSQL2="SELECT * FROM ZoneStatsSum WHERE ZoneID=" & strZoneID  & " AND UserID= " & CLng(Session("BanManProSiteID"))
		'Set rsZoneStatsSum=connBanManPro.Execute(strSQL2)
	Else
		strTask2="Insert"
		strButtonText="Submit New Zone"
	End If
	
	'create array of data
	intCounter=0
	blnNoCampaigns=True
	blnDefaultCampaignExists=False
	Do While Not rsCampaigns.EOF
		If rsCampaigns("CampaignSiteDefault") <> True Then
			blnNoCampaigns=False
			intCounter=intCounter+1
			ReDim Preserve arrCampaignID(intCounter)
			ReDim Preserve strCampaignName(intCounter)
			ReDim Preserve blnSelected(intCounter)
			ReDim Preserve strWeighting(intCounter)
			arrCampaignID(intCounter)=rsCampaigns("CampaignID")
			strCampaignName(intCounter)=rsCampaigns("CompanyName") & ": " & rsCampaigns("CampaignName")
		Else
			blnDefaultCampaignExists=True
		End If
		rsCampaigns.MoveNext
	Loop
	If strTask="Edit" Then
		'find matching selected banners
		Do While Not rsZoneCampaigns.EOF
			intCounter=1
			If IsArray(arrCampaignID) Then
			Do While intCounter <= Ubound(arrCampaignID)
				If rsZoneCampaigns("CampaignID")=arrCampaignID(intCounter) And rsZoneCampaigns("Even")<>True Then
					blnSelected(intCounter)=True
					strWeighting(intCounter)=rsZoneCampaigns("ZoneCampaignWeighting")
					Exit Do
				End If
				intCounter=intCounter+1
			Loop
			End if
			rsZoneCampaigns.MoveNext
		Loop
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
			arrEvenCampaignID(intCounter)=rsEvenCampaigns("CampaignID")
			strEvenCampaignName(intCounter)=rsEvenCampaigns("CompanyName") & ": " & rsEvenCampaigns("CampaignName")
			blnEvenSelected(intCounter)=False
		End If
		rsEvenCampaigns.MoveNext
	Loop
	If strTask="Edit" Then
		'find matching selected banners
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
        <option <% If strTask="Edit" Then %><%If rsz("ZoneMode")="HTML" Then%>selected<%End If%><%Else%>selected<%End If%> value="HTML">HTML</option>
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
        <p align="left"><a href="help/zones.htm#EvenlyDistributed" target="_blank"><img border="0" src="images/evenlydistributedcampaigns.gif" WIDTH="586" HEIGHT="25">
        </a>
      </td>
</tr>
   
   

<tr>
      <td align="right" colspan="2">
        <p align="left"><font face="Arial" size="2"> The following campaigns are evenly distributed
        campaigns.&nbsp; These campaigns always take priority over weighted
        campaigns in order to ensure that all impressions will be delivered
        during the flight dates.&nbsp; The remaining inventory will be given to
        weighted campaigns.&nbsp; Use the CTRL key to select multiple campaigns.&nbsp;
        <a href="http://www.banmanpro.com/support/smoothing.asp">More
        information on smoothing algorithm.</a></font></td>
</tr>
    <tr align="center">
      <td align="right" colspan="2">
        <p align="center">&nbsp;
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="500" background="images/tableback.gif">
            <tr>
              <td align="center">
        <p align="center"><select size="6" name="EvenCampaigns" multiple>
<%
intCounter=1
If blnNoEvenCampaigns=False Then
Do While intCounter <= Ubound(arrEvenCampaignID)
%>
          <option <%If strTask="Edit" Then%><% If blnEvenSelected(intCounter)=True Then %>selected<%End If%><%End If%> value="<%=arrEvenCampaignID(intCounter)%>"><%=strEvenCampaignName(intCounter)%></option>
<%  intCounter=intCounter+1
Loop
End If %>
        </select></td>
    </tr>

  </center>

          </table>
        </div>
<%End If%>
    <tr align="center">
      <td align="right" colspan="2">
        <p align="left"><a href="help/zones.htm#Weighted" target="_blank"><img border="0" src="images/weightedcampaigns.gif" WIDTH="586" HEIGHT="25">
        </a>
      </td>
    </tr>
    <tr align="center">
      <td align="right" colspan="2">
      <p align="left"><font face="Arial" size="2">
      The following campaigns are weighted campaigns.&nbsp; The sum of all
      weighted campaigns should be 100 unless you have included a default.&nbsp;
      If the sum is less than 100 and you have not included a default, there
      will be times when a 1 X 1 pixel blank image is show.&nbsp;</font></td>
    </tr>
    <center>
    <tr align="center">
      <td width="578" align="right" colspan="2"><div align="center"><center><table border="1" cellpadding="2" cellspacing="0" width="467" bgcolor="#B6B6B6" bordercolor="#000000">
        <tr>
          <td align="center" width="314"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><strong><font face="Arial" size="2">Company:
            Campaign</font></strong></td>
          <td align="center" width="62"><font face="Arial" size="2"><strong>Selected</strong></font></td>
          <td align="center" width="71"><font face="Arial" size="2"><strong>Weighting</strong></font></td>
        </tr>
<% 
intCounter=1
If blnNoCampaigns=False Then
Do While intCounter <= Ubound(arrCampaignID)
	  %>
        <tr>
          <td align="center" width="314"><font face="Arial" size="2"><%=strCampaignName(intCounter)%>
            </font>
</td>
          <td align="center" width="62"><font face="Arial" size="2"><input type="checkbox" name="chkCampaignSelected<%=intCounter%>" value="<%=arrCampaignID(intCounter)%>" <%If strTask="Edit" Then%><% If blnSelected(intCounter)=True Then %>checked<%End If%><%End If%>></font></td>
          <td align="center" width="71"><font face="Arial" size="2"><input type="text" name="ZoneCampaignWeighting<%=intCounter%>" size="5" value="<%If strTask="Edit" Then%><% If blnSelected(intCounter)=True Then %><%=strWeighting(intCounter)%><%End If%><%End If%>"></font></td>
        </tr>
<%  intCounter=intCounter+1
Loop
End If %>
      </table>
      </center></div></td>
    </tr>
<% ' campaign site default *****************************************************************  %>
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
      <td width="391" align="left"><input type="submit" value="<%=strButtonText%>" name="B1"></td>
    </tr>
    <tr align="center">
      <td width="175" align="right">&nbsp;</td>
      <td width="391" align="left"><font face="Arial" size="2"><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Indicates Required Fields</font></td>
    </tr>
  </table>
  </center></div>
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


