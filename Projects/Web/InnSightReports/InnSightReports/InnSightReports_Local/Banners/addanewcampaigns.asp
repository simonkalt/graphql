<% 	

	If strTask="Edit" Then
		strCampaignID=Clng(strCampaignID)
		strAdvertiserID=Clng(strAdvertiserID)

		'determine if an overall default campaign exists
		strSQLL="SELECT * FROM Campaigns WHERE CampaignSiteDefault <>0"
		Set rsDefaultCampaign=connBanManPro.Execute(strSQLL)
		If rsDefaultCampaign.EOF=True Then
			blnFoundDefault=False
		Else
			blnFoundDefault=True
		End If
		blnFoundBanner=False
		strTask2="Update"
		strButtonText="Update Campaign"
		strAdvertiserID=rsc("AdvertiserID")
		strCompany=rsc("CompanyName")
		lngUserID=rsc("UserID")
		strSQLB="Select Banners.BannerID, Banners.UserID, Banners.AdvertiserID, Banners.AdTextLink,Banners.AdDescription, Banners.AdTargetURL, Banners.AdAltText, "
		strSQLB=strSQLB & " Banners.AdImageURL, Banners.AdBorder, Banners.AdWidth, Banners.AdHeight, Banners.AdAlign, Banners.AdNewWindow, Banners.AdTextUnderneath, Banners.AdFragment "
		strSQLB=strSQLB & " From Banners WHERE (((Banners.AdvertiserID)=" & strAdvertiserID & ") AND (Banners.UserID=" & CLng(Session("BanManProSiteID")) & " Or Banners.UserID=0) AND Banners.AdTextLink<>0) ORDER BY Banners.AdWidth,Banners.AdHeight DESC"

		Set rsBanner=connBanManPro.Execute(strSQLB)
		strSQLB="Select * From CampaignBanners WHERE CampaignBanners.CampaignID=" & strCampaignID
		Set rsCampaignBanners=connBanManPro.Execute(strSQLB)

		'create array of data
		intCounter=0
		Do While Not rsBanner.EOF
			blnFoundBanner=True
			intCounter=intCounter+1
			ReDim Preserve strBannerIDTemp(intCounter)
			ReDim Preserve strAdDescription(intCounter)
			ReDim Preserve strAdImageURL(intCounter)
			ReDim Preserve blnSelected(intCounter)
			ReDim Preserve strWeighting(intCounter)
			ReDim Preserve strSize(intCounter)
			strBannerIDTemp(intCounter)=rsBanner("BannerID")
			strAdDescription(intCounter)=rsBanner("AdDescription")
			strAdImageURL(intCounter)=rsBanner("AdImageURL")
			'strWeighting(intCounter)=0
			If rsBanner("AdTextLink")<>0 Then
				strSize(intCounter)="(Text Link)" 
			Else
				strSize(intCounter)="(" & rsBanner("AdWidth") & " X " & rsBanner("AdHeight") & ")"
			End If
			rsBanner.MoveNext
		Loop
		'find matching selected banners
		If blnFoundBanner=True Then
			Do While Not rsCampaignBanners.EOF
				intCounter=1
				Do While intCounter <= Ubound(strBannerIDTemp)
					If rsCampaignBanners("BannerID")=strBannerIDTemp(intCounter) Then
						blnSelected(intCounter)=True
						strWeighting(intCounter)=rsCampaignBanners("CampaignBannerWeighting")
						Exit Do
					End If
					intCounter=intCounter+1
				Loop
				rsCampaignBanners.MoveNext
			Loop
		End If
	Else
		strTask2="Insert"
		strButtonText="Submit New Campaign"
		'retrieve Company Name
		strSQL="SELECT * FROM Advertisers WHERE (((Advertisers.AdvertiserID)=" & Clng(Request.Form("AdvertiserID")) & ")) AND (UserID=" & CLng(Session("BanManProSiteID")) & " Or UserID=0)"
	   	strAdvertiserID=Request.Form("AdvertiserID")
		Set rsAdvertiser=connBanManPro.Execute(strSQL)
		strCompany=rsAdvertiser("CompanyName")
		lngUserID=rsAdvertiser("UserID")
		strSQLB="Select Banners.BannerID, Banners.UserID, Banners.AdvertiserID, Banners.AdTextLink, Banners.AdDescription, Banners.AdTargetURL, Banners.AdAltText, "
		strSQLB=strSQLB & " Banners.AdImageURL, Banners.AdBorder, Banners.AdWidth, Banners.AdHeight, Banners.AdAlign, Banners.AdNewWindow, Banners.AdTextUnderneath, Banners.AdFragment "
		strSQLB=strSQLB & " From Banners WHERE (((Banners.AdvertiserID)=" & Clng(Request.Form("AdvertiserID")) & ") AND (Banners.UserID=" & CLng(Session("BanManProSiteID")) & " Or Banners.UserID=0) AND Banners.AdTextLink<>0) ORDER BY Banners.AdWidth,Banners.AdHeight DESC"
		Set rsBanner=connBanManPro.Execute(strSQLB)
		If rsBanner.EOF=True Then
			Response.Write "You must add a text banner for this advertiser before you can define a campaign."
			Response.End
		End If
		If Clng(Application("StandardCampaignLength")) > 0 Then
			strDate=DateAdd("d",Clng(Application("StandardCampaignLength")),Date())
			strEndYear=Year(strDate)
			strEndMonth=Month(strDate)
			strEndDay=Day(strDate)
		Else
			strEndYear=Year(Date())
			strEndMonth=Month(Date())
			strEndDay=Day(Date())
		End If
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

  if (theForm.CampaignName.value == "")
  {
    alert("Please enter a value for the \"Campaign Name\" field.");
    theForm.CampaignName.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="campaigns.asp?Task=<%=strTask2%>&amp;CampaignID=<%=strCampaignID%>&amp;TextLink=True" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1">
  <input type="hidden" name="AdvertiserID" value="<%=strAdvertiserID%>">
  <div align="center">
    <center>
    <table border="0" cellpadding="0" cellspacing="0" width="590">
      <tr>
        <td><a href="help/campaigns.htm" target="_new"><img border="0" src="images/ListingofAllCampaigns.gif" WIDTH="590" HEIGHT="30"></a></td>
      </tr>
    </table>
    </center>
  </div>
  <div align="center"><center><table border="0" cellpadding="5" cellspacing="0" width="590" background="images/tableback.gif">
    <tr>
      <td width="134" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Advertiser:</strong></font></td>
      <td width="366"><font face="Arial" size="2"><%=strCompany%></font></td>
    </tr>
    <tr>
      <td width="134" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Campaign
      Name:</strong></font></td>
      <td width="366"><!--webbot bot="Validation" S-Display-Name="Campaign Name" B-Value-Required="TRUE" --><input type="text" name="CampaignName" size="45" <% If strTask="Edit" Then %>value="<%=rsc("CampaignName")%>" <%End If%>></td>
    </tr>
    <!--Multi-Site option only -->
    <% If Application("BanManProMultiSite")=True  And Clng(lngUserID)=0 Then%>
    <!--Multi-Site option only -->
    <tr>
      <td width="134" align="right"><font face="Arial" size="2"><strong>Run of
        Network:</strong></font></td>
      <td width="366"><input type="checkbox" name="RunOfNetwork" value="ON" <% If strTask="Edit" Then %><%If rsc("UserID")=0 Then%>checked<%End If%><%End If%>><font face="Arial" size="2">
        </font><font face="Arial" size="1">(Available to all sites if checked)</font></td>
    </tr>
    <!--Multi-Site option only -->
    <% End If %>
    <!--Multi-Site option only -->
    <tr>
      <td width="134" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Text
        Link:</strong></font></td>
      <td width="366"><div align="left"><table border="1" cellpadding="2" cellspacing="0" width="369" bordercolor="#000000">
        <tr>
          <td width="238" align="center"><strong><font face="Arial" size="2">Text
            Link</font></strong></td>
          <td width="58" align="center"><strong><font face="Arial" size="2">Selected</font></strong></td>
        </tr>
<% If strTask="Edit" Then 
intCounter=1
If blnFoundBanner=True Then
Do While intCounter <= Ubound(strBannerIDTemp)
%>
        <tr>
          <td width="238" align="center"><a href="<%=strAdImageURL(intCounter)%>"><font face="Arial" size="2"><%=strAdDescription(intCounter) & " " & strSize(intCounter)%></font></a></td>
          <td width="58" align="center"><font face="Arial" size="2"><input type="radio" <%If blnSelected(intCounter)=True Then%>checked<%End If%> value="<%=strBannerIDTemp(intCounter)%>" name="Banners"></font></td>
        </tr>
<% intCounter=intCounter+1
	Loop
End If %>
<% Else 
intCnt=0
Do While Not rsBanner.EOF 
intCnt=intCnt+1 %>
        <tr>
          <td width="238" align="center"><font face="Arial" size="2"><%=rsBanner("AdDescription")%></font></td>
          <td width="58" align="center"><font face="Arial" size="2"><input type="radio" value="<%=rsBanner("BannerID")%>" checked name="Banners"></font></td>
        </tr>
<% rsBanner.MoveNext
	Loop %>
<% End If %>
      </table>
      </div></td>
    </tr>
<% 'Default campaign ***********************************************************************
'Determine if a default exists
If strTask="Edit" Then 
'	If rsc("CampaignSiteDefault")=-1 Or rsc("CampaignSiteDefault")=1 Or blnFoundDefault=False Then
'	If rsc("CampaignSiteDefault")=-1 Or rsc("CampaignSiteDefault")=1 Then
%>
    <%'   End If 
Else 
'If blnFoundDefault=False Then%><%' End If
End If 
'End Default campaign ******************************************************************** %>
    <tr>
      <td width="134" align="right"><font face="Arial" size="2"><strong>Cost:</strong></font></td>
      <td width="366"><input type="text" name="CampaignCost" size="10" <% If strTask="Edit" Then %>value="<%=rsc("CampaignCost")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="134" align="right">&nbsp;</td>
      <td width="366"><font face="Arial" size="2"><input type="submit" value="<%=strButtonText%>" name="B1"></font></td>
    </tr>
    <tr>
      <td width="134" align="right">&nbsp;</td>
      <td width="366"><font face="Arial" size="2"><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Indicates Required Fields</font></td>
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
  <input type="hidden" name="CampaignDailyEnd" value="12:00:00 AM"><input type="hidden" name="CampaignDailyStart" value="12:00:00 AM"><input type="hidden" name="CampaignDistribution" value="Text"><input type="hidden" name="CampaignEndDate" value="01/01/2099"><input type="hidden" name="CampaignQuantitySold" value="0"><input type="hidden" name="CampaignSiteDefault" value="0"><input type="hidden" name="CampaignStartDate" value="01/01/1999"><input type="hidden" name="CampaignType" value="Flat Rate">
</form>

</body>
</html>

<%

Set rsDefaultCampaign=Nothing
Set rsBanner=Nothing
Set rsCampaignBanners=Nothing
Set rsAdvertiser=Nothing

%>