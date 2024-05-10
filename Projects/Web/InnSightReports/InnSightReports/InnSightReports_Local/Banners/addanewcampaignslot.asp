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
		lngUserID=rsc("UserID")
		strCompany=rsc("CompanyName")
		strSQLB="Select Banners.BannerID, Banners.UserID, Banners.AdvertiserID, Banners.AdTextLink,Banners.AdDescription, Banners.AdTargetURL, Banners.AdAltText, "
		strSQLB=strSQLB & " Banners.AdImageURL, Banners.AdBorder, Banners.AdWidth, Banners.AdHeight, Banners.AdAlign, Banners.AdNewWindow, Banners.AdTextUnderneath, Banners.AdFragment "
		strSQLB=strSQLB & " From Banners WHERE (((Banners.AdvertiserID)=" & strAdvertiserID & ") AND (Banners.UserID=" & CLng(Session("BanManProSiteID")) & " Or Banners.UserID=0) AND Banners.AdTextLink<>1) ORDER BY Banners.AdWidth,Banners.AdHeight DESC"

		Set rsBanner=connBanManPro.Execute(strSQLB)
		strSQLB="Select * From CampaignBanners WHERE CampaignBanners.CampaignID=" & strCampaignID
		Set rsCampaignBanners=connBanManPro.Execute(strSQLB)

		'Obtain list of zones this campaign is included in
		strSQL="SELECT Zones.ZoneDescription, ZoneCampaigns.ZoneID FROM ZoneCampaigns INNER JOIN Zones ON ZoneCampaigns.ZoneID = Zones.ZoneID Where ZoneCampaigns.CampaignID="  & strCampaignID & " AND (ZoneCampaigns.UserID=" & CLng(Session("BanManProSiteID")) & " Or ZoneCampaigns.UserID=0)  Order By ZoneDescription ASC"
		Set rsZoneList=connBanManpro.Execute(strSQL)

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
		strSQLB=strSQLB & " From Banners WHERE (((Banners.AdvertiserID)=" & Clng(Request.Form("AdvertiserID")) & ") AND (Banners.UserID=" & CLng(Session("BanManProSiteID")) & " Or Banners.UserID=0) AND Banners.AdTextLink<>1) ORDER BY Banners.AdWidth,Banners.AdHeight DESC"
		Set rsBanner=connBanManPro.Execute(strSQLB)
		If rsBanner.EOF=True Then
			Response.Write "You must add a banner for this advertiser before you can define a campaign."
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
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="campaigns.asp?Task=<%=strTask2%>&amp;CampaignID=<%=strCampaignID%>" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1">
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
  <div align="center"><center><table border="0" cellpadding="4" cellspacing="0" width="590" background="images/tableback.gif">
    <tr>
      <td width="134" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Advertiser:</strong></font></td>
      <td width="366"><font face="Arial" size="2"><%=strCompany%></font></td>
    </tr>
    <tr>
      <td width="134" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Campaign
      Name:</strong></font></td>
      <td width="366"><!--webbot bot="Validation" S-Display-Name="Campaign Name" B-Value-Required="TRUE" --><input type="text" name="CampaignName" size="45" <% If strTask="Edit" Then %>value="<%=rsc("CampaignName")%>" <%Else%>value="<%=strCompany%>" <%End If%>></td>
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
      <td width="134" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Banner(s):</strong></font></td>
      <td width="366"><div align="left"><table border="1" cellpadding="2" cellspacing="0" width="369" bordercolor="#000000">
        <tr>
          <td width="238" align="center"><strong><font face="Arial" size="2">Banner Name
            (Size)</font></strong></td>
          <td width="58" align="center"><strong><font face="Arial" size="2">Selected</font></strong></td>
          <td width="73" align="center"><font face="Arial" size="2"><strong>Weighting</strong></font></td>
        </tr>
<% If strTask="Edit" Then 
intCounter=1
If blnFoundBanner=True Then
Do While intCounter <= Ubound(strBannerIDTemp)
%>
        <tr>
          <td width="238" align="center"><a href="<%=strAdImageURL(intCounter)%>"><font face="Arial" size="2"><%=strAdDescription(intCounter) & " " & strSize(intCounter)%></font></a></td>
          <td width="58" align="center"><font face="Arial" size="2"><input type="checkbox" name="chkBannerSelected<%=intCounter%>" value="<%=strBannerIDTemp(intCounter)%>" <%If strWeighting(intCounter)<>0 Then %>checked<%End IF%>></font></td>
          <td width="73" align="center"><font face="Arial" size="2"><input type="text" name="txtBannerWeighting<%=intCounter%>" size="6" value="<%=strWeighting(intCounter)%>"></font></td>
        </tr>
<% intCounter=intCounter+1
	Loop
End If %>
<% Else 
intCnt=0
Do While Not rsBanner.EOF 
If rsBanner("AdTextLink")=True Then
	strTemp="Text Link"
Else
	strTemp= rsBanner("AdWidth") & " X " & rsBanner("AdHeight") 
End If
intCnt=intCnt+1 %>
        <tr>
          <td width="238" align="center"><font face="Arial" size="2"><%=rsBanner("AdDescription") & " (" & strTemp & ")"%></font></td>
          <td width="58" align="center"><font face="Arial" size="2"><input type="checkbox" name="chkBannerSelected<%=intCnt%>" value="<%=rsBanner("BannerID")%>"></font></td>
          <td width="73" align="center"><font face="Arial" size="2"><input type="text" name="txtBannerWeighting<%=intCnt%>" size="6"></font></td>
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
    <tr>
      <td width="134" align="right"><font face="Arial" size="2"><strong>Default Campaign:</strong></font></td>
      <td width="366"><font face="Arial" size="2"><input type="checkbox" name="CampaignSiteDefault" value="-1" <% If strTask="Edit" Then %><% If rsc("CampaignSiteDefault")=True Then%>checked<%End If%><%Else%><%End If%>> </font><font face="Arial" size="1">(Note: All
      Parameters Below will be ignored for default campaigns)</font><font face="Arial" size="2"></font></td>
    </tr>
<%'   End If 
Else 
'If blnFoundDefault=False Then%>
    <tr>
      <td width="134" align="right"><font face="Arial" size="2"><strong>Default Campaign:</strong></font></td>
      <td width="366"><font face="Arial" size="2"><input type="checkbox" name="CampaignSiteDefault" value="-1"> </font><font face="Arial" size="1">(Note: All
      Parameters Below will be ignored for default campaigns)</font><font face="Arial" size="2"></font></td>
    </tr>
<%' End If
End If 
'End Default campaign ******************************************************************** %>
  </center>
    <tr>
      <td align="right" colspan="2">
        <p align="left"><a href="help/campaigns.htm#datetime" target="_blank"><img border="0" src="images/datetimeinformation.gif" WIDTH="586" HEIGHT="25">
        </a>
      </td>
    </tr>
    <center>
    <tr>
      <td width="134" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Start Date:</strong></font></td>
      <td width="366"><font face="Arial" size="2">Year: <select name="StartYear" size="1">
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignStartDate"))="1999" Then%>selected<%End If %><% End If%> value="1999">1999</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignStartDate"))="2000" Then%>selected<%End If %><%Else%><%If Year(Date())="2000" Then%>selected<%End If%><% End If%> value="2000">2000</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignStartDate"))="2001" Then%>selected<%End If %><%Else%><%If Year(Date())="2001" Then%>selected<%End If%><% End If%> value="2001">2001</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignStartDate"))="2002" Then%>selected<%End If %><%Else%><%If Year(Date())="2002" Then%>selected<%End If%><% End If%> value="2002">2002</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignStartDate"))="2003" Then%>selected<%End If %><%Else%><%If Year(Date())="2003" Then%>selected<%End If%><% End If%> value="2003">2003</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignStartDate"))="2004" Then%>selected<%End If %><%Else%><%If Year(Date())="2004" Then%>selected<%End If%><% End If%> value="2004">2004</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignStartDate"))="2005" Then%>selected<%End If %><%Else%><%If Year(Date())="2005" Then%>selected<%End If%><% End If%> value="2005">2005</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignStartDate"))="2006" Then%>selected<%End If %><%Else%><%If Year(Date())="2006" Then%>selected<%End If%><% End If%> value="2006">2006</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignStartDate"))="2007" Then%>selected<%End If %><%Else%><%If Year(Date())="2007" Then%>selected<%End If%><% End If%> value="2007">2007</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignStartDate"))="2008" Then%>selected<%End If %><%Else%><%If Year(Date())="2008" Then%>selected<%End If%><% End If%> value="2008">2008</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignStartDate"))="2009" Then%>selected<%End If %><%Else%><%If Year(Date())="2009" Then%>selected<%End If%><% End If%> value="2009">2009</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignStartDate"))="2010" Then%>selected<%End If %><%Else%><%If Year(Date())="2010" Then%>selected<%End If%><% End If%> value="2010">2010</option>
      </select>&nbsp; Month:<select name="StartMonth" size="1">
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignStartDate"))="1" Then%>selected<%End If %><%Else%><%If Month(Date())="1" Then%>selected<%End If%><% End If%> value="01">Jan</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignStartDate"))="2" Then%>selected<%End If %><%Else%><%If Month(Date())="2" Then%>selected<%End If%><% End If%> value="02">Feb</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignStartDate"))="3" Then%>selected<%End If %><%Else%><%If Month(Date())="3" Then%>selected<%End If%><% End If%> value="03">Mar</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignStartDate"))="4" Then%>selected<%End If %><%Else%><%If Month(Date())="4" Then%>selected<%End If%><% End If%> value="04">Apr</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignStartDate"))="5" Then%>selected<%End If %><%Else%><%If Month(Date())="5" Then%>selected<%End If%><% End If%> value="05">May</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignStartDate"))="6" Then%>selected<%End If %><%Else%><%If Month(Date())="6" Then%>selected<%End If%><% End If%> value="06">Jun</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignStartDate"))="7" Then%>selected<%End If %><%Else%><%If Month(Date())="7" Then%>selected<%End If%><% End If%> value="07">Jul</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignStartDate"))="8" Then%>selected<%End If %><%Else%><%If Month(Date())="8" Then%>selected<%End If%><% End If%> value="08">Aug</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignStartDate"))="9" Then%>selected<%End If %><%Else%><%If Month(Date())="9" Then%>selected<%End If%><% End If%> value="09">Sep</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignStartDate"))="10" Then%>selected<%End If %><%Else%><%If Month(Date())="10" Then%>selected<%End If%><% End If%> value="10">Oct</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignStartDate"))="11" Then%>selected<%End If %><%Else%><%If Month(Date())="11" Then%>selected<%End If%><% End If%> value="11">Nov</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignStartDate"))="12" Then%>selected<%End If %><%Else%><%If Month(Date())="12" Then%>selected<%End If%><% End If%> value="12">Dec</option>
      </select>&nbsp; Day:<select name="StartDay" size="1">
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="1" Then%>selected<%End If %><%Else%><%If Day(Date())="1" Then%>selected<%End If%><% End If%> value="1">1</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="2" Then%>selected<%End If %><%Else%><%If Day(Date())="2" Then%>selected<%End If%><% End If%> value="2">2</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="3" Then%>selected<%End If %><%Else%><%If Day(Date())="3" Then%>selected<%End If%><% End If%> value="3">3</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="4" Then%>selected<%End If %><%Else%><%If Day(Date())="4" Then%>selected<%End If%><% End If%> value="4">4</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="5" Then%>selected<%End If %><%Else%><%If Day(Date())="5" Then%>selected<%End If%><% End If%> value="5">5</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="6" Then%>selected<%End If %><%Else%><%If Day(Date())="6" Then%>selected<%End If%><% End If%> value="6">6</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="7" Then%>selected<%End If %><%Else%><%If Day(Date())="7" Then%>selected<%End If%><% End If%> value="7">7</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="8" Then%>selected<%End If %><%Else%><%If Day(Date())="8" Then%>selected<%End If%><% End If%> value="8">8</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="9" Then%>selected<%End If %><%Else%><%If Day(Date())="9" Then%>selected<%End If%><% End If%> value="9">9</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="10" Then%>selected<%End If %><%Else%><%If Day(Date())="10" Then%>selected<%End If%><% End If%> value="10">10</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="11" Then%>selected<%End If %><%Else%><%If Day(Date())="11" Then%>selected<%End If%><% End If%> value="11">11</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="12" Then%>selected<%End If %><%Else%><%If Day(Date())="12" Then%>selected<%End If%><% End If%> value="12">12</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="13" Then%>selected<%End If %><%Else%><%If Day(Date())="13" Then%>selected<%End If%><% End If%> value="13">13</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="14" Then%>selected<%End If %><%Else%><%If Day(Date())="14" Then%>selected<%End If%><% End If%> value="14">14</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="15" Then%>selected<%End If %><%Else%><%If Day(Date())="15" Then%>selected<%End If%><% End If%> value="15">15</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="16" Then%>selected<%End If %><%Else%><%If Day(Date())="16" Then%>selected<%End If%><% End If%> value="16">16</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="17" Then%>selected<%End If %><%Else%><%If Day(Date())="17" Then%>selected<%End If%><% End If%> value="17">17</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="18" Then%>selected<%End If %><%Else%><%If Day(Date())="18" Then%>selected<%End If%><% End If%> value="18">18</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="19" Then%>selected<%End If %><%Else%><%If Day(Date())="19" Then%>selected<%End If%><% End If%> value="19">19</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="20" Then%>selected<%End If %><%Else%><%If Day(Date())="20" Then%>selected<%End If%><% End If%> value="20">20</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="21" Then%>selected<%End If %><%Else%><%If Day(Date())="21" Then%>selected<%End If%><% End If%> value="21">21</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="22" Then%>selected<%End If %><%Else%><%If Day(Date())="22" Then%>selected<%End If%><% End If%> value="22">22</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="23" Then%>selected<%End If %><%Else%><%If Day(Date())="23" Then%>selected<%End If%><% End If%> value="23">23</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="24" Then%>selected<%End If %><%Else%><%If Day(Date())="24" Then%>selected<%End If%><% End If%> value="24">24</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="25" Then%>selected<%End If %><%Else%><%If Day(Date())="25" Then%>selected<%End If%><% End If%> value="25">25</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="26" Then%>selected<%End If %><%Else%><%If Day(Date())="26" Then%>selected<%End If%><% End If%> value="26">26</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="27" Then%>selected<%End If %><%Else%><%If Day(Date())="27" Then%>selected<%End If%><% End If%> value="27">27</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="28" Then%>selected<%End If %><%Else%><%If Day(Date())="28" Then%>selected<%End If%><% End If%> value="28">28</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="29" Then%>selected<%End If %><%Else%><%If Day(Date())="29" Then%>selected<%End If%><% End If%> value="29">29</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="30" Then%>selected<%End If %><%Else%><%If Day(Date())="30" Then%>selected<%End If%><% End If%> value="30">30</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignStartDate"))="31" Then%>selected<%End If %><%Else%><%If Day(Date())="31" Then%>selected<%End If%><% End If%> value="31">31</option>
      </select></font></td>
    </tr>
    <tr>
      <td width="134" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>End Date:</strong></font></td>
      <td width="366"><font face="Arial" size="2">Year: <select name="EndYear" size="1">
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignEndDate"))="1999" Then%>selected<%End If %><% End If%> value="1999">1999</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignEndDate"))="2000" Then%>selected<%End If %><%Else%><%If strEndYear="2000" Then%>selected<%End If%><% End If%> value="2000">2000</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignEndDate"))="2001" Then%>selected<%End If %><%Else%><%If strEndYear="2001" Then%>selected<%End If%><% End If%> value="2001">2001</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignEndDate"))="2002" Then%>selected<%End If %><%Else%><%If strEndYear="2002" Then%>selected<%End If%><% End If%> value="2002">2002</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignEndDate"))="2003" Then%>selected<%End If %><%Else%><%If strEndYear="2003" Then%>selected<%End If%><% End If%> value="2003">2003</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignEndDate"))="2004" Then%>selected<%End If %><%Else%><%If strEndYear="2004" Then%>selected<%End If%><% End If%> value="2004">2004</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignEndDate"))="2005" Then%>selected<%End If %><%Else%><%If strEndYear="2005" Then%>selected<%End If%><% End If%> value="2005">2005</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignEndDate"))="2006" Then%>selected<%End If %><%Else%><%If strEndYear="2006" Then%>selected<%End If%><% End If%> value="2006">2006</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignEndDate"))="2007" Then%>selected<%End If %><%Else%><%If strEndYear="2007" Then%>selected<%End If%><% End If%> value="2007">2007</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignEndDate"))="2008" Then%>selected<%End If %><%Else%><%If strEndYear="2008" Then%>selected<%End If%><% End If%> value="2008">2008</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignEndDate"))="2009" Then%>selected<%End If %><%Else%><%If strEndYear="2009" Then%>selected<%End If%><% End If%> value="2009">2009</option>
        <option <% If strTask="Edit" Then %> <% If Year(rsc("CampaignEndDate"))="2010" Then%>selected<%End If %><%Else%><%If strEndYear="2010" Then%>selected<%End If%><% End If%> value="2010">2010</option>
      </select>&nbsp; Month:<select name="EndMonth" size="1">
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignEndDate"))="1" Then%>selected<%End If %><%Else%><%If strEndMonth="1" Then%>selected<%End If%><% End If%> value="01">Jan</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignEndDate"))="2" Then%>selected<%End If %><%Else%><%If strEndMonth="2" Then%>selected<%End If%><% End If%> value="02">Feb</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignEndDate"))="3" Then%>selected<%End If %><%Else%><%If strEndMonth="3" Then%>selected<%End If%><% End If%> value="03">Mar</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignEndDate"))="4" Then%>selected<%End If %><%Else%><%If strEndMonth="4" Then%>selected<%End If%><% End If%> value="04">Apr</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignEndDate"))="5" Then%>selected<%End If %><%Else%><%If strEndMonth="5" Then%>selected<%End If%><% End If%> value="05">May</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignEndDate"))="6" Then%>selected<%End If %><%Else%><%If strEndMonth="6" Then%>selected<%End If%><% End If%> value="06">Jun</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignEndDate"))="7" Then%>selected<%End If %><%Else%><%If strEndMonth="7" Then%>selected<%End If%><% End If%> value="07">Jul</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignEndDate"))="8" Then%>selected<%End If %><%Else%><%If strEndMonth="8" Then%>selected<%End If%><% End If%> value="08">Aug</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignEndDate"))="9" Then%>selected<%End If %><%Else%><%If strEndMonth="9" Then%>selected<%End If%><% End If%> value="09">Sep</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignEndDate"))="10" Then%>selected<%End If %><%Else%><%If strEndMonth="10" Then%>selected<%End If%><% End If%> value="10">Oct</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignEndDate"))="11" Then%>selected<%End If %><%Else%><%If strEndMonth="11" Then%>selected<%End If%><% End If%> value="11">Nov</option>
        <option <% If strTask="Edit" Then %> <% If Month(rsc("CampaignEndDate"))="12" Then%>selected<%End If %><%Else%><%If strEndMonth="12" Then%>selected<%End If%><% End If%> value="12">Dec</option>
      </select>&nbsp; Day:<select name="EndDay" size="1">
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="1" Then%>selected<%End If %><%Else%><%If strEndDay="1" Then%>selected<%End If%><% End If%> value="1">1</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="2" Then%>selected<%End If %><%Else%><%If strEndDay="2" Then%>selected<%End If%><% End If%> value="2">2</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="3" Then%>selected<%End If %><%Else%><%If strEndDay="3" Then%>selected<%End If%><% End If%> value="3">3</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="4" Then%>selected<%End If %><%Else%><%If strEndDay="4" Then%>selected<%End If%><% End If%> value="4">4</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="5" Then%>selected<%End If %><%Else%><%If strEndDay="5" Then%>selected<%End If%><% End If%> value="5">5</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="6" Then%>selected<%End If %><%Else%><%If strEndDay="6" Then%>selected<%End If%><% End If%> value="6">6</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="7" Then%>selected<%End If %><%Else%><%If strEndDay="7" Then%>selected<%End If%><% End If%> value="7">7</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="8" Then%>selected<%End If %><%Else%><%If strEndDay="8" Then%>selected<%End If%><% End If%> value="8">8</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="9" Then%>selected<%End If %><%Else%><%If strEndDay="9" Then%>selected<%End If%><% End If%> value="9">9</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="10" Then%>selected<%End If %><%Else%><%If strEndDay="10" Then%>selected<%End If%><% End If%> value="10">10</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="11" Then%>selected<%End If %><%Else%><%If strEndDay="11" Then%>selected<%End If%><% End If%> value="11">11</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="12" Then%>selected<%End If %><%Else%><%If strEndDay="12" Then%>selected<%End If%><% End If%> value="12">12</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="13" Then%>selected<%End If %><%Else%><%If strEndDay="13" Then%>selected<%End If%><% End If%> value="13">13</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="14" Then%>selected<%End If %><%Else%><%If strEndDay="14" Then%>selected<%End If%><% End If%> value="14">14</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="15" Then%>selected<%End If %><%Else%><%If strEndDay="15" Then%>selected<%End If%><% End If%> value="15">15</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="16" Then%>selected<%End If %><%Else%><%If strEndDay="16" Then%>selected<%End If%><% End If%> value="16">16</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="17" Then%>selected<%End If %><%Else%><%If strEndDay="17" Then%>selected<%End If%><% End If%> value="17">17</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="18" Then%>selected<%End If %><%Else%><%If strEndDay="18" Then%>selected<%End If%><% End If%> value="18">18</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="19" Then%>selected<%End If %><%Else%><%If strEndDay="19" Then%>selected<%End If%><% End If%> value="19">19</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="20" Then%>selected<%End If %><%Else%><%If strEndDay="20" Then%>selected<%End If%><% End If%> value="20">20</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="21" Then%>selected<%End If %><%Else%><%If strEndDay="21" Then%>selected<%End If%><% End If%> value="21">21</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="22" Then%>selected<%End If %><%Else%><%If strEndDay="22" Then%>selected<%End If%><% End If%> value="22">22</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="23" Then%>selected<%End If %><%Else%><%If strEndDay="23" Then%>selected<%End If%><% End If%> value="23">23</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="24" Then%>selected<%End If %><%Else%><%If strEndDay="24" Then%>selected<%End If%><% End If%> value="24">24</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="25" Then%>selected<%End If %><%Else%><%If strEndDay="25" Then%>selected<%End If%><% End If%> value="25">25</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="26" Then%>selected<%End If %><%Else%><%If strEndDay="26" Then%>selected<%End If%><% End If%> value="26">26</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="27" Then%>selected<%End If %><%Else%><%If strEndDay="27" Then%>selected<%End If%><% End If%> value="27">27</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="28" Then%>selected<%End If %><%Else%><%If strEndDay="28" Then%>selected<%End If%><% End If%> value="28">28</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="29" Then%>selected<%End If %><%Else%><%If strEndDay="29" Then%>selected<%End If%><% End If%> value="29">29</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="30" Then%>selected<%End If %><%Else%><%If strEndDay="30" Then%>selected<%End If%><% End If%> value="30">30</option>
        <option <% If strTask="Edit" Then %> <% If Day(rsc("CampaignEndDate"))="31" Then%>selected<%End If %><%Else%><%If strEndDay="31" Then%>selected<%End If%><% End If%> value="31">31</option>
      </select></font></td>
    </tr>
    <tr>
      <td width="134" align="right"><font face="Arial" size="2"><strong>Daily Start Time:</strong></font></td>
      <td width="366"><font face="Arial" size="2"><select name="CampaignDailyStart" size="1">
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="12:00:00 AM" Then%>selected<%End If %><% End If%> value="12:00 AM">12:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="1:00:00 AM" Then%>selected<%End If %><% End If%> value="1:00 AM">1:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="2:00:00 AM" Then%>selected<%End If %><% End If%> value="2:00 AM">2:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="3:00:00 AM" Then%>selected<%End If %><% End If%> value="3:00 AM">3:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="4:00:00 AM" Then%>selected<%End If %><% End If%> value="4:00 AM">4:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="5:00:00 AM" Then%>selected<%End If %><% End If%> value="5:00 AM">5:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="6:00:00 AM" Then%>selected<%End If %><% End If%> value="6:00 AM">6:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="7:00:00 AM" Then%>selected<%End If %><% End If%> value="7:00 AM">7:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="8:00:00 AM" Then%>selected<%End If %><% End If%> value="8:00 AM">8:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="9:00:00 AM" Then%>selected<%End If %><% End If%> value="9:00 AM">9:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="10:00:00 AM" Then%>selected<%End If %><% End If%> value="10:00 AM">10:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="11:00:00 AM" Then%>selected<%End If %><% End If%> value="11:00 AM">11:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="12:00:00 PM" Then%>selected<%End If %><% End If%> value="12:00 PM">12:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="1:00:00 PM" Then%>selected<%End If %><% End If%> value="1:00 PM">1:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="2:00:00 PM" Then%>selected<%End If %><% End If%> value="2:00 PM">2:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="3:00:00 PM" Then%>selected<%End If %><% End If%> value="3:00 PM">3:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="4:00:00 PM" Then%>selected<%End If %><% End If%> value="4:00 PM">4:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="5:00:00 PM" Then%>selected<%End If %><% End If%> value="5:00 PM">5:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="6:00:00 PM" Then%>selected<%End If %><% End If%> value="6:00 PM">6:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="7:00:00 PM" Then%>selected<%End If %><% End If%> value="7:00 PM">7:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="8:00:00 PM" Then%>selected<%End If %><% End If%> value="8:00 PM">8:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="9:00:00 PM" Then%>selected<%End If %><% End If%> value="9:00 PM">9:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="10:00:00 PM" Then%>selected<%End If %><% End If%> value="10:00 PM">10:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)="11:00:00 PM" Then%>selected<%End If %><% End If%> value="11:00 PM">11:00 PM</option>
      </select></font></td>
    </tr>
    <tr>
      <td width="134" align="right"><font face="Arial" size="2"><strong>Daily End Time:</strong></font></td>
      <td><font face="Arial" size="2"><select name="CampaignDailyEnd" size="1">
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="12:00:00 AM" Then%>selected<%End If %><% End If%> value="12:00 AM">12:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="1:00:00 AM" Then%>selected<%End If %><% End If%> value="1:00 AM">1:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="2:00:00 AM" Then%>selected<%End If %><% End If%> value="2:00 AM">2:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="3:00:00 AM" Then%>selected<%End If %><% End If%> value="3:00 AM">3:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="4:00:00 AM" Then%>selected<%End If %><% End If%> value="4:00 AM">4:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="5:00:00 AM" Then%>selected<%End If %><% End If%> value="5:00 AM">5:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="6:00:00 AM" Then%>selected<%End If %><% End If%> value="6:00 AM">6:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="7:00:00 AM" Then%>selected<%End If %><% End If%> value="7:00 AM">7:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="8:00:00 AM" Then%>selected<%End If %><% End If%> value="8:00 AM">8:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="9:00:00 AM" Then%>selected<%End If %><% End If%> value="9:00 AM">9:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="10:00:00 AM" Then%>selected<%End If %><% End If%> value="10:00 AM">10:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="11:00:00 AM" Then%>selected<%End If %><% End If%> value="11:00 AM">11:00 AM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="12:00:00 PM" Then%>selected<%End If %><% End If%> value="12:00 PM">12:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="1:00:00 PM" Then%>selected<%End If %><% End If%> value="1:00 PM">1:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="2:00:00 PM" Then%>selected<%End If %><% End If%> value="2:00 PM">2:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="3:00:00 PM" Then%>selected<%End If %><% End If%> value="3:00 PM">3:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="4:00:00 PM" Then%>selected<%End If %><% End If%> value="4:00 PM">4:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="5:00:00 PM" Then%>selected<%End If %><% End If%> value="5:00 PM">5:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="6:00:00 PM" Then%>selected<%End If %><% End If%> value="6:00 PM">6:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="7:00:00 PM" Then%>selected<%End If %><% End If%> value="7:00 PM">7:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="8:00:00 PM" Then%>selected<%End If %><% End If%> value="8:00 PM">8:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="9:00:00 PM" Then%>selected<%End If %><% End If%> value="9:00 PM">9:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="10:00:00 PM" Then%>selected<%End If %><% End If%> value="10:00 PM">10:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="11:00:00 PM" Then%>selected<%End If %><% End If%> value="11:00 PM">11:00 PM</option>
        <option <% If strTask="Edit" Then %> <% If FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)="11:59:00 PM" Then%>selected<%End If %><% End If%> value="11:59 PM">11:59 PM</option>
	</select></font><font face="Arial" size="1"> (Setting both equal to 12:00
        AM indicates
      24-hour rotation)</font></td>
    </tr>
    <tr>
      <td width="134" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Days Selected:</strong></font></td>
      <td width="366"><div align="left"><table border="0" cellpadding="0" cellspacing="0" width="413">
        <tr>
          <td width="157"><font face="Arial" size="2"><input type="checkbox" name="CampaignSunday" value="-1" <% If strTask="Edit" Then %><% If rsc("CampaignSunday")=True Then%>checked<%End If%><%Else%>checked<%End If%>>Sunday</font></td>
          <td width="256"><font face="Arial" size="2"><input type="checkbox" name="CampaignThursday" value="-1" <% If strTask="Edit" Then %><% If rsc("CampaignThursday")=True Then%>checked<%End If%><%Else%>checked<%End If%>>Thursday</font></td>
        </tr>
        <tr>
          <td width="157"><font face="Arial" size="2"><input type="checkbox" name="CampaignMonday" value="-1" <% If strTask="Edit" Then %><% If rsc("CampaignMonday")=True Then%>checked<%End If%><%Else%>checked<%End If%>>Monday</font></td>
          <td width="256"><font face="Arial" size="2"><input type="checkbox" name="CampaignFriday" value="-1" <% If strTask="Edit" Then %><% If rsc("CampaignFriday")=True Then%>checked<%End If%><%Else%>checked<%End If%>>Friday</font></td>
        </tr>
        <tr>
          <td width="157"><font face="Arial" size="2"><input type="checkbox" name="CampaignTuesday" value="-1" <% If strTask="Edit" Then %><% If rsc("CampaignTuesday")=True Then%>checked<%End If%><%Else%>checked<%End If%>>Tuesday</font></td>
          <td width="256"><font face="Arial" size="2"><input type="checkbox" name="CampaignSaturday" value="-1" <% If strTask="Edit" Then %><% If rsc("CampaignSaturday")=True Then%>checked<%End If%><%Else%>checked<%End If%>>Saturday</font></td>
        </tr>
        <tr>
          <td width="157"><font face="Arial" size="2"><input type="checkbox" name="CampaignWednesday" value="-1" <% If strTask="Edit" Then %><% If rsc("CampaignWednesday")=True Then%>checked<%End If%><%Else%>checked<%End If%>>Wednesday</font></td>
          <td width="256"></td>
        </tr>
      </table>
      </div></td>
    </tr>
  </center>
    <tr>
      <td align="right" colspan="2">
        <p align="left"><a href="help/campaigns.htm#distribution" target="_blank"><img border="0" src="images/Distribution_OtherInformation.gif" WIDTH="586" HEIGHT="25">
        </a>
      </td>
    </tr>
  </table>
    <table border="0" cellpadding="4" cellspacing="0" width="594" background="images/tableback.gif">
      <tr>
        <td width="134" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Distribution:</strong></font></td>
        <td width="366"><font face="Arial" size="2"><select name="CampaignDistribution" size="1">
            <option <% If strTask="Edit" Then %> <% If rsc("CampaignDistribution")="Weighted" Then%>selected<%End If %><% End If%>value="Weighted" value="Weighted">Slot</option>
            <option <% If strTask="Edit" Then %> <% If rsc("CampaignDistribution")="Keyword" Then%>selected<%End If %><% End If%> value="Keyword">Called
            By Keyword Only</option>
          </select></font></td>
      </tr>
    </table>
    <table border="0" cellpadding="4" cellspacing="0" width="590" background="images/tableback.gif">
    <center>
    <tr>
      <td width="134" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><strong><font face="Arial" size="2">Quantity
        or Slots:</font></strong></td>
      <td width="366"><input type="text" name="CampaignQuantitySold" size="10" <% If strTask="Edit" Then %>value="<%=rsc("CampaignQuantitySold")%>" <%End If%>>
        &nbsp;</td>
    </tr>
    <tr>
      <td width="134" align="right"><font face="Arial" size="2"><strong>Cost:</strong></font></td>
      <td width="366"><input type="text" name="CampaignCost" size="10" <% If strTask="Edit" Then %>value="<%=rsc("CampaignCost")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="134" align="right" valign="top"><font face="Arial" size="2"><strong>Optional
        Keywords:</strong></font></td>
      <td width="366"><input type="text" name="CampaignKeywords" size="50" <% If strTask="Edit" Then %>value="<%=rsc("CampaignKeywords")%>" <%End If%>><br>
        <font face="Arial" size="1">Note: Use comma&nbsp; to separate multiple
        words/phrases.&nbsp; This field is reserved for campaigns of
        distribution type &quot;Call by Keyword Only&quot;.</font></td>
    </tr>
    <tr>
      <td width="134" align="right">&nbsp;</td>
      <td width="366"><font face="Arial" size="2"><input type="submit" value="<%=strButtonText%>" name="B1"></font></td>
    </tr>
    <tr>
      <td width="134" align="right">&nbsp;</td>
      <td width="366"><font face="Arial" size="2"><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Indicates Required Fields</font></td>
    </tr>
<!--Editing Campaign so show list of zones-->
<% If  strTask="Edit" Then %>
    <tr>
      <td width="500" align="right" colspan="2">
        <p align="center"><font face="Arial" size="3"><b><img border="0" src="images/zonesdisplayingcampaignbar.gif" WIDTH="586" HEIGHT="25"></b></font></td>
    </tr>
    <tr>
      <td align="center" colspan="2">
        <p align="center"><font face="Arial" size="2">
  <table border="1" cellpadding="2" cellspacing="0" width="301" bordercolor="#000000">
  <%Do While Not rsZoneList.EOF %>
    <tr>
      <td width="38"><a href="zones.asp?Task=Edit&amp;ZoneID=<%=rsZoneList("ZoneID")%>"><img border="0" src="images/Editsmall.gif" width="38" height="18"></a></td>
      <td width="259">
        <p align="center"><font face="Arial" size="3"><%=rsZoneList("ZoneDescription")%></font></td>
    </tr>
    	<%rsZoneList.MoveNext
	Loop %>
  </table>
	</font></td>
    </tr>
<% End If %>
<!--END Editing Campaign so show list of zones-->
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
  <input type="hidden" name="CampaignType" value="Flat Rate">
</form>

</body>
</html>

<%

Set rsDefaultCampaign=Nothing
Set rsBanner=Nothing
Set rsCampaignBanners=Nothing
Set rsAdvertiser=Nothing
Set rsZoneList=Nothing
%>