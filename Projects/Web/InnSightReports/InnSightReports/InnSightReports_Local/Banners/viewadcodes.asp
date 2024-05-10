<% 'create ad code

			'Site ID
			If Application("BanManProMultiSite")=True Then
				strExtraTag= "&SiteID=" & CLng(Session("BanManProSiteID"))
			Else
				strExtraTag=""
			End If

			strMode="TEXT"
			'get information for zone
			strSQL="SELECT CampaignBanners.CampaignID,CampaignBanners.BannerID, CampaignBanners.UserID, "
    		strSQL=strSQL & "Campaigns.AdvertiserID, Banners.AdTextLinkText FROM CampaignBanners INNER JOIN "
    		strSQL=strSQL & "Campaigns ON CampaignBanners.CampaignID = Campaigns.CampaignID INNER JOIN "
    		strSQL=strSQL & "Banners ON CampaignBanners.BannerID = Banners.BannerID "
			strSQL=strSQL & "WHERE (CampaignBanners.UserID = " & CLng(Session("BanManProSiteID")) & ") AND "
    		strSQL=strSQL & " (CampaignBanners.CampaignID = " & Clng(Request("CampaignID")) & ")"
			Set rsBanners=connBanManPro.Execute(strSQL)
			strTextLinkText=rsBanners("AdTextLinkText")
			strBannerID=rsBanners("BannerID")
			strAdvertiserID=rsBanners("AdvertiserID")
			strCampaignID=Clng(Request("CampaignID"))

			strImageURL=Application("DomainURL") & "?ZoneID=0&BannerID=" & strBannerID & "&AdvertiserID=" & strAdvertiserID & "&CampaignID=" & strCampaignID & "&Task=Get&Mode=TEXT" & strExtraTag
			strClickURL=Application("DomainURL") & "?ZoneID=0&BannerID=" & strBannerID & "&AdvertiserID=" & strAdvertiserID & "&CampaignID=" & strCampaignID & "&Task=Click&Mode=TEXT" & strExtraTag

			strAdCode="<!-- Begin Ban Man Pro Text Link Code -->" & vbCRLF
			strAdCode=strAdCode & "<a href=" & Chr(34) & strClickURL & Chr(34) & ">" & strTextLinkText & "</a>"
			strAdCode=strAdCode & "<img src=" & Chr(34) & strImageURL & Chr(34) & " width=" & Chr(34) & "1" & Chr(34) & " height=" & Chr(34) & "1" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & ">" & vbCRLF
			strAdCode=strAdCode & "<!-- End Ban Man Pro Text Link Code -->"
			


%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title></title>
</head>

<body>
<div align="center">
  <center>
    <table border="0" cellpadding="0" cellspacing="0" width="590">
      <tr>
        <td><img border="0" src="images/viewad1.gif" WIDTH="590" HEIGHT="30"></td>
      </tr>
    </table>
  </center>
</div>
<div align="center"><center>

<table border="0" cellpadding="5" cellspacing="0" width="590" background="images/tableback.gif">
  <tr>
    <td><strong><font face="Arial" size="4">Static Link Ad Code </font>-- </strong><font face="Arial" size="2">Copy
      and paste the following code to your web pages where you want to show this
      static text link.</font>
    </td>
  </tr>
  <tr>
    <td><form method="POST" action>
      <p align="center"><textarea rows="8" name="ZoneCode" cols="60"><%=strAdCode%></textarea></p>
    </form>
    </td>
  </tr>
  <tr>
    <td>&nbsp;
    </td>
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
</body>
</html>
<%
Function getFilePath()
	Dim lsPath, arPath

	' Obtain the virtual file path. The SCRIPT_NAME
	' item in the ServerVariables collection in the
	' Request object has the complete virtual file path
	lsPath = Request.ServerVariables("SCRIPT_NAME")
                           
	' Split the path along the /s. This creates an
	' This creates an one-dimensional array 
	arPath = Split(lsPath, "/")

	' Set the last item in the array to blank string
	' (The last item actually is the file name)
	arPath(UBound(arPath,1)) = ""
	
	' Join the items in the array. This will
	' give you the virtual path of the file
	GetFilePath = Join(arPath, "/")
End Function

%>