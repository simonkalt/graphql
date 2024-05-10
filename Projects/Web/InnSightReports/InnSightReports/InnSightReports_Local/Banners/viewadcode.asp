<% 'create ad code

	strSQLMode="SELECT * FROM Zones WHERE Zones.ZoneID=" & strZoneID 
	Set rsZoneMode=connBanManPro.Execute(strSQLMode)

	If Not rsZoneMode.EOF Then

			'obtain random number
			Randomize
			lngRandom=Int((100000 - 1 + 1) * Rnd + 1)

			'SSI mode
			strTarget=getFilePath() & "zones/banmanzone" & strZoneID & ".asp"
			strAdCodeSSI="<!--#include " & "virtual=" & Chr(34) & strTarget & Chr(34) & "-->"

			'Advanced Code
			'get heights/widths for zone
			strSQL="SELECT ZoneCampaigns.ZoneID, Banners.* "
			strSQL=strSQL & "FROM ZoneCampaigns INNER JOIN (Banners RIGHT JOIN CampaignBanners "
 			strSQL=strSQL & "ON Banners.BannerID = CampaignBanners.BannerID) ON ZoneCampaigns.CampaignID "
			strSQL=strSQL & "= CampaignBanners.CampaignID WHERE ZoneCampaigns.ZoneID=" & strZoneID
			Set rsBanners=connBanManPro.Execute(strSQL)
			strBanWidth="468"
			strBanHeight="60"
			strBanBorder="0"
			If NOT rsBanners.EOF Then
				strBanWidth=rsBanners("AdWidth")
				strBanHeight=rsBanners("AdHeight")
				strBanBorder=rsBanners("AdBorder")
			End If
			If rsZoneMode("ZoneWidth")=0 Or Trim(rsZoneMode("ZoneWidth"))="" or IsNull(rsZoneMode("ZoneWidth")) Then
				strBanWidth=468
			Else
				strBanWidth=rsZoneMode("ZoneWidth")
			End If
			If rsZoneMode("ZoneHeight")=0 Or Trim(rsZoneMode("ZoneHeight"))="" Or IsNull(rsZoneMode("ZoneHeight")) Then
				strBanHeight=60
			Else
				strBanHeight=rsZoneMode("ZoneHeight")
			End If

			'Site ID
			If Application("BanManProMultiSite")=True Then
				strExtraTag= "&SiteID=" & CLng(Session("BanManProSiteID"))
				lngSiteID= CLng(Session("BanManProSiteID"))
				Set rsTemp=connBanManPro.Execute("Select SiteName From BanManProWebSites Where SiteID=" & CLng(Session("BanManProSiteID")))
				If Not rsTemp.EOF Then
					strSiteName=" Site: " & rsTemp("SiteName") & " "
				End If
			Else
				strExtraTag=""
				lngSiteID=0
				strSiteName=""
			End If

			'IMage URL For HTML Mode
			strImageURL=Application("DomainURL") & "?ZoneID=" & strZoneID & "&Task=Get&Mode=HTML" & strExtraTag
			'Click URL
			strClickURL=Application("DomainURL") & "?ZoneID=" & strZoneID & "&Task=Click&Mode=HTML" & strExtraTag
			'strAdCode="<a href=" & Chr(34) & strClickURL & Chr(34) & "><img src=" & Chr(34) & strImageURL & Chr(34) & "></a>"
			strAdCode="<SCRIPT LANGUAGE=" & Chr(34) & "Javascript" & Chr(34) & ">" & Chr(10)
			strAdCode=strAdCode & "<!-- " & Chr(10)
			strAdCode=strAdCode & "document.write('<A HREF=" & Chr(34) & strClickURL & Chr(34) & "><IMG SRC=" & Chr(34) & strImageURL & "&fightcache=' + (new Date()).getTime() + '" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></A>');" & Chr(10)
			strAdCode=strAdCode & "//-->" & Chr(10)
			strAdCode=strAdCode & "</SCRIPT>"
			strAdCode=strAdCode & "<noscript>" & Chr(10)
			strAdCode=strAdCode & "    <a href=" & Chr(34) & strClickURL & Chr(34) & " target=" & Chr(34) & "_new" & Chr(34) & ">" & Chr(10) 
			strAdCode=strAdCode & "    <img src=" & Chr(34) & strImageURL & Chr(34) & " width=" & Chr(34) & strBanWidth & Chr(34) & " height=" & Chr(34) & strBanHeight & Chr(34) & " border=" & Chr(34) & strBanBorder & Chr(34) & "></a>" & Chr(10)
			strAdCode=strAdCode & "</noscript>" & Chr(10) 
			
			'Create Advanced Java Code
			strAdCodeJavaNew="<!-- Begin Ban Man Pro Banner Code - " & strSiteName & " Zone: " & rsZoneMode("ZoneDescription") & " -->" & Chr(10)
			strAdCodeJavaNew=strAdCodeJavaNew & "<SCRIPT LANGUAGE=" & Chr(34) & "JAVASCRIPT" & Chr(34) & ">" & Chr(10) 
			strAdCodeJavaNew=strAdCodeJavaNew & "<!--" & Chr(10)
			strAdCodeJavaNew=strAdCodeJavaNew & "var browName = navigator.appName;" & Chr(10) 
			strAdCodeJavaNew=strAdCodeJavaNew & "var browVersion = parseInt(navigator.appVersion);" & Chr(10) 
			strAdCodeJavaNew=strAdCodeJavaNew & "var ua=navigator.userAgent.toLowerCase();" & Chr(10) 
			strAdCodeJavaNew=strAdCodeJavaNew & "var adcode='';" & Chr(10) 
			strAdCodeJavaNew=strAdCodeJavaNew & "if (browName=='Netscape'){" & Chr(10)
  			strAdCodeJavaNew=strAdCodeJavaNew & "     if ((browVersion>=4)&&(ua.indexOf(" & Chr(34) & "mac" & Chr(34) & ")==-1))" & Chr(10)
 			strAdCodeJavaNew=strAdCodeJavaNew & "          { document.write('<S'+'CRIPT src=" & Chr(34) & Application("DomainURL") & "?ZoneID=" & strZoneID & "&Task=Get&Browser=NETSCAPE4" & strExtraTag & Chr(34) & ">');" & Chr(10)
       			strAdCodeJavaNew=strAdCodeJavaNew & "          document.write('</'+'scr'+'ipt>');" & Chr(10)
			strAdCodeJavaNew=strAdCodeJavaNew & "          document.write(adcode); }" & Chr(10)
     			strAdCodeJavaNew=strAdCodeJavaNew & "     else if (browVersion>=3) " & Chr(10) 
			strAdCodeJavaNew=strAdCodeJavaNew & "          { document.write('<A HREF=" & Chr(34) & strClickURL & Chr(34) & " target=" & Chr(34) & "_new" & Chr(34) & "><IMG SRC=" & Chr(34) & strImageURL & "&fightcache=' + (new Date()).getTime() + '" & Chr(34) & " width=" & Chr(34) & strBanWidth & Chr(34) & " height=" & Chr(34) & strBanHeight & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></A>'); } }" & Chr(10) 
			strAdCodeJavaNew=strAdCodeJavaNew & "if (browName=='Microsoft Internet Explorer')" & Chr(10)
   			strAdCodeJavaNew=strAdCodeJavaNew & "     { document.write('<ifr'+'ame src=" & chr(34) & Application("DomainURL") & "?ZoneID=" & strZoneID & "&Task=Get" & strExtraTag & chr(34) & " width=" & strBanWidth + strIEBanBorder & " height=" & strBanHeight + strIEBanBorder & " Marginwidth=0 Marginheight=0 Hspace=0 Vspace=0 Frameborder=0 Scrolling=No></ifr'+'ame>'); }" & Chr(10)
			strAdCodeJavaNew=strAdCodeJavaNew & "// --> " & Chr(10)
			strAdCodeJavaNew=strAdCodeJavaNew & "</script>" & Chr(10) 
			strAdCodeJavaNew=strAdCodeJavaNew & "<noscript>" & Chr(10)
			strAdCodeJavaNew=strAdCodeJavaNew & "    <a href=" & Chr(34) & strClickURL & "&PageID=" & lngRandom & Chr(34) & " target=" & Chr(34) & "_new" & Chr(34) & ">" & Chr(10) 
			strAdCodeJavaNew=strAdCodeJavaNew & "    <img src=" & Chr(34) & strImageURL & "&PageID=" & lngRandom &  Chr(34) & " width=" & Chr(34) & strBanWidth & Chr(34) & " height=" & Chr(34) & strBanHeight & Chr(34) & " border=" & Chr(34) & strBanBorder & Chr(34) & "></a>" & Chr(10)
			strAdCodeJavaNew=strAdCodeJavaNew & "</noscript>" & Chr(10) 
			strAdCodeJavaNew=strAdCodeJavaNew & "<!-- End Ban Man Pro Banner Code - " & strSiteName & " Zone: " & rsZoneMode("ZoneDescription") & " -->"	


			'non-cache defeating code
			strAdCodeNoCache="<!-- Begin Ban Man Pro Banner Code - " & strSiteName & " Zone: " & rsZoneMode("ZoneDescription") & " -->" & Chr(10)
			strAdCodeNoCache=strAdCodeNoCache & "<IFRAME SRC=" & chr(34) & Application("DomainURL") & "?ZoneID=" & strZoneID & "&Task=Get" & "&PageID=" & lngRandom & strExtraTag &  chr(34) & " width=" & strBanWidth + strIEBanBorder & " height=" & strBanHeight + strIEBanBorder & " Marginwidth=0 Marginheight=0 Hspace=0 Vspace=0 Frameborder=0 Scrolling=No>" & Chr(10)
			strAdCodeNoCache=strAdCodeNoCache & "<SCRIPT LANGUAGE='JAVASCRIPT1.1' SRC=" & Chr(34) & Application("DomainURL") & "?ZoneID=" & strZoneID & "&Task=Get&Browser=NETSCAPE4&NoCache=True" & "&PageID=" & lngRandom & strExtraTag &  Chr(34) & ">"
			strAdCodeNoCache=strAdCodeNoCache & "</SCRIPT>" & Chr(10)
			strAdCodeNoCache=strAdCodeNoCache & "<NOSCRIPT>"
			strAdCodeNoCache=strAdCodeNoCache & "<a href=" & Chr(34) & strClickURL & "&PageID=" & lngRandom &  Chr(34) & " target=" & Chr(34) & "_new" & Chr(34) & ">" & Chr(10) 
			strAdCodeNoCache=strAdCodeNoCache & "<img src=" & Chr(34) & strImageURL & "&PageID=" & lngRandom &  Chr(34) & " width=" & Chr(34) & strBanWidth & Chr(34) & " height=" & Chr(34) & strBanHeight & Chr(34) & " border=" & Chr(34) & strBanBorder & Chr(34) & "></a>" & Chr(10)
			strAdCodeNoCache=strAdCodeNoCache & "</NOSCRIPT>" & chr(10)
			strAdCodeNoCache=strAdCodeNoCache & "</IFRAME>" & chr(10)
			strAdCodeNoCache=strAdCodeNoCache & "<!-- End Ban Man Pro Banner Code - " & strSiteName & " Zone: " & rsZoneMode("ZoneDescription") & " -->"	

			'simple,simple HTML
			strSimplest="<!-- Begin Ban Man Pro Banner Code - " & strSiteName & " Zone: " & rsZoneMode("ZoneDescription") & " -->" & Chr(10)
			strSimplest=strSimplest & "<a href=" & Chr(34) & strClickURL & "&PageID=" & lngRandom &  Chr(34) & " target=" & Chr(34) & "_new" & Chr(34) & ">" & Chr(10) 
			strSimplest=strSimplest & "<img src=" & Chr(34) & strImageURL & "&PageID=" & lngRandom &  Chr(34) & " width=" & Chr(34) & strBanWidth & Chr(34) & " height=" & Chr(34) & strBanHeight & Chr(34) & " border=" & Chr(34) & strBanBorder & Chr(34) & "></a>" & Chr(10)
			strSimplest=strSimplest & "<!-- End Ban Man Pro Banner Code - " & strSiteName & " Zone: " & rsZoneMode("ZoneDescription") & " -->"	
	
			'Call Ban Man Pro By Function
			strCallByFunc="<" & Chr(37) & "'Place the Ban Man Pro function include at the beginning of the ASP page." & Chr(37) & ">" & Chr(10)
			strTarget=getFilePath() & "banmanfunc.asp"
			strCallByFunc=strCallByFunc & "<!--#include " & "virtual=" & Chr(34) & strTarget & Chr(34) & "-->" & Chr(10) & Chr(10)
			If Application("BanManProMultiSite")=True Then
				strCallByFunc=strCallByFunc & "<" & Chr(37) & "'Call Ban Man Pro Ad Using Subroutine GetBanManProAd(ZoneID,ZoneName,Keywords,Mode,SiteID)" & Chr(10)
			Else
				strCallByFunc=strCallByFunc & "<" & Chr(37) & "'Call Ban Man Pro Ad Using Subroutine GetBanManProAd(ZoneID,ZoneName,Keywords,Mode,0)" & Chr(10)
			End If
			strCallByFunc=strCallByFunc & "GetBanManProAd " & strZoneID & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & "SSI" & Chr(34) & "," & lngSiteID & Chr(37) & ">"

	End If

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
        <td><a href="help/zones.htm" target="_new"><img border="0" src="images/ListingofAllZones.gif" WIDTH="590" HEIGHT="30"></a></td>
      </tr>
    </table>
  </center>
</div>
<div align="center"><center>

<table border="0" cellpadding="5" cellspacing="0" width="590" background="images/tableback.gif">
  <tr>
    <td><font face="Arial" size="3">Copy and paste the code below to the page
    where you wish to display the banners for this zone.&nbsp; Need help
      deciding which code is right for you?&nbsp; <a href="http://www.banmanpro.com/support/codesummary.asp" target="_new">Click
      here for more information</a>.<br>
      </font>
      <hr>
      <p align="center"><%=strAdCodeJavaNew%></td>
  </tr>
  <tr>
    <td><strong><font face="Arial" size="4">Advanced Java Code (Highly Recommended)
      --</font><font face="Arial" size="2"> </font></strong><font face="Arial" size="2">Serves
      rich media ads and defeats cache. It is recommended that you replace
      PageID with a unique number on each page.&nbsp; The number must be
      identical within the same Ad Code snippet and appears in two locations
      near the end of the code.</font></td>
  </tr>
  <tr>
    <td><form method="POST" action>
      <p align="center"><textarea rows="10" name="ZoneCode" cols="60"><%=strAdCodeJavaNew%></textarea></p>
    </form>
    </td>
  </tr>
  <tr>
    <td><strong><font face="Arial" size="4">Non-Cache Defeating Code
      </font>-- </strong><font face="Arial" size="2">Serves rich media ads but
      does not defeat cache.&nbsp; It is recommended that you replace PageID
      with a unique number on each page.&nbsp; The number must be identical
      within the same Ad Code snippet and appears in FOUR locations throughout the code.</font></td>
  </tr>
  <tr>
    <td><form method="POST" action>
      <p align="center"><textarea rows="8" name="ZoneCode" cols="60"><%=strAdCodeNoCache%></textarea></p>
    </form>
    </td>
  </tr>
  <tr>
    <td><strong><font face="Arial" size="4">Simple HTML </font>-- </strong><font face="Arial" size="2">Serves
      only image ads.&nbsp; You must replace the PageID with a unique number on
      EVERY single page in your web site. The number must be identical within
      the same Ad Code snippet and appears in two locations.</font>
    </td>
  </tr>
  <tr>
    <td><form method="POST" action>
      <p align="center"><textarea rows="4" name="ZoneCode" cols="60"><%=strSimplest%></textarea></p>
    </form>
    </td>
  </tr>
  <tr>
    <td><strong><font face="Arial" size="4">SSI Code </font>-- </strong><font face="Arial" size="2">Works
      only on pages with an ASP extension.&nbsp; Also, Ban Man Pro must be
      installed on the same server as the web server or a virtual directory must
      be setup to point to the ad server.</font>
    </td>
  </tr>
  <tr>
    <td><form method="POST" action>
      <p align="center"><textarea rows="4" name="ZoneCode" cols="60"><%=strAdCodeSSI%></textarea></p>
    </form>
    </td>
  </tr>
  <tr>
    <td><strong><font face="Arial" size="4">Call By Function (ASP Only) </font>-- </strong><font face="Arial" size="2">Works
      only on pages with an ASP extension.&nbsp; Also, Ban Man Pro must be
      installed on the same server as the web server or a virtual directory must
      be setup to point to the ad server.</font>
    </td>
  </tr>
  <tr>
    <td><form method="POST" action>
      <p align="center"><textarea rows="6" name="ZoneCodeFunction" cols="60"><%=strCallByFunc%></textarea></p>
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