<%

%>
	
	<!--#include file="dbconnect.asp"-->
	
<%

	'connect to database
	Set connBanManPro=Server.CreateObject("ADODB.Connection") 
	connBanManPro.Mode = 3      '3 = adModeReadWrite
	connBanManPro.Open Application("BannerManagerConnectString")

	If Request.QueryString("Login")="True" then
		Session("UserName")=Request.Form("UserName")
		Session("Password")=Request.Form("Password")
		If Trim(Request.Form("lstWebSites"))<>"" Then
			Session("BanManProSiteID")=CLng(Request.Form("lstWebSites"))
			strSQL="Select SiteName From BanManProWebSites Where SiteID=" & CLng(Request.Form("lstWebSites"))
			Set rs=connBanManPro.Execute(strSQL)
			Session("BanManProSiteName")=rs("SiteName")
		Else
			Session("BanManProSiteID")=0
		End If
		If Request.Form("StoreCookie")="ON" Then
			Response.Cookies ("BanManPro")("UserName") = Session("UserName")
			Response.Cookies ("BanManPro")("Password") = Session("Password")
		Else
			Response.Cookies ("BanManPro")("UserName") = ""
			Response.Cookies ("BanManPro")("Password") = ""
		End If
		Response.Cookies ("BanManPro").Expires=Date() + 180
		'Response.Cookies ("BanManPro").Domain=Request.ServerVariables("SERVER_NAME")
		'Response.Cookies ("BanManPro").Path=""
	End If


%>
