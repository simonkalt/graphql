<% 

Application.Lock
Application("AdministratorName")=""
Application("AdministratorPassword")=""
Application("CacheBustingMode")=""
Application("BannerManagerConnectString")=""
Application("BanManProStatsInProgress")=False
Application.Unlock
Session("UserName")=""
Session("Password")=""
Session.Abandon

'Find homepage of ban man pro
If IsNull(Application("DomainURL")) Then
	strTargetURL=""
Else
	lngPos=Instr(UCASE(Application("DomainURL")),"BANMAN.ASP")
	strTargetURL=Left(Application("DomainURL"),lngPos-1)
End If

Response.Redirect strTargetURL & "default.asp"

%>