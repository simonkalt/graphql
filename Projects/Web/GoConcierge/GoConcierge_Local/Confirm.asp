<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Application("DotNetDSN") = "Password=sequoia;User ID=sa;Initial Catalog=Innsight;Data Source=Shaq2" '

function GenerateUserKey(v)
	tmpKey = ""
	i = 3
	do while len (tmpKey) < 9 and (i < len(v) AND i <= 15)
		key = cint(mid(v,i,1))
		if key > 0 and key < 10 Then
			tmpKey = tmpKey & key
		End if
		i = i + 1
	loop
	
	GenerateUserKey = tmpKey
On Error Goto 0
end function


'Response.Cookies ("UserKey").expires = Date()+1

Set cnMain = Server.CreateObject("ADODB.Connection")

cnMain.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")

set rsKey = cnMain.Execute ("select Rand()")
vtmpKey = rsKey(0)
set rsKey = Nothing
If Request.Cookies("UserKey") = "" Then
	tmpKey = GenerateUserKey(vtmpKey)
	Response.Cookies ("UserKey") = tmpKey
Else
	tmpKey = Request.Cookies("UserKey")		
End IF


Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (tmpKey)

cid = request.querystring("cid")

remote.Session("DefaultCalView") = 0

' New remote.Session Stuff
remote.Session("MultiCo") = "False"

Dim strSQL, x
Set rs = Server.CreateObject("ADODB.Recordset")

vk = "lkasj79834lk23472lk73427lkj23498lkasjdf-238z"

validkey = Request.QueryString("ValidKey")

if cid <> "" and validkey = vk then
	set rsUser = server.CreateObject("Adodb.recordset")
	set rsUser = cnMain.Execute("Select * from tblUser where username = 'Super' and userLName = 'Concierge'")
	un = rsUser.Fields("LoginName").Value
	pw = rsUser.Fields("password").Value
	rsUser.Close
	set rsUser = nothing
else
	un = request("logon")
	pw = request("password")
end if

strSQL = "sp_LoginConfirm '" & replace(un,"'","''") & "', '" & replace(pw,"'","''") & "'"
rs.Open strSQL,cnMain,adOpenKeyset,adLockReadOnly
'Check for no companies assigned to this user
If rs.EOF Then
	remote.Session("LoginOK") = "False"
	'Response.Redirect "Login.asp"%>
	<script language=vbscript>
		msgbox "Invalid Username or Password",,"Login Failed"
	</script>
	<%		
Else
Response.Write tmpKey & " " & rs("Admin")
	
	remote.Session("Admin") = rs("Admin")

	'Used for multiple companies
    remote.Session("UserID") = rs("UserID")
        
    remote.Session("VCT") = rs("ViewTasks")
    remote.Session("ACT") = rs("AddTask")
    remote.Session("ECT") = rs("EditTask")
    remote.Session("CCT")=rs("CloseTask")
    remote.Session("ELD") = rs("EditLocDir")
    remote.Session("BPW") = null2bool(rs("BypassPW"))
        
    remote.Session("Login") = request("logon") 'Actual login value from previous screen
    remote.Session("Password") = request("password")

    remote.Session("FloatingUser_Login") = request("logon")
    remote.Session("FloatingUser_Password") = request("password")
    remote.Session("FloatingUser_UserID") = rs("UserID").Value
    remote.Session("FloatingUser_UserName") = rs("UserName").Value
    remote.Session("FloatingUser_UserLName") = rs("UserLName").Value
    remote.Session("FloatingUser_Admin") = rs("Admin").Value
    
	remote.Session("FloatingUser_SuperUser") = rs("SuperUser").Value
	
	
	Response.Write "<br>" & remote.Session("FloatingUser_SuperUser")
	'Response.End 
	
	
	remote.Session("RolloverDefault") = rs("Rollover").Value
	
	remote.Session("FloatingUser_EMail") = rs("EmailAddress").Value
	remote.Session("FloatingUser_CCPrivate") = rs("CCPrivate").Value
	remote.Session("FloatingUser_CCPublic") = rs("CCPublic").Value
	remote.Session("FloatingUser_Title") = rs("Title").Value
	remote.Session("FloatingUser_OTUserID") = rs("OTUserID").Value
	remote.Session("FloatingUser_VPPN") = rs("ViewPrivate").Value
	remote.Session("FloatingUser_Phone") = rs("Phone").Value
	remote.Session("FloatingUser_DDID") = rs("DefaultDepartmentID").Value
	remote.session("DefaultDepartmentID") = rs("DefaultDepartmentID").Value
	
	
	remote.Session("_Login") = remote.Session("FloatingUser_Login")     
	remote.Session("_Password") = remote.Session("FloatingUser_Password")
	remote.Session("_UserID") = remote.Session("FloatingUser_UserID")   
	remote.Session("_UserName") = remote.Session("FloatingUser_UserName")  
	remote.Session("_UserLName") = remote.Session("FloatingUser_UserLName")
	remote.Session("_Admin") = remote.Session("FloatingUser_Admin")    
	remote.Session("_SuperUser") = remote.Session("FloatingUser_SuperUser")
	remote.Session("_EMail") = remote.Session("FloatingUser_EMail")    
	remote.Session("_CCPrivate") = remote.Session("FloatingUser_CCPrivate")
	remote.Session("_CCPublic") = remote.Session("FloatingUser_CCPublic")  
	remote.Session("_Title") = remote.Session("FloatingUser_Title")     
	
	remote.Session("UseGuestProfile") = rs.Fields("UseGuestProfile").Value
	remote.Session("GPSearchID") = rs.Fields("GPSearchID").Value
        
    remote.Session("FloatingUser_VCCN")=rs("ViewCC")

    remote.Session("SQLNOORDER") = ""
    
    remote.Session("CompanyID") = rs("CompanyID").Value
    Response.Cookies ("CompanyID") = rs("CompanyID").Value
    
    remote.Session("CompanyName") = rs("CompanyName").Value
    remote.Session("HotelLocationID") = rs("LocationID").Value

    if isnull(rs("LogoBGColor").value) then
		remote.Session("LogoBGColor") = ""
	else
		remote.Session("LogoBGColor") = trim(rs("LogoBGColor").Value)
	end if

    remote.Session("CompanyState") = rs("State").Value
    remote.Session("CompanyCity") = rs("City").Value
    If Len(rs("TimeZone"))<>0 Then
		remote.Session("TimeZone") = rs("TimeZone").Value/24
	Else
		remote.Session("TimeZone") = 0
	End IF
	
	'strMapUNC = rs("MapUNC").Value
	'strMapUNC = Right(strMapUNC,len(strMapUNC)-2)
	
	'remote.Session("MapVP")= Mid(strMapUNC,instr(1,strMapUNC,"\")+1,len(strMapUNC))
	'remote.Session("MapPath") = rs("MapUNC").Value & "\"
	
	set rsCCType = server.CreateObject("adodb.recordset")
	rsCCType.Open "select ChargeTypeID from tlkpChargeType where ChargeType = 'Cash'", cnMain
	remote.session("CashID") = rsCCType.Fields("ChargeTypeID").Value
	rsCCType.Close
	set rsCCType = nothing
	
	if cid <> "" and validkey = vk then
		remote.Session("ScreenHeight") = 768
		remote.Session("AvailHeight") = Request.QueryString("rsa")
		remote.Session("AvailWidth") = Request.QueryString("rsw")
	else
		remote.Session("ScreenHeight") = Request.Form("txtScreenHeight")
		remote.Session("AvailHeight") = Request.Form("txtAvailHeight")
		remote.Session("AvailWidth") = Request.Form("txtAvailWidth")
	end if
	
	remote.Session("MapTempPath") = Application("ENV_Path") & "Temp\"
	remote.Session("MapTempURL") = Application("HomePage") & "\Temp\"
    remote.Session("ScreenLogoLocation") = rs("ScreenLogoLocation").Value
    remote.Session("LogoLocation") = rs("LogoLocation").Value
        
    if rs("UseCompanyLetterHead").Value then
		remote.Session("LetterHead") = "Yes"
	else
		remote.Session("LetterHead") = "No"
	end if

    'New for Super User 4/7/01 7:11 PM.
	remote.Session("Admin") = rs("Admin").Value
	remote.Session("SuperUser")= rs("SuperUser").Value

	remote.Session("LoginOK") = "True"
	
	remote.Session("Loc_Recordset") = ""
	remote.Session("rsBrowseSelect") = ""
		
	remote.Session("PrinterList") = Request.Form ("PrinterList")
		
	remote.Session("CompanyZip") = rs("PostalCode").Value
	remote.Session("CompanyOTID") = rs("OTHotelID").Value 
	
	remote.Session("DefaultCategory") = rs("DefaultCategory").Value
	remote.Session("DefaultBCK") = rs("DefaultBCK").Value
	remote.Session("DefaultSortBy") = rs("DefaultSortBy").Value
	remote.Session("DefaultState") = rs("DefaultState").Value
	
	remote.Session("booModifyReport") = false
	
	remote.Session("DefaultCalView") = rs("DefaultCalView").Value
	

    function isFilled( v )
		retVal = true
		if isnull(v) then
			retVal = false
		else
			if len(trim(v)) = 0 then
				retVal = false
			end if
		end if
		isFilled = retVal
    end function

	if not isFilled(rs.Fields("WeatherURL").Value) then
		remote.session("WeatherURL") = "www.weather.com/weather/local/" & left(rs("PostalCode").Value,5)
	else
		remote.session("WeatherURL") = rs.Fields("WeatherURL").Value
	end if
	if not isFilled(rs.Fields("TicketsURL").Value) then
		remote.session("TicketsURL") = "www.ticketmaster.com"
	else
		remote.session("TicketsURL") = rs.Fields("TicketsURL").Value
	end if
	if not isFilled(rs.Fields("MoviesURL").Value) then
		'remote.session("MoviesURL") = "www.moviefone.com/showtimes/closesttheaters.adp?csz=" & left(rs("PostalCode").Value,5) & "&_action=setLocation"
		remote.session("MoviesURL") = "movies.channel.aol.com/search/locationresults.adp?csz=" & left(rs("PostalCode").Value,5) & "&_action=setLocation"
	else
		remote.session("MoviesURL") = rs.Fields("MoviesURL").Value
	end if
	
	if not isFilled(rs.Fields("ZagatURL").Value) then
		remote.session("ZagatURL") = "www.zagat.com"
	else
		remote.session("ZagatURL") = rs.Fields("ZagatURL").Value
	end if
	
	if not isFilled(rs.Fields("FlightsURL").Value) then
		remote.session("FlightsURL") = "www.flytecomm.com/cgi-bin/trackflight"
	else
		remote.session("FlightsURL") = rs.Fields("FlightsURL").Value
	end if

	'Response.Write remote.Session("PrinterList")
	'arPrint = Split(Request.Form ("PrinterList"),"|")
    'for i = LBound(arPrint) to UBound(arPrint) - 1
		'cnMain.Execute "sp_HotelPrinter_Add " & remote.Session("CompanyID") & "," & remote.Session("UserID") & ", '" & arPrint(i) & "'"
    'next
    function null2bool(x)
		if isnull(x) or isempty(x) then
			null2bool = false
		else
			null2bool = x
		end if
    end function

if cid <> "" and validkey = vk then
	Response.Redirect "SelectCompanyConfirm.asp?lid=" & Request.QueryString("lid") & "&id=" & cid
else
%>
	<script language=vbscript>
		parent.document.FormPassword.logon.value = ""
		parent.document.FormPassword.password.value = ""
		dim v, h, w
		h = screen.height-80
		w = screen.width-10
		v = "scrollbars=no,status=yes,location=no,menubar=no,toolbar=no,top=0,left=0,height=" & h & ",width=" & w
		<%If rs.Fields("RowsReturned").Value > 1 Then%>
			parent.OpenApp 1, v, "", "<%=tmpKey%>"
		<%Else%>
			parent.OpenApp 2, v, "<%=FormatDateTime(Date(),2)%>", "<%=tmpKey%>"
		<%End If%>
	</script>
<%end if
end If
rs.Close
set rs = nothing
cnMain.Close
set cnMain = nothing
%>
<!--#INCLUDE FILE="header.inc" -->
