<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))

dim rs, cn, strSQL, strField

gid = Request.QueryString("id")

select case Request.QueryString("mode")
	case 0:
		' should never happen
	case 1:
		strField = "GuestID"
	case 2:
		strField = "PMSGuestID"
		gid = "'" & Replace(Request.QueryString("id"),"'","''") & "'"
	case 3:
		strField = "HotelGuestID"
		gid = "'" & Replace(Request.QueryString("id"),"'","''") & "'"
end select

set cn = server.CreateObject("adodb.connection")
cn.Open Application("sqlInnSight_ConnectionString")

set rs = server.CreateObject("adodb.recordset")

strSQL = "select * from vw_GuestHotel where " & strField & " = " & gid & " and HotelID = " & Remote.Session("CompanyID")

set rs = cn.Execute(strSQL)
if rs.EOF then
	Response.Write "EOF"
else
	Response.Write rs.Fields("Salutation").Value & "|" & rs.Fields("LastName").Value & "|" & rs.Fields("FirstName").Value & "|" & trim(rs.Fields("PrimaryPhone").Value) & "|" & trim(rs.Fields("PhoneExt").Value) & "|" & rs.Fields("EMail1").Value & "|" & rs.Fields("ChargeTypeID").Value & "|" & rs.Fields("ChargeNumber").Value & "|" & rs.Fields("Expiration").Value & "|" & rs.Fields("GuestID").Value
end if

rs.Close
set rs = nothing
cn.Close
set cn = nothing
%>
