<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))

dim rs, cn, strSQL, strField

last = Request.QueryString("last")
first = Request.QueryString("first")

set cn = server.CreateObject("adodb.connection")
cn.Open Application("sqlInnSight_ConnectionString")

set rs = server.CreateObject("adodb.recordset")

strSQL = "select *, case when lastname = '" & last & "' and firstname = '" & first & "' then 1 else 0 end as Score from vw_GuestHotel where LastName like '" & left(last,4) & "%' and FirstName like '" & left(first,1) & "%' and HotelID = " & Remote.Session("CompanyID")

set rs = cn.Execute(strSQL)
if rs.EOF then
	Response.Write "EOF"
else
	Response.Write rs.Fields("Salutation").Value & "|" & rs.Fields("LastName").Value & "|" & rs.Fields("FirstName").Value & "|" & trim(rs.Fields("PrimaryPhone").Value) & "|" & rs.Fields("EMail1").Value & "|" & rs.Fields("ChargeTypeID").Value & "|" & rs.Fields("ChargeNumber").Value & "|" & rs.Fields("Expiration").Value & "|" & rs.Fields("GuestID").Value & "|" & rs.Fields("Score").value
end if

rs.Close
set rs = nothing
cn.Close
set cn = nothing
%>
