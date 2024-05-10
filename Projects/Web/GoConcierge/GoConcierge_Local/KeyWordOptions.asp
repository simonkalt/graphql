<%@ Language=VBScript %>
<%
Response.Expires = -1

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))

Set cnSQL = Server.CreateObject("ADODB.Connection")
cnSQL.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")

set rs = Server.CreateObject ("Adodb.Recordset")

SQL = Request.QueryString ("SQL")

set rs = cnSQL.Execute (SQL)



if rs.State = 1 Then
	Do while not rs.EOF 

		tmpVal = ""
		for i = 0 to rs.Fields.count - 1
			If Len(rs(i).Value) > 0 Then
				tmpVal = tmpVal & rs(i).Value & "|"
			Else
				tmpVal = tmpVal &  "(null)|"
			End If
			
		next 
		
		tmpVal = Left(tmpVal,len(tmpVal) - 1)
		Response.Write tmpVal & "||"

	rs.MoveNext 

	Loop
End If


%>
