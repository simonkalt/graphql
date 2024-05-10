<%@ Language=VBScript %>

<%
Response.Expires = -1

dim cn, rs, Mode
dim strTable, strIDFieldName, intID, strSQL
dim strDescriptionFieldName, strDescription
strTable = Request.QueryString("Table")
strIDFieldName = Request.QueryString("IDFieldName")
intID = Request.QueryString("ID")
strDescriptionFieldName = Request.QueryString("DescriptionFieldName")
strDescription = Request.QueryString("Description")

set cn = server.CreateObject("ADODB.connection")
'set rs = server.CreateObject("ADODB.recordset")

cn.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")

Mode = Request.QueryString("Mode")
roll = Request.QueryString("roll")
'if Mode = "Edit" then
If strTable="tlkpNotesFields" Then
	
	include="0"
	If Ucase(Trim(Request.QueryString("Include")))="TRUE" Then
		include = "1"
	End If
	
	strSQL = "UPDATE " & strTable & " SET " & strDescriptionFieldName & " = '" & replace(strDescription,"'","''") & "', DisplayText='"& Request.QueryString("txtDisplay") & "', Include=" & include & "  WHERE " & strIDFieldName & " = " & intID
Else
	if strTable = "tlkpAction" then
		strSQL = "UPDATE " & strTable & " SET rollover=" & roll & ", " & strDescriptionFieldName & " = '" & replace(strDescription,"'","''") & "' WHERE " & strIDFieldName & " = " & intID
	else
		strSQL = "UPDATE " & strTable & " set " & strDescriptionFieldName & " = '" & replace(strDescription,"'","''") & "' WHERE " & strIDFieldName & " = " & intID
	end if
End If
'else
	'strSQL = "INSERT INTO " & strTable & " (" & strDescriptionFieldName & ") VALUES ('" & strDescription & "')"
'end if

'Response.Write Request.QueryString & "<br>"
'Response.Write "<script>alert(" & strSQL & ")</script>"

%>
<HTML>
<HEAD>
</HEAD>

<BODY bgcolor=#FAD667 class="StandardFont" leftmargin=10 topmargin=5>
<%
cn.Execute strSQL
'set rs = cn.Execute("SELECT @@IDENTITY AS ID")
%>
<!--input type=hidden id=txtID name=txtID value="<%'=rs("ID")%>"-->
</BODY>
</HTML>
<%
'rs.close
'set rs = nothing
cn.Close
set cn = nothing

Response.Write "<script>parent.window.close()</script>"
%>