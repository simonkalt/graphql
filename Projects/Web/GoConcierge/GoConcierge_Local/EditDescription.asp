<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
dim cn, rs, Mode
dim strTable, strIDFieldName, intID, strDescriptionFieldName, strDescription
strTable = Request.QueryString("Table")

'Response.Write strTable 

strIDFieldName = Request.QueryString("IDFieldName")
intID = Request.QueryString("ID")
strDescriptionFieldName = Request.QueryString("DescriptionFieldName")

set cn = server.CreateObject("ADODB.connection")
cn.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")

Mode = Request.QueryString("Mode")

'if Mode = "Edit" then
	set rs = server.CreateObject("ADODB.recordset")
	strSQL = "SELECT * FROM " & strTable & " WHERE " & strIDFieldName & " = " & intID
	'Response.Write strSQL
	rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
	strDescription = rs(strDescriptionFieldName)

	If strTable="tlkpNotesFields" Then 
	
		strDisplay = rs("DisplayText")
		
		chkInclude = " checked "
		If rs("Include")=0 Then
			chkInclude = " "
		End If

	Else 
	
		If strTable="tlkpAction" Then
			
			if rs("rollover") Then
				chkRoll = " checked "			
			End If
			
		End If

	End IF
	
'else
'	strDescription = ""
'end if
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<style>
<!--
	.StandardFont	{ font-face: tahoma; font-size: 14 }
-->
</style>
<TITLE><%=Mode%>&nbsp;Description</TITLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub cmdCancel_onclick
	window.returnValue = "<%=strDescription%>"
	window.close
End Sub

Sub cmdOK_onclick

	<% If strTable="tlkpNotesFields" Then %>
	
			str = "EditDescriptionCommit.asp?Mode=<%=Mode%>&Table=<%=strTable%>&IDFieldName=<%=strIDFieldName%>&ID=<%=intID%>&DescriptionFieldName=<%=strDescriptionFieldName%>&Description=" & escape(document.all("txt").value)
			str = str & "&txtDisplay=" & escape(document.all("txtDisplay").value) & "&include=" & document.all("chkInclude").checked
			document.all("frmSubmit").src =  str
	<% Else %>
			if document.all("chkRoll").checked Then
				booRoll = 1
			Else
				booRoll = 0
			End If
			
			str = "EditDescriptionCommit.asp?Mode=<%=Mode%>&Table=<%=strTable%>&IDFieldName=<%=strIDFieldName%>&ID=<%=intID%>&DescriptionFieldName=<%=strDescriptionFieldName%>&Description=" & escape(document.all("txt").value) & "&roll=" & booRoll
			document.all("frmSubmit").src = str
			
	<%End IF%>

	<%'if Mode = "Edit" then%>
		window.returnValue = document.all("txt").value
	<%'else%>
		'window.returnValue = document.all("txt").value & "|" & window.frmSubmit.document.all("txtID").value
		'msgbox window.returnValue
	<%'end if%>
End Sub

Sub Body_onkeydown
	if window.event.keyCode = 13 then
		cmdOK_onclick
	end if
End Sub

-->
</SCRIPT>
</HEAD>

<BODY id=Body bgcolor=#FAD667 class="StandardFont" leftmargin=10 topmargin=5>
<!--#include file=Global.asp-->
<iframe src="nullSrc()" id=frmSubmit name=frmSubmit border=0 style="height: 1px; width: 1px"></iframe>
<table cellpadding=5 align="center" style="font-face: tahoma; font-size: 16">
	<tr align="center">
		<td>Text:&nbsp;<INPUT style="font-face: tahoma; font-size: 14; width: 330px" class="StandardFont" type="text" id=txt name=txt value="<%=strDescription%>"></td>
	</tr>
	
	<% If strTable="tlkpNotesFields" Then %>
		<tr align="center">
			<td>Disp:&nbsp;<INPUT style="font-face: tahoma; font-size: 14; width: 330px" class="StandardFont" type="text" id=txtDisplay name=txtDisplay value="<%=strDisplay%>"></td>
		</tr>
		<tr align="left">
			<td>Include:&nbsp;<INPUT style="font-face: tahoma; font-size: 14;" type="CheckBox" class="StandardFont"  id=chkInclude <%=chkInclude%>></td>
		</tr>
	<%Else%>
		<tr align="left">
			<td>Rollover:&nbsp;<INPUT style="font-face: tahoma; font-size: 14;" type="CheckBox" class="StandardFont"  id=chkRoll <%=chkRoll%>></td>
		</tr>
	<%End IF%>
	
	
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr align="center">
		<td><INPUT style="width: 120" type="button" value="OK" id=cmdOK name=cmdOK>&nbsp;<INPUT style="width: 120" type="button" value="Cancel" id=cmdCancel name=cmdCancel></td>
	</tr>
</table>

</BODY>
</HTML>

<%
'if Mode = "Edit" then
	rs.Close
	set rs = nothing
'end if

cn.Close
set cn = nothing
%>