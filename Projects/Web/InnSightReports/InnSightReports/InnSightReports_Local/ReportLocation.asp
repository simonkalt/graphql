<%@ Language=VBScript %>
<!-- #include file="Data/adovbs.asp" -->
<!--#INCLUDE file="checkuser.asp"-->

<%
  Set cnSQL = Server.CreateObject("ADODB.Connection")
  cnSQL.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")
%>

<%
	Dim strList
	'Display list of locations
	Select Case Request.QueryString("Mode")
		Case "v" 'view
		  strList = Request.Form("txtViewLocationList")

		Case "p" 'print
		  strList = Request.Form("txtPrintLocationList")
		  
		Case "e" 'email
			strList = Request.Form("txtViewLocationList")
	End Select

	' Tweak to allow viewing and printing from Add/Edit Location forms
	if Request.QueryString("Loc") <> "" then
		  strList = Request.QueryString("Loc")
	End if
	
	
	Dim aList()
	ReDim aList(0)
	Dim intPos
	intPos = InStr(strList, ",")
	If intPos > 0 Then
		Do While intPos > 0
			ReDim Preserve aList(UBound(aList) + 1)
			aList(UBound(aList)) = Left(strList, intPos - 1)
			strList = Right(strList, Len(strList) - intPos)
			intPos = InStr(strList, ",")
			If intPos = 0 Then
				ReDim Preserve aList(UBound(aList) + 1)
				aList(UBound(aList)) = strList
			End If
		Loop
	Else
		ReDim Preserve aList(UBound(aList) + 1)
		aList(UBound(aList)) = strList
	End If

%>
<%
	Response.Expires = 0
	Server.ScriptTimeout = 1200
	reportTitle = "Location Summary"
	reportObject = "InnSightReportsAll.Report" 'ActiveX DLL public createable class

	'Create the ActiveX report object and set properties one by one.
	Set report = server.createObject(reportObject)

    report.ReportName = "Location"


	'Write Stats
	For intCnt = (LBound(aList) + 1) to UBound(aList)
	  If intCnt > 0 Then
		strSQL = "sp_InsertReportStat " & Session("UserID") & ", " & Session("CompanyID") & ", " & aList(intCnt) & ", '" & Request.QueryString("type") & "'" & ", '" & Request.QueryString("mode") & "'"
		cnSQL.Execute(strSQL)
	  End If
	Next

	'Set report object's locations
	For intCnt = (LBound(aList) + 1) to UBound(aList)
	  report.addLocation(aList(intCnt))
	Next

	Report.ReportName = "Location"
	Report.DSN = "DSN=InnSightReports"

	Report.CompanyID = Session("CompanyID")
	If Request.Form("txtLetterhead") = "Yes" Then
		Report.Letterhead = True
	Else
		Report.Letterhead = False
	End If

	Report.dataDirectory = server.mapPath("reports") 'The "reports" subdirectory should only contain report output files.
	Report.fileType = 0 '0=RDF, 1=RTF, 2=PDF, 3=TXT

	'Not Asynchronous
	Result = report.run()

%>
<%	
	Select Case Request.QueryString("Mode")
		Case "v" 'View
%>
	
	
	<html>
	<head>

	
	<script event="onload" for="window" language=vbscript>
		reportViewer.printer.orientation= 2 '<%=report.orientation%>
		'msgbox "<%=report.orientation%>"
		reportViewer.dataPath="reports/<%=report.filename %>"
		'msgbox "<%=report.filename%>"
		reportViewer.TOCEnabled = False
		reportViewer.tocVisible = False
	</script>

	<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub reportViewer_ToolbarClick( pvTool)
	If pvTool.caption = "Back to Search Screen" Then
		window.history.back(1)
	End If
End Sub

-->
</SCRIPT>

	</head>
	<body bgcolor=silver topmargin="0" leftmargin="0" marginwidth = "0" marginheight = "0" link="black" vlink="black" alink="black">
	<!--#include file = "Header.inc" ---> 
	
	<object class=Bordered id="reportViewer" width=100% height=90%
 	classid="clsid:00C7C2A0-8B82-11D1-8B57-00A0C98CD92B"
	     codebase="arviewer.cab#Version=1,2,0,1057">
	</object>
<%
	Case "p" 'Print
%>
	<html>
	<head>
<META name=VI60_defaultClientScript content=VBScript>

	
	<script event="onload" for="window" language=vbscript>
		reportViewer.printer.orientation= 2
		reportViewer.dataPath="reports/<%=report.filename %>"
		reportViewer.tocVisible = False
		
		reportViewer.PrintReport True
		location.href = "NoReport.asp"
		'window.parent.location = document.referrer
		
	</script>
	
	</head>
	<body bgcolor=silver topmargin="0" leftmargin="0" marginwidth = "0" marginheight = "0" link="black" vlink="black" alink="black">
	<!--#include file = "Header.inc" ---> 
	
	<object class=Bordered id="reportViewer" width=0% height=0%
 	classid="clsid:00C7C2A0-8B82-11D1-8B57-00A0C98CD92B"
	      codebase="../../arviewer.cab#Version=1,2,0,1057">
	</object>

<%
	Case "e" 'E-Mail %>
		<html>
		<head>
		<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		</head>
		<body bgcolor=silver>
		<!--#include file = "Header.inc" ---> 
		<style>
			<!-- .MyFont { font-family: Tahoma; font-size: 11 } -->
		</style>

		<form id="frmSend" name="frmSend" action="EMailLocationReport.asp" method="post">
			<table WIDTH="450" ALIGN="center" BORDER="1" CELLSPACING="1" CELLPADDING="1" bgcolor=#d4d0c8>
				<tr>
					<td><INPUT type="button" value="Send" id=cmdSend name=cmdSend style="height: 40; width: 64; font-family: Tahoma; font-size: 11" onclick="frmSend.submit()"><!--img name="imgSend" id="imgSend" src="images/Send_Up.gif" onmouseover="imgSend.src=imgSendOver.src" onmouseout="imgSend.src=imgSendUp.src" onmousedown="imgSend.src=imgSendDown.src" onmouseup="imgSend.src=imgSendOver.src" onclick="frmSend.submit()"--></tr>
				</tr>
				<tr>
					<td>
						<table bgcolor=#d4d0c8 background border="0">
							<tr>
								<td><p class="MyFont"><img src="images/book.jpg" align="absMiddle" border="0" valign="center" WIDTH="23" HEIGHT="20">&nbsp;To:</p></td>
								<td><input id="to" style="WIDTH: 445px; HEIGHT: 22px" size="63" name="to"></td>
							</tr>
							<tr>
								<td><p class="MyFont"><img style="WIDTH: 23px; HEIGHT: 20px" height="20" alt hspace="0" src="images/book.jpg" width="23" align="absMiddle" useMap border="0" valign="center">&nbsp;cc:</p></td>
								<td><input id="cc" style="WIDTH: 445px; HEIGHT: 22px" size="63" name="cc"></td>
							</tr>
							<tr>
								<td><p class="MyFont">Subject:</p></td>
								<td><input id="subject" style="WIDTH: 445px; HEIGHT: 22px" size="63" name="subject"</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<% for i = lbound(aList) to ubound(aList)
						strstr = strstr & ", " & aList(i)
					next%>
					<td><input id="text1" name="text1"  value="<%=strstr%>" style="WIDTH: 518px; HEIGHT: 138px" size="74"></td>
				</tr>
			</table>
		</form>
		</p>
		</body>
		</html>
<%	Case Else
		Response.Write "Invalid report mode.<BR>"
		Response.End
	end select
	
	'We're done with ActiveX server object.
	Set report = NOTHING
	
%>

</body>
</html>
<%if Request.QueryString("Mode") <> "e" then%>
<SCRIPT LANGUAGE=vbscript>
<!--
		reportViewer.ToolBar.Tools.Add "Back to Search Screen"
		reportViewer.ToolBar.Tools(4).visible = False
		reportViewer.ToolBar.Tools(5).visible = False
		reportViewer.ToolBar.Tools(6).visible = False
		reportViewer.ToolBar.Tools(12).visible = False
		reportViewer.ToolBar.Tools(13).visible = False
-->
</SCRIPT>
<%end if%>
