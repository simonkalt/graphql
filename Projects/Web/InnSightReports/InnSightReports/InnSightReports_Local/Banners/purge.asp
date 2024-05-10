<%Server.ScriptTimeout=24000
On Error Resume Next 
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body>
<%If Request("Task2")<> "Perform" Then %>
<p align="center"><font face="Arial" size="3">Use this tool to purge the
database.&nbsp; Purging<br>
will remove old statistics from the database<br>
which are no longer needed.&nbsp; Stats will be deleted<br>
for all campaigns that have been removed since<br>
the last database purging.</font></p>
<p align="center"><a href="tools.asp?Task=Purge&amp;Task2=Perform"><img border="0" src="images/PurgeNow.gif" WIDTH="135" HEIGHT="30"></a></p>
<%Else

	strSQL="Select CampaignID From Campaigns"
	Set rsCampaigns=connBanManPro.Execute(strSQL)

	'Build Where clause
	strPlus=""
	strString=" CampaignID<> 0 AND "
	Do While Not rsCampaigns.EOF

		strString=strString & strPlus & " CampaignID <> " & rsCampaigns("CampaignID")
		strPlus=" AND "	
		rsCampaigns.MoveNext
	Loop

	strCount="Select Count(ID) As CountOfClicks From Clicks Where " & strString
	Set rsCount=connBanManPro.Execute(strCount)
	If rsCount("CountOfClicks") > 0 Then

		intUpper=rsCount("CountOfClicks")
		intCount=0
		intCnt=0
		strCount="Select ID From Clicks Where " & strString & " ORDER By Id Asc"
		Set rsCount=connBanManPro.Execute(strCount)
		Do While Not rsCount.EOF
			'move forward 1000 ID counts
			Do While intCnt< 1000
				rsCount.MoveNext
				intCnt=intCnt+1
			Loop
			intCnt=0

			strSQL="Delete From Clicks Where " & strString & " And ID < " & rsCount("ID")
			connBanManPro.Execute strSQL,,adExecuteNoRecords
			Response.Write "Deleted " & intCount & " - " & intCount +1000 & " Clicks <br>"
			intCount=intCount+1000
		Loop
	
	End If

	strCount="Select Count(ImpressionID) As CountOfImpressions From Impressions Where " & strString
	Set rsCount=connBanManPro.Execute(strCount)
	strSQL="Delete From Impressions Where " & strString
	connBanManPro.Execute strSQL,,adExecuteNoRecords
	%>
	<p align="center"><b><font face="Arial" size="5">Deleted <%=rsCount("CountOfImpressions")%> Impression Records</font></b></p>
	<%
	%>
	<p align="center"><b><font face="Arial" size="5">Purge successfully completed.</font></b></p>
	<%
End If %>

</body>

</html>
