<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Product:  Ban Man Pro
'   Author:   Joe Rohrbach of Brookfield Consultants
'   Notes:    None
'                  
'
'                         COPYRIGHT NOTICE
'
'   The contents of this file are protected under the United States
'   copyright laws as an unpublished work, and are confidential and
'   proprietary to Brookfield Consultants.  Its use or disclosure in 
'   whole or in part without the expressed written permission of 
'   Brookfield Consultants is prohibited.
'
'   (c) Copyright 1999 by Brookfield Consultants.  All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	
	strMessage=""
	blnFoundError=False

	Select Case strTask
		Case "AddNew" 
			If Application("BanManProMultiSite")=True Then
				If CLng(Session("BanManProSiteID"))=0 Then
					Response.Write "<p align=center>You must first select a site then click go."
					blnFoundError=True
				End If
			End If
			'obtain list of campaigns
			strSQL="SELECT Campaigns.*, Advertisers.CompanyName FROM Campaigns INNER JOIN Advertisers ON Campaigns.AdvertiserID = Advertisers.AdvertiserID Where (Campaigns.UserID=" & CLng(Session("BanManProSiteID")) & " Or Campaigns.UserID=0) AND (CampaignDistribution='Normal' OR CampaignDistribution='Weighted')  AND CampaignDistribution<>'Keyword' ORDER BY Advertisers.CompanyName ASC,Campaigns.[CampaignName] ASC"
			Set rsCampaigns=connBanManPro.Execute(strSQL)
			If Application("SlotOption")=True Then
				'obtain list of Slot Campaigns
				strSQL="SELECT Campaigns.*, Advertisers.CompanyName FROM Campaigns INNER JOIN Advertisers ON Campaigns.AdvertiserID = Advertisers.AdvertiserID Where (Campaigns.UserID=" & CLng(Session("BanManProSiteID")) & " Or Campaigns.UserID=0) AND CampaignDistribution='Weighted' ORDER BY Advertisers.CompanyName ASC,Campaigns.[CampaignName] ASC"
				Set rsEvenCampaigns=connBanManPro.Execute(strSQL)
			Else
				'obtain list of Even Campaigns
				strSQL="SELECT Campaigns.*, Advertisers.CompanyName FROM Campaigns INNER JOIN Advertisers ON Campaigns.AdvertiserID = Advertisers.AdvertiserID Where (Campaigns.UserID=" & CLng(Session("BanManProSiteID")) & " Or Campaigns.UserID=0) AND CampaignDistribution='Even' ORDER BY Advertisers.CompanyName ASC,Campaigns.[CampaignName] ASC"
				Set rsEvenCampaigns=connBanManPro.Execute(strSQL)
			End If
			'obtain list of defaults
			strSQL="SELECT Campaigns.*, Advertisers.CompanyName FROM Campaigns INNER JOIN Advertisers ON Campaigns.AdvertiserID = Advertisers.AdvertiserID Where CampaignSiteDefault<>0 AND (Campaigns.UserID= " & CLng(Session("BanManProSiteID")) & " Or Campaigns.UserID=0)"
			Set rsAllDefaults=connBanManPro.Execute(strSQL)
			If Application("SlotOption")=True Then%>
				<!--#include file="addanewzoneslot.asp"-->
			<% Else %>
				<!--#include file="addanewzone.asp"-->
			<% End If %>
			<%
			Set rsCampaigns=Nothing
			Set rsEvenCampaigns=Nothing
			Set rsAllDefaults=Nothing
		Case "Edit"
			'edit record
			If Trim(strZoneID) <> "" Then
				strSQL2="SELECT * FROM Zones WHERE Zones.ZoneID=" & strZoneID & " AND UserID=" & CLng(Session("BanManProSiteID"))  
				Set rsz=connBanManPro.Execute(strSQL2)
				'obtain list of Campaigns Weighted
				'strSQL="SELECT * FROM Campaigns Where (Campaigns.UserID=" & CLng(Session("BanManProSiteID")) & "Or Campaigns.UserID=0) AND CampaignDistribution<>'Even' AND CampaignDistribution<>'Keyword' ORDER BY Advertisers.CompanyName ASC,Campaigns.[CampaignName] ASC"
				strSQL="SELECT Campaigns.*, Advertisers.CompanyName FROM Campaigns INNER JOIN Advertisers ON Campaigns.AdvertiserID = Advertisers.AdvertiserID Where ( Campaigns.UserID=" & CLng(Session("BanManProSiteID")) & "Or Campaigns.UserID=0) AND (CampaignDistribution='Normal' OR CampaignDistribution='Weighted')  AND CampaignDistribution<>'Keyword' ORDER BY Advertisers.CompanyName ASC,Campaigns.[CampaignName] ASC"
				Set rsCampaigns=connBanManPro.Execute(strSQL)
				If Application("SlotOption")=True Then
					'obtain list of Slot Campaigns
					strSQL="SELECT Campaigns.*, Advertisers.CompanyName FROM Campaigns INNER JOIN Advertisers ON Campaigns.AdvertiserID = Advertisers.AdvertiserID Where (Campaigns.UserID=" & CLng(Session("BanManProSiteID")) & " Or Campaigns.UserID=0) AND CampaignDistribution='Weighted' ORDER BY Advertisers.CompanyName ASC,Campaigns.[CampaignName] ASC"
					Set rsEvenCampaigns=connBanManPro.Execute(strSQL)
				Else
					'obtain list of Even Campaigns
					strSQL="SELECT Campaigns.*, Advertisers.CompanyName FROM Campaigns INNER JOIN Advertisers ON Campaigns.AdvertiserID = Advertisers.AdvertiserID Where (Campaigns.UserID=" & CLng(Session("BanManProSiteID")) & " Or Campaigns.UserID=0) AND CampaignDistribution='Even' ORDER BY Advertisers.CompanyName ASC,Campaigns.[CampaignName] ASC"
					Set rsEvenCampaigns=connBanManPro.Execute(strSQL)
				End If				'obtain ZoneCampaigns
				strSQL="SELECT * FROM ZoneCampaigns WHERE ZoneCampaigns.[ZoneID]=" & strZoneID & " AND (UserID=" & CLng(Session("BanManProSiteID")) & " Or UserID=0)"
				Set rsZoneCampaigns=connBanManPro.Execute(strSQL)
				'obtain list of defaults
				strSQL="SELECT Campaigns.*, Advertisers.CompanyName FROM Campaigns INNER JOIN Advertisers ON Campaigns.AdvertiserID = Advertisers.AdvertiserID Where CampaignSiteDefault<>0 AND (Campaigns.UserID= " & CLng(Session("BanManProSiteID")) & " Or Campaigns.UserID=0)"
				Set rsAllDefaults=connBanManPro.Execute(strSQL)
				'obtain default for this zone
				strSQL="Select * From ZoneDefaults Where ZoneID=" & strZoneID & " AND UserID= " & CLng(Session("BanManProSiteID"))
				Set rsSelectedDefaults=connBanManPro.Execute(strSQL)
				If Application("SlotOption")=True Then%>
					<!--#include file="addanewzoneslot.asp"-->
				<% Else %>
					<!--#include file="addanewzone.asp"-->
				<% End If 
				Set rsz=Nothing
				Set rsCampaigns=Nothing
				Set rsEvenCampaigns=Nothing
				Set rsZoneCampaigns=Nothing
				Set rsAllDefaults=Nothing
				Set rsSelectedDefaults=Nothing
			End If	
		Case "Insert"
			'error checks *************************
			If Trim(Request.Form("ZoneDescription"))="" Then %>
				<p>Zone Description is a required field</p>
				<% blnFoundError=True
			End If	
			'Check if zone name is already in use
			strZoneName=FixBlank(Request.Form("ZoneDescription"))
			strSQL="Select ZoneDescription From Zones Where ZoneDescription='" & strZoneName & "' And UserID= " & CLng(Session("BanManProSiteID"))
			Set rsTemp=connBanManPro.Execute(strSQL)
			If Not rsTemp.EOF Then%>
				<p align="center">Zone Name is already in use for another zone.</p>
				<% blnFoundError=True
			End If
			Set rsTemp=Nothing
		    If (UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN") And blnFoundError<>True Then

			If Clng(Request.Form("DefaultCampaign")) <> 0 Then
				blnIncludeDefault=1
			Else
				blnIncludeDefault=0
			End If
			'end error checks *********************
 			strSQL2="INSERT INTO Zones ("
			strSQL2=strSQL2 & "ZoneDescription,"
			strSQL2=strSQL2 & "ZoneMode,"
			strSQL2=strSQL2 & "ZoneIncludeDefaults,"
			strSQL2=strSQL2 & "ZonePageURL,UserID,ZoneWidth,ZoneHeight) VALUES ('"
			strSQL2=strSQL2 & FixBlank(Request.Form("ZoneDescription")) & "','" 
			strSQL2=strSQL2 & FixBlank(Request.Form("ZoneMode")) & "'," 
			strSQL2=strSQL2 & blnIncludeDefault & ",'" 
			strSQL2=strSQL2 & FixBlank(Request.Form("ZonePageURL")) & "'," & CLng(Session("BanManProSiteID")) 
			strSQL2=strSQL2 & "," & Request.Form("ZoneWidth") & "," & Request.Form("ZoneHeight") & ")" 
			connBanManPro.Execute strSQL2

			'get newly added Zone ID
			strSQL2="SELECT Zones.ZoneID FROM Zones WHERE Zones.ZoneDescription In ('" & FixBlank(Request.Form("ZoneDescription")) & "') And Zones.ZonePageURL in ('" & FixBlank(Request.Form("ZonePageURL")) & "') And Zones.UserID=" & CLng(Session("BanManProSiteID")) 		
			Set rsZoneID=connBanManPro.Execute(strSQL2)
			strZoneID=rsZoneID("ZoneID")

			'now add Weighted campaigns to ZoneCampaigns
			intCnt=0
			sngSum=0
			Do Until intCnt=2000
				intCnt=intCnt+1
				strSelected="chkCampaignSelected" & intCnt
				strWeighting="ZoneCampaignWeighting" & intCnt
				If Trim(Request.Form(strSelected)) <> "" Then
					If IsNumeric(Request.Form(strWeighting)) Then
						sngCampWeight=CInt(Request.Form(strWeighting))
					Else
						sngCampWeight=0
					End If
					sngSum=sngSum+Csng(sngCampWeight)
					strSQL2="INSERT INTO ZoneCampaigns ("
					strSQL2=strSQL2 & "ZoneID,"
					strSQL2=strSQL2 & "CampaignID,"
					strSQL2=strSQL2 & "ZoneCampaignWeighting,UserID) VALUES ("
					strSQL2=strSQL2 & strZoneID & ","
					strSQL2=strSQL2 & Request.Form(strSelected) & "," 
					strSQL2=strSQL2 & sngCampWeight & "," & CLng(Session("BanManProSiteID"))  & ")" 
					connBanManPro.Execute strSQL2
				End If
			Loop

			'Add EVEN Campaigns
			varCampaigns=Split(Request.Form("EvenCampaigns"),",")
			intCnt=0
			Do While intCnt<= Ubound(varCampaigns)
				'If Slot Option, get #Slots
				If Application("SlotOption")=True Then
					strSQL="Select CampaignQuantitySold From Campaigns Where CampaignID=" & Clng(varCampaigns(intCnt))
					Set rsTemp=connBanManPro.Execute(strSQL)
					lngSlots=rsTemp("CampaignQuantitySold")
				Else
					lngSlots=0
				End If
				strSQL2="INSERT INTO ZoneCampaigns ("
				strSQL2=strSQL2 & "ZoneID,"
				strSQL2=strSQL2 & "CampaignID,"
				strSQL2=strSQL2 & "ZoneCampaignWeighting,UserID,Even) VALUES ("
				strSQL2=strSQL2 & strZoneID & ","
				strSQL2=strSQL2 & varCampaigns(intCnt) & "," & lngSlots & "," 
				strSQL2=strSQL2 & CLng(Session("BanManProSiteID"))  & ",1)" 
				connBanManPro.Execute strSQL2
				
				intCnt=intCnt+1
			Loop		

			'Add Default Campaigns to ZoneDefaults
			If blnIncludeDefault<>0 Then
				strSQL="Insert Into ZoneDefaults (ZoneID,CampaignID,UserID) Values ("
				strSQL=strSQL & strZoneID & "," 
				strSQL=strSQL & Request.Form("DefaultCampaign") & ","
				strSQL=strSQL & CLng(Session("BanManProSiteID")) & ")"
				connBanManPro.Execute strSQL
			End If
			%>
			<!--#Include File="banmanfunc.asp"-->
			<%
			'Calculate Expected Quantity for Even Campaigns
			If Application("SlotOption")<>True Then
				CalculateBanManProExpectedQuantity
			End If
			If Request.Form("ZoneMode")<> "HTML" Then	
				CreateZoneFile strZoneID,Session("BanManProSiteID")
			End If
			Set rsZoneID=Nothing
			Set rsTemp=Nothing
		    End If
			If blnFoundError=False Then
			If sngSum > 101 Then
				%>
				<p align="center"><font face="Arial" size="5">Added Zone: <%=Request.Form("ZoneDescription")%>.</font></p>
				<p align="center"><font face="Arial" size="5">Alert** Sum of all campaigns must be <= 100, please edit.</font></p>
				<%  	 
			Else
				%>
				<p align="center"><font face="Arial" size="5">Successfully added Zone: <%=Request.Form("ZoneDescription")%>.</font></p>
				<%  	 
			End If
			End If
		Case "Update"
			'error checks *************************
			If Trim(Request.Form("ZoneDescription"))="" Then
				Response.Write "<p align=center>Zone Description is a required field"
				blnFoundError=True
			End If	
			'Check if zone name is already in use
			strZoneName=FixBlank(Request.Form("ZoneDescription"))
			strSQL="Select ZoneDescription From Zones Where ZoneDescription='" & strZoneName & "' And ZoneID<>" & strZoneID & " AND UserID= " & CLng(Session("BanManProSiteID"))
			Set rsTemp=connBanManPro.Execute(strSQL)
			If Not rsTemp.EOF Then%>
				<p align="center">Zone Name is already in use for another zone.</p>
				<% blnFoundError=True
			End If
			Set rsTemp=Nothing
			'end error checks *********************
		    If (UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN") And blnFoundError<>True Then

			If Clng(Request.Form("DefaultCampaign")) <> 0 Then
				blnIncludeDefault=1
			Else
				blnIncludeDefault=0
			End If
 			strSQL2="UPDATE Zones SET "
			strSQL2=strSQL2 & "ZoneDescription='" &  FixBlank(Request.Form("ZoneDescription")) & "',"  
			strSQL2=strSQL2 & "ZoneMode='" &  FixBlank(Request.Form("ZoneMode")) & "'," 
			strSQL2=strSQL2 & "ZoneIncludeDefaults=" &  blnIncludeDefault & ","  
			strSQL2=strSQL2 & "ZonePageURL='" &  FixBlank(Request.Form("ZonePageURL")) & "', "
			strSQL2=strSQL2 & "ZoneWidth=" &  Request.Form("ZoneWidth") & ","
			strSQL2=strSQL2 & "ZoneHeight=" &   Request.Form("ZoneHeight")
			strSQL2=strSQL2 & "WHERE Zones.[ZoneID]=" & strZoneID & " AND UserID=" & CLng(Session("BanManProSiteID")) 
			connBanManPro.Execute strSQL2

			'now Update ZoneCampaigns
			intCnt=0
			'first delete all records for this ZoneID
			strSQL2="DELETE FROM ZoneCampaigns WHERE ZoneCampaigns.[ZoneID]=" & strZoneID & " AND UserID=" & CLng(Session("BanManProSiteID")) & " AND Even=0"
			connBanManPro.Execute strSQL2
			sngSum=0
			Do Until intCnt=2000
				intCnt=intCnt+1
				strSelected="chkCampaignSelected" & intCnt
				strWeighting="ZoneCampaignWeighting" & intCnt
				If Trim(Request.Form(strSelected)) <> "" Then
					If IsNumeric(Request.Form(strWeighting)) Then
						sngCampWeight=CInt(Request.Form(strWeighting))
					Else
						sngCampWeight=0
					End If
					sngSum=sngSum+Csng(sngCampWeight)
					strSQL2="INSERT INTO ZoneCampaigns ("
					strSQL2=strSQL2 & "ZoneID,"
					strSQL2=strSQL2 & "CampaignID,"
					strSQL2=strSQL2 & "ZoneCampaignWeighting,UserID) VALUES ("
					strSQL2=strSQL2 & strZoneID & ","
					strSQL2=strSQL2 & Request.Form(strSelected) & "," 
					strSQL2=strSQL2 & sngCampWeight & "," & CLng(Session("BanManProSiteID")) &  ")" 
					connBanManPro.Execute strSQL2
				End If
			Loop

			'Add EVEN Campaigns
			varCampaigns=Split(Request.Form("EvenCampaigns"),",")
			intCnt=0
			strTemp=" Delete From ZoneCampaigns Where ZoneID=" & Clng(strZoneID) & " And Even<>0  "
			'Must Account for case where all even campaigns are deleted
			If Trim(Request.Form("EvenCampaigns")="") Then
				Set rs=connBanManPro.Execute(strTemp)
			End If
			Do While intCnt<= Ubound(varCampaigns)
				'Determine if exists
				strTemp=strTemp & " And CampaignID<>" &  Clng(varCampaigns(intCnt))
				strSQL="Select ZoneCampaignWeighting From ZoneCampaigns Where ZoneID=" & Clng(strZoneID)
				strSQL=strSQL & " And CampaignID=" & Clng(varCampaigns(intCnt)) & " And Even<>0"
				Set rs=connBanManPro.Execute(strSQL)
				If Not rs.EOF Then
					'already exists, don't touch
				Else
					If Application("SlotOption")=True Then
						strSQL="Select CampaignQuantitySold From Campaigns Where CampaignID=" & Clng(varCampaigns(intCnt))
						Set rsTemp=connBanManPro.Execute(strSQL)
						lngSlots=rsTemp("CampaignQuantitySold")
					Else
						lngSlots=0
					End If
					strSQL2="INSERT INTO ZoneCampaigns ("
					strSQL2=strSQL2 & "ZoneID,"
					strSQL2=strSQL2 & "CampaignID,"
					strSQL2=strSQL2 & "ZoneCampaignWeighting,UserID,Even) VALUES ("
					strSQL2=strSQL2 & strZoneID & ","
					strSQL2=strSQL2 & varCampaigns(intCnt) & "," & lngSlots & "," 
					strSQL2=strSQL2 & CLng(Session("BanManProSiteID"))  & ",1)" 
					connBanManPro.Execute strSQL2
				End If
				intCnt=intCnt+1
			Loop	
			'Add Default Campaigns to ZoneDefaults, first delete any existing
			strSQL="Delete from ZoneDefaults Where ZoneID=" & strZoneID
			connBanManPro.Execute strSQL
			If blnIncludeDefault<>0 Then
				strSQL="Insert Into ZoneDefaults (ZoneID,CampaignID,UserID) Values ("
				strSQL=strSQL & strZoneID & "," 
				strSQL=strSQL & Request.Form("DefaultCampaign") & ","
				strSQL=strSQL & CLng(Session("BanManProSiteID")) & ")"
				connBanManPro.Execute strSQL
			End If


			'Delete Even Campaigns Which have been removed
			If intCnt>0 Then
				connBanManPro.Execute strTemp
			End If
			%>
			<!--#Include File="banmanfunc.asp"-->
			<%
			'Calculate Expected Quantity for Even Campaigns
			If Application("SlotOption")<>True Then
				CalculateBanManProExpectedQuantity
			End If


			If Request.Form("ZoneMode")<> "HTML" Then	
				CreateZoneFile strZoneID,Session("BanManProSiteID")
			End If
			Set rs=Nothing
			Set rsTemp=Nothing
		    End If
			If blnFoundError=False Then
			If sngSum > 101 Then
				%>
				<p align="center"><font face="Arial" size="5">Updated Zone.</font></p>
				<p align="center"><font face="Arial" size="5">Alert** Sum of all campaigns must be <= 100, please edit.</font></p>
				<%  	 
			Else
				%>
				<p align="center"><font face="Arial" size="5">Successfully updated zone.</font></p>
				<%  	 
			End If
			End If
		Case "Delete"
			'delete entry
			If Trim(strZoneID) <> "" Then
			    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then
				strSQL2="DELETE FROM Zones WHERE Zones.[ZoneID]=" & strZoneID & " AND UserID=" & CLng(Session("BanManProSiteID"))
				connBanManPro.Execute strSQL2
				strSQL="Delete from ZoneDefaults Where ZoneID=" & strZoneID
				connBanManPro.Execute strSQL
			    End If
				%>
				<p align="center"><font face="Arial" size="5">Record Deleted.</font></p>
				<%  	
			Else 	%>
				<p align="center"><font face="Arial" size="5">Nothing to Delete.</font></p>
				<%
			End If		
		Case "ViewAll", ""
			'Check if viewing by letter
			If Trim(Request("Letter")) <> "" Then
				If Trim(Request("Letter")) <> "Other" Then
					strSQL2="SELECT * FROM Zones WHERE (((Zones.[ZoneDescription] Like '" & UCase(Request("Letter")) & "%') OR (Zones.[ZoneDescription] Like '" & LCase(Request("Letter")) & "%'))) AND UserID=" & CLng(Session("BanManProSiteID")) & "  ORDER BY Zones.[ZoneDescription] ASC"
				Else
					strSQL2="SELECT * FROM Zones WHERE (((Zones.[ZoneDescription] < 'a%') OR (Zones.[ZoneDescription] > 'z%'))) AND UserID=" & CLng(Session("BanManProSiteID")) & "  ORDER BY Zones.[ZoneDescription] ASC"
				End If
			Else
				strSQL2="SELECT * FROM Zones Where UserID=" & CLng(Session("BanManProSiteID")) &  " ORDER BY Zones.[ZoneDescription] ASC" 
			End If

			Set rsz=connBanManPro.Execute(strSQL2)
			'update zone statistics
			If Not rsz.EOF Then
				strMessage="Listing of all Zones in Database."	
				'call include file and create table of all data
				%>
				<!--#include file="showallzones.asp"-->
				<%
			End If
			Set rsz=Nothing
		Case "ViewCode"
			strSQL2="SELECT * FROM Zones WHERE Zones.[ZoneID]=" & strZoneID & " And UserID=" & CLng(Session("BanManProSiteID"))
			Set rsz=connBanManPro.Execute(strSQL2)
			If Not rsz.EOF Then
				'call include file and create table of all data
				%>
				<!--#include file="viewadcode.asp"-->
				<%
			End If
			Set rsz=Nothing
	End Select
%>


<% ''''''''''''''''''''Change blank fields to " "   '''''''''''''''''''''''''''''''''''''''''''
Function FixBlank(strParameter)
	If Trim(strParameter)="" Then
		FixBlank=" "
	Else
		FixBlank=Replace(strParameter, "'", "''")
		FixBlank=Trim(FixBlank)
	End If
End Function  %>
<% ''''''''''''''''''''Create Zone File''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateZoneFile(strZoneID,SiteID)

On Error Resume Next

        strPath=Application("ServerPath")
	'Server.MapPath("/zones/banmanzone" & strZoneID & ".asp")
	Set fs = CreateObject("Scripting.FileSystemObject")
        Set objFile = fs.CreateTextFile(strPath & "zones\banmanzone" & strZoneID & ".asp", True)
	objFile.WriteLine(Chr(60) & "%")
	'objFile.WriteLine("Dim strZoneID")
	'objFile.WriteLine("Dim strTask")
	objFile.WriteLine("strZoneID=" & strZoneID)
	objFile.WriteLine("lngBMPSiteID=" & SiteID)
	objFile.WriteLine("strTask=" & Chr(34) & "Get" & Chr(34))
	objFile.WriteLine("%" & Chr(62))
	strInclude=Chr(60) & "!--#include virtual=" & Chr(34) & getFilePath() & "banman.asp" & Chr(34) & "-->"
	objFile.WriteLine(strInclude)
        objFile.Close

	'check for error
	If Err.Number >0 Then
		'An error occurred attempting to write file
		'include file with directions on fixing problem
		%>
		<!--#include file="errorzone.asp"-->
		<%
	End If
	Err.Clear

End Sub
%><%
Function getFilePath()
	Dim lsPath, arPath

	' Obtain the virtual file path. The SCRIPT_NAME
	' item in the ServerVariables collection in the
	' Request object has the complete virtual file path
	lsPath = Request.ServerVariables("SCRIPT_NAME")
                           
	' Split the path along the /s. This creates an
	' This creates an one-dimensional array 
	arPath = Split(lsPath, "/")

	' Set the last item in the array to blank string
	' (The last item actually is the file name)
	arPath(UBound(arPath,1)) = ""
	
	' Join the items in the array. This will
	' give you the virtual path of the file
	GetFilePath = Join(arPath, "/")
End Function

%>
<% ''''''''''''''''''''Set False Check box to 0 " "   '''''''''''''''''''''''''''''''''''''''''''
Function SetTrueFalse(strParameter)
	If Trim(strParameter)="-1" Then
		SetTrueFalse=-1
	Else
		SetTrueFalse=0
	End If
End Function %>