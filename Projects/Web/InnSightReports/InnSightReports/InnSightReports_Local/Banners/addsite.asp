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

	Select Case strTask
		Case "AddNew" 
			'Add A New Site
			%>
			<!--#include file="addanewsite.asp"-->
			<%
		Case "Edit"
			'edit record
			If Trim(strSiteID) <> "" Then
				strSQL2="SELECT * FROM BanManProWebSites Where SiteID=" & strSiteID 
				Set rss=connBanManPro.Execute(strSQL2)
				If Not rss.EOF Then
 					%>
					<!--#include file="addanewsite.asp"-->
					<%
					Set rss=Nothing
				End If
			End If	
		Case "Insert"

		    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then
			'end error checks *********************
 			strSQL2="INSERT INTO BanManProWebSites ("
			strSQL2=strSQL2 & "SiteName,"
			strSQL2=strSQL2 & "SiteURL) VALUES ('"
			strSQL2=strSQL2 & FixBlank(Request.Form("SiteName")) & "','" 
			strSQL2=strSQL2 & FixBlank(Request.Form("SiteURL")) & "')"
			connBanManPro.Execute strSQL2

		    End If
			%>
			<p align="center"><font face="Arial" size="3">Successfully added Site: <%=Request.Form("SiteName")%>.</font></p>
			<%  	 
		Case "Update"
		    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then
 			strSQL2="UPDATE BanManProWebSites SET "
			strSQL2=strSQL2 & "SiteName='" &  FixBlank(Request.Form("SiteName")) & "',"  
			strSQL2=strSQL2 & "SiteURL='" &  FixBlank(Request.Form("SiteURL"))  & "' "
			strSQL2=strSQL2 & "WHERE SiteID=" & strSiteID
			connBanManPro.Execute strSQL2
		    End If
			%>
			<p align="center"><font face="Arial" size="3">Successfully updated site.</font></p>
			<%  	 
		Case "Delete"
			'delete entry
			If Trim(strSiteID) <> "" Then
			    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then
				strSQL2="DELETE FROM BanManProWebSites WHERE SiteID=" & strSiteID
				connBanManPro.Execute strSQL2,,AdExecuteNoRecords
				'This could be damaging but delete all Advertisers which will cascade to all other fields
				strSQL2="DELETE FROM Advertisers Where UserID=" & strSiteID
				connBanManPro.Execute strSQL2,,AdExecuteNoRecords
				'This could be damaging but delete all Zones which will cascade to all other fields
				strSQL2="DELETE FROM Zones Where UserID=" & strSiteID
				connBanManPro.Execute strSQL2,,AdExecuteNoRecords
			    End If
				%>
				<p align="center"><font face="Arial" size="5">Record Deleted.</font></p>
				<%  	
			Else 	%>
				<p align="center"><font face="Arial" size="5">Nothing to Delete.</font></p>
				<%
			End If		
		Case "ViewAll", ""
			'get sites
			strSQL2="SELECT * FROM BanManProWebSites" 
			Set rss=connBanManPro.Execute(strSQL2)
			If Not rss.EOF Then
				strMessage="Listing of all sites in Database."	
				'call include file and create table of all data
				%>
				<!--#include file="showallsites.asp"-->
				<%
			End If

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
<% ''''''''''''''''''''Set False Check box to 0 " "   '''''''''''''''''''''''''''''''''''''''''''
Function SetTrueFalse(strParameter)
	If Trim(strParameter)="-1" Then
		SetTrueFalse=-1
	Else
		SetTrueFalse=0
	End If
End Function %>