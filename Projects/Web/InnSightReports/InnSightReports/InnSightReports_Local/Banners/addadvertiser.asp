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
			If Application("BanManProMultiSite")=True Then
				If CLng(Session("BanManProSiteID"))=0 Then
					Response.Write "<p align=center>You must first select a site then click go."
					Response.End
				End If
			End If
			%>
			<!--#include file="addanewadvertiser.asp"-->
			<%
		Case "Edit"
			'edit record
			If Trim(strAdvertiserID) <> "" Then
				strSQL2="SELECT * FROM Advertisers WHERE Advertisers.[AdvertiserID]=" & strAdvertiserID 
				Set rss=connBanManPro.Execute(strSQL2)
				If Not rss.EOF Then
 					%>
					<!--#include file="addanewadvertiser.asp"-->
					<%
				End If
				Set rss=Nothing
			End If	
		Case "Insert"
			'error checks *************************
			If Trim(Request.Form("CompanyName"))="" Then
				Response.Write "Company Name is a required field"
				Response.End
			End If	
			'end error checks *********************
		    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then
 			strSQL2="INSERT INTO Advertisers ("
			strSQL2=strSQL2 & "AdvertiserDesc,"
			strSQL2=strSQL2 & "LoginName,"
			strSQL2=strSQL2 & "LoginPassword,"
			strSQL2=strSQL2 & "Email,"
			strSQL2=strSQL2 & "Contact,"
			strSQL2=strSQL2 & "CompanyWebSite,"
			strSQL2=strSQL2 & "CompanyName,"
			strSQL2=strSQL2 & "CompanyAddress1,"
			strSQL2=strSQL2 & "CompanyAddress2,"
			strSQL2=strSQL2 & "Country,"
			strSQL2=strSQL2 & "City,"
			strSQL2=strSQL2 & "State,"
			strSQL2=strSQL2 & "Zip,"
			strSQL2=strSQL2 & "Telephone,"
			strSQL2=strSQL2 & "Fax,UserID,DailyReport,WeeklyReport) VALUES ('"
			strSQL2=strSQL2 & FixBlank(Request.Form("AdvertiserDesc")) & "','" 
			strSQL2=strSQL2 & FixBlank(Request.Form("LoginName")) & "','"  
			strSQL2=strSQL2 & FixBlank(Request.Form("LoginPassword")) & "','"  
			strSQL2=strSQL2 & FixBlank(Request.Form("Email")) & "','" 
			strSQL2=strSQL2 & FixBlank(Request.Form("Contact")) & "','" 
        		strSQL2=strSQL2 & FixBlank(Request.Form("CompanyWebSite")) & "','"  
			strSQL2=strSQL2 & FixBlank(Request.Form("CompanyName")) & "','"  
			strSQL2=strSQL2 & FixBlank(Request.Form("CompanyAddress1")) & "','"  
			strSQL2=strSQL2 & FixBlank(Request.Form("CompanyAddress2")) & "','"  
			strSQL2=strSQL2 & FixBlank(Request.Form("Country")) & "','" 
        		strSQL2=strSQL2 & FixBlank(Request.Form("City")) & "','"  
			strSQL2=strSQL2 & FixBlank(Request.Form("State")) & "','"  
			strSQL2=strSQL2 & FixBlank(Request.Form("Zip")) & "','"  
			strSQL2=strSQL2 & FixBlank(Request.Form("Telephone")) & "','"  
			strSQL2=strSQL2 & FixBlank(Request.Form("Fax"))  & "'," 
			If Request.Form("RunOfNetwork")="ON" Then
				strSQL2=strSQL2 & Clng(0) 
			Else
				strSQL2=strSQL2 & CLng(Session("BanManProSiteID")) 
			End If
			strSQL2=strSQL2 & "," & SetTrueFalse(Request.Form("DailyReport")) & "," & SetTrueFalse(Request.Form("WeeklyReport")) & ")"
			connBanManPro.Execute strSQL2
		    End If
			'Response.Write strSQL2
			%>
			<p align="center"><font face="Arial" size="5">Successfully added new advertiser: <%=Request.Form("CompanyName")%>.</font></p>
			<%  	 
		Case "Update"
		    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then
 			strSQL2="UPDATE Advertisers SET "
			strSQL2=strSQL2 & "AdvertiserDesc='" &  FixBlank(Request.Form("AdvertiserDesc")) & "',"  
			strSQL2=strSQL2 & "LoginName='" &  FixBlank(Request.Form("LoginName")) & "',"
			strSQL2=strSQL2 & "LoginPassword='" &  FixBlank(Request.Form("LoginPassword")) & "',"
			strSQL2=strSQL2 & "Email='" &  FixBlank(Request.Form("Email")) & "'," 
			strSQL2=strSQL2 & "Contact='" &  FixBlank(Request.Form("Contact")) & "'," 
			strSQL2=strSQL2 & "CompanyWebSite='" &  FixBlank(Request.Form("CompanyWebSite")) & "'," 
			strSQL2=strSQL2 & "CompanyName='" &  FixBlank(Request.Form("CompanyName")) & "'," 
			strSQL2=strSQL2 & "CompanyAddress1='" &  FixBlank(Request.Form("CompanyAddress1")) & "'," 
			strSQL2=strSQL2 & "CompanyAddress2='" &  FixBlank(Request.Form("CompanyAddress2")) & "'," 
			strSQL2=strSQL2 & "Country='" &  FixBlank(Request.Form("Country")) & "'," 
			strSQL2=strSQL2 & "City='" &  FixBlank(Request.Form("City")) & "'," 
			strSQL2=strSQL2 & "State='" &  FixBlank(Request.Form("State")) & "'," 
			strSQL2=strSQL2 & "Zip='" &  FixBlank(Request.Form("Zip")) & "'," 
			strSQL2=strSQL2 & "Telephone='" &  FixBlank(Request.Form("Telephone")) & "'," 
			strSQL2=strSQL2 & "Fax='" &  FixBlank(Request.Form("Fax")) & "',"
			If Request.Form("RunOfNetwork")="ON" Then
				strSQL2=strSQL2 & "UserID=0,"
			Else
				strSQL2=strSQL2 & "UserID=" & CLng(Session("BanManProSiteID"))  & ","
			End If
			strSQL2=strSQL2 & "DailyReport=" &  SetTrueFalse(Request.Form("DailyReport"))  & ","
			strSQL2=strSQL2 & "WeeklyReport=" &  SetTrueFalse(Request.Form("WeeklyReport"))  
			strSQL2=strSQL2 & " WHERE Advertisers.[AdvertiserID] =" & strAdvertiserID
			connBanManPro.Execute strSQL2
		    End If
			%>
			<p align="center"><font face="Arial" size="5">Record For <%=Request.Form("CompanyName")%> Updated.</font></p>
			<%  	 
			'Response.Redirect "Advertisers.asp?Task=Details&AdvertiserID=" & strAdvertiserID
		Case "Delete"
			'delete entry
			If Trim(strAdvertiserID) <> "" Then
			    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then
				strSQL2="DELETE FROM Advertisers WHERE Advertisers.[AdvertiserID]=" & strAdvertiserID 
				connBanManPro.Execute strSQL2
			    End If
				%>
				<p align="center"><font face="Arial" size="5">Record Deleted.</font></p>
				<%  	
			Else 	%>
				<p align="center"><font face="Arial" size="5">Nothing to Delete.</font></p>
				<%
			End If		
		Case "ViewAll",""
			'Check if viewing by letter
			If Trim(Request("Letter")) <> "" Then
				If Trim(Request("Letter")) <> "Other" Then
					strSQL2="SELECT * FROM Advertisers WHERE (((Advertisers.CompanyName Like '" & UCase(Request("Letter")) & "%') OR (Advertisers.CompanyName Like '" & LCase(Request("Letter")) & "%'))) AND (UserID=" & CLng(Session("BanManProSiteID")) & " OR UserID=0)  ORDER BY Advertisers.[CompanyName] ASC"
				Else
					strSQL2="SELECT * FROM Advertisers WHERE (((Advertisers.CompanyName < 'a%') OR (Advertisers.CompanyName > 'z%'))) AND (UserID=" & CLng(Session("BanManProSiteID")) & " OR UserID=0)  ORDER BY Advertisers.[CompanyName] ASC"
				End If
			Else
				strSQL2="SELECT * FROM Advertisers Where (UserID=" & CLng(Session("BanManProSiteID")) & " OR UserID=0)  ORDER BY Advertisers.[CompanyName] ASC" 
			End If
			Set rss=connBanManPro.Execute(strSQL2)
			If Not rss.EOF Then
				strMessage="Listing of all Advertisers in Database."	
				'call include file and create table of all data
				%>
				<!--#include file="showalladvertisers.asp"-->
				<%
			End If
			Set rss=Nothing
		Case "Details"
			'show advertiser entry
			If Trim(strAdvertiserID) <> "" Then
				strSQL2="SELECT * FROM Advertisers WHERE Advertisers.[AdvertiserID]=" & strAdvertiserID   
				Set rss=connBanManPro.Execute(strSQL2)
				If Not rss.EOF Then
					strMessage="Advertiser information for: " & rss("CompanyName")  	
					'call include file to create table of data
					%>
					<!--#include file="showadvertiserdetails.asp"-->
					<%
				End If
				Set rss=Nothing
			End If	
	End Select


''''''''''''''''''''Change blank fields to " "   '''''''''''''''''''''''''''''''''''''''''''
Function FixBlank(strParameter)
	If Trim(strParameter)="" Then
		FixBlank=" "
	Else
		FixBlank=Replace(strParameter, "'", "''")
		FixBlank=Trim(FixBlank)
	End If
End Function
''''''''''''''''''''Set False Check box to 0 " "   '''''''''''''''''''''''''''''''''''''''''''
Function SetTrueFalse(strParameter)
	If Trim(strParameter)="-1" Then
		SetTrueFalse=-1
	Else
		SetTrueFalse=0
	End If
End Function  
%>