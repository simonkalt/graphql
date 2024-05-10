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
		Case "Advanced"
			'Allow user to ad code rather than banner information
			'obtain list of advertisers
			strSQL="SELECT * FROM Advertisers WHERE (UserID=" & CLng(Session("BanManProSiteID")) & " OR UserID=0) ORDER BY Advertisers.[CompanyName] ASC"
			Set rsa=connBanManPro.Execute(strSQL)
			%>
			<!--#include file="advancedbanner.asp"-->
			<%
			Set rsa=Nothing
		Case "ViewCode"
			'show the code for this banner in a text box
			'edit record
			If Trim(strBannerID) <> "" Then
				strSQL2="SELECT Banners.UserID,Banners.AdDescription,Banners.AdWidth,Banners.AdHeight,Banners.BannerID,Banners.AdFragment,Banners.AdImageURL,Banners.AdNewWindow,Banners.AdTargetURL,Advertisers.CompanyName,Advertisers.AdvertiserID,Banners.AdCode,Banners.AdCodeNetscape "	
				strSQL2=strSQL2 & " FROM Advertisers RIGHT JOIN Banners ON Advertisers.AdvertiserID = Banners.AdvertiserID"
				strSQL2=strSQL2 & " WHERE (((Banners.BannerID)=" & strBannerID & "))"
				Set rsb=connBanManPro.Execute(strSQL2)
				'obtain list of advertisers
				strSQL="SELECT * FROM Advertisers Where (UserID=" & CLng(Session("BanManProSiteID")) & " Or UserID=0) ORDER BY Advertisers.[CompanyName] ASC"
				Set rsa=connBanManPro.Execute(strSQL)
				If Not rsb.EOF Then
 					%>
					<!--#include file="advancedbanner.asp"-->
					<%
				End If
				Set rsa=Nothing
				Set rsb=Nothing
			End If	
		Case "InsertAdvanced"
		    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then
			If Request.Form("Netscape4")="-1" Then
				lngNetscape4=-1
			Else
				lngNetscape4=0
			End If
 			strSQL2="INSERT INTO Banners ("
			strSQL2=strSQL2 & "AdvertiserID,"
			strSQL2=strSQL2 & "AdDescription,"
			strSQL2=strSQL2 & "AdFragment,"
			strSQL2=strSQL2 & "AdImageURL,"
			strSQL2=strSQL2 & "AdTargetURL,"
			strSQL2=strSQL2 & "AdNewWindow,"
			strSQL2=strSQL2 & "AdWidth,"
			strSQL2=strSQL2 & "AdHeight,"
			strSQL2=strSQL2 & "AdCode,UserID,AdCodeNetscape) VALUES ("
			strSQL2=strSQL2 & FixBlank(Request.Form("AdvertiserID")) & ",'" 
			strSQL2=strSQL2 & FixBlank(Request.Form("AdDescription")) & "'," 
			strSQL2=strSQL2 & "-1,'"
			strSQL2=strSQL2 & FixBlank(Request.Form("ImageSource")) & "','" 
			strSQL2=strSQL2 & FixBlank(Request.Form("TargetURL")) & "'," 
			strSQL2=strSQL2 & lngNetscape4 & "," 
			strSQL2=strSQL2 & FixZero(Request.Form("AdWidth")) & "," 
			strSQL2=strSQL2 & FixZero(Request.Form("AdHeight")) & ",'" 
			strSQL2=strSQL2 & FixBlank(Request.Form("AdCode")) & "',"
			If Request.Form("RunOfNetwork")="ON" Then
				'Check if Advertiser is global
				strTemp="Select AdvertiserID From Advertisers Where AdvertiserID=" & Clng(Request.Form("AdvertiserID")) & " And UserID=0"
				set rsTemp=connBanManPro.Execute(strTEmp)
				If Not rsTemp.EOF Then
					strSQL2=strSQL2 & Clng(0) 
				Else
					strSQL2=strSQL2 & CLng(Session("BanManProSiteID")) 
				End If
				Set rsTemp=Nothing
			Else
				strSQL2=strSQL2 & CLng(Session("BanManProSiteID"))  
			End If		
			strSQL2=strSQL2 & ",'" & FixBlank(Request.Form("AdCodeNetscape")) & "')" 
			connBanManPro.Execute strSQL2
		    End If
			%>
			<p align="center"><font face="Arial" size="5">Successfully added new Banner.</font></p>
			<%  	 
		Case "UpdateAdvanced"
		    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then
			If Request.Form("Netscape4")="-1" Then
				lngNetscape4=-1
			Else
				lngNetscape4=0
			End If
 			strSQL2="UPDATE Banners SET "
			strSQL2=strSQL2 & "AdDescription='" &  FixBlank(Request.Form("AdDescription")) & "',"  
			strSQL2=strSQL2 & "AdvertiserID=" &  FixBlank(Request.Form("AdvertiserID")) & ","
			strSQL2=strSQL2 & "AdCode='" &  FixBlank(Request.Form("AdCode")) & "', "
			strSQL2=strSQL2 & "AdImageURL='" &  FixBlank(Request.Form("ImageSource")) & "', "
			strSQL2=strSQL2 & "AdTargetURL='" &  FixBlank(Request.Form("TargetURL")) & "', "
			strSQL2=strSQL2 & "AdNewWindow=" &  lngNetscape4 & ","
			strSQL2=strSQL2 & "AdWidth=" &  FixBlank(Request.Form("AdWidth"))  & ","
			strSQL2=strSQL2 & "AdHeight=" &  FixBlank(Request.Form("AdHeight"))   & ","
			If Request.Form("RunOfNetwork")="ON" Then
				'Check if Advertiser is global
				strTemp="Select AdvertiserID From Advertisers Where AdvertiserID=" & Clng(Request.Form("AdvertiserID")) & " And UserID=0"
				set rsTemp=connBanManPro.Execute(strTEmp)
				If Not rsTemp.EOF Then
					strSQL2=strSQL2 & "UserID=0,"
				Else
					strSQL2=strSQL2 & "UserID=" & CLng(Session("BanManProSiteID"))  & ","
				End If
				Set rsTemp=Nothing
			Else
				strSQL2=strSQL2 & "UserID=" & CLng(Session("BanManProSiteID"))  & ","
			End If
			strSQL2=strSQL2 & "AdCodeNetscape='" &  FixBlank(Request.Form("AdCodeNetscape")) & "'"
			strSQL2=strSQL2 & " WHERE Banners.[BannerID]=" & strBannerID
			connBanManPro.Execute strSQL2
		    End If
			%>
			<p align="center"><font face="Arial" size="5">Banner Updated.</font></p>
			<%  	 
		Case "AddNew" 
			'obtain list of advertisers
			strSQL="SELECT * FROM Advertisers Where (UserID=" & CLng(Session("BanManProSiteID")) & " Or UserID=0) ORDER BY Advertisers.[CompanyName] ASC"
			Set rsa=connBanManPro.Execute(strSQL)
			If NOT rsa.EOF Then
				%>
				<!--#include file="addanewbanner.asp"-->
				<%
			Else %>
				<p align="center"><font face="Arial" size="5">You must first add atleast one advertiser.</font></p>
			<% End If
			Set rsa=Nothing
		Case "Edit"
			'edit record
			If Trim(strBannerID) <> "" Then
				strSQL2="SELECT Banners.UserID,Banners.AdDescription,Banners.AdWidth,Banners.AdHeight,Banners.BannerID,Banners.AdFragment,Banners.AdImageURL,Banners.AdNewWindow,Banners.AdTargetURL,Advertisers.CompanyName,Banners.AdAltText,Banners.AdAlign,Banners.AdBorder,Banners.AdBorder,Banners.AdTextUnderneath,Banners.AdTextLink,Banners.AdTextLinkText,Advertisers.AdvertiserID,Banners.AdCodeNetscape,Banners.AdCode FROM Advertisers RIGHT JOIN Banners ON Advertisers.AdvertiserID = Banners.AdvertiserID WHERE (((Banners.BannerID)=" & strBannerID & "))"
				Set rsb=connBanManPro.Execute(strSQL2)
				'obtain list of advertisers
				strSQL="SELECT * FROM Advertisers Where (UserID=" & CLng(Session("BanManProSiteID")) & " OR UserID=0) ORDER BY Advertisers.[CompanyName] ASC"
				Set rsa=connBanManPro.Execute(strSQL)
				If Not rsb.EOF Then
 					%>
					<!--#include file="addanewBanner.asp"-->
					<%
				End If
				Set rsb=Nothing
			End If	
		Case "Insert"
			If Request.Form("AdTextLink") <> "-1" Then
				'error checks *************************
				If Trim(Request.Form("AdTargetURL"))="" Then
					Response.Write "Invalid Target URL"
					Response.End
				End If
				If Trim(Request.Form("AdImageURL"))="" Then
					Response.Write "Invalid Image URL"
					Response.End
				End If			
			Else
				If Trim(Request.Form("AdTextLinkText"))="" Then
					Response.Write "Invalid Link Text"
					Response.End
				End If		
			End If
			'end error checks *********************
		    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then
 			strSQL2="INSERT INTO Banners ("
			strSQL2=strSQL2 & "AdvertiserID,"
			strSQL2=strSQL2 & "AdDescription,"
			strSQL2=strSQL2 & "AdTargetURL,"
			strSQL2=strSQL2 & "AdAltText,"
			strSQL2=strSQL2 & "AdImageURL,"
			strSQL2=strSQL2 & "AdBorder,"
			strSQL2=strSQL2 & "AdWidth,"
			strSQL2=strSQL2 & "AdHeight,"
			strSQL2=strSQL2 & "AdAlign,"
			strSQL2=strSQL2 & "AdNewWindow,"
			strSQL2=strSQL2 & "AdTextUnderneath,"
			strSQL2=strSQL2 & "AdTextLink,"
			strSQL2=strSQL2 & "AdTextLinkText,"
			strSQL2=strSQL2 & "UserID,"
			strSQL2=strSQL2 & "AdCode) VALUES ("
			strSQL2=strSQL2 & FixBlank(Request.Form("AdvertiserID")) & ",'" 
			strSQL2=strSQL2 & FixBlank(Request.Form("AdDescription")) & "','" 
			strSQL2=strSQL2 & FixBlank(Request.Form("AdTargetURL")) & "','" 
			strSQL2=strSQL2 & FixBlank(Request.Form("AdAltText")) & "','" 
			strSQL2=strSQL2 & FixBlank(Request.Form("AdImageURL")) & "'," 
			strSQL2=strSQL2 & FixBlank(Request.Form("AdBorder")) & "," 
			strSQL2=strSQL2 & FixZero(Request.Form("AdWidth")) & "," 
			strSQL2=strSQL2 & FixZero(Request.Form("AdHeight")) & ",'" 
			strSQL2=strSQL2 & FixBlank(Request.Form("AdAlign")) & "'," 
			strSQL2=strSQL2 & FixBlank(Request.Form("AdNewWindow")) & ",'" 
			strSQL2=strSQL2 & FixBlank(Request.Form("AdTextUnderneath")) & "',"
			strSQL2=strSQL2 & SetTrueFalse(Request.Form("AdTextLink")) & ",'" 
			strSQL2=strSQL2 & FixBlank(Request.Form("AdTextLinkText")) & "'," 
			If Request.Form("RunOfNetwork")="ON" Then
				'Check if Advertiser is global
				strTemp="Select AdvertiserID From Advertisers Where AdvertiserID=" & Clng(Request.Form("AdvertiserID")) & " And UserID=0"
				set rsTemp=connBanManPro.Execute(strTEmp)
				If Not rsTemp.EOF Then
					strSQL2=strSQL2 & Clng(0) 
				Else
					strSQL2=strSQL2 & CLng(Session("BanManProSiteID")) 
				End If
				Set rsTemp=Nothing
			Else
				strSQL2=strSQL2 & CLng(Session("BanManProSiteID"))  
			End If		
			strSQL2=strSQL2 & ",'" & FixBlank(CreateAdCode(Request.Form("AdTargetURL"),Request.Form("AdImageURL"),Request.Form("AdWidth"),Request.Form("AdHeight"),Request.Form("AdAltText"),Request.Form("AdAlign"),Request.Form("AdBorder"),Request.Form("AdTextUnderneath"),Request.Form("AdTextLink"),Request.Form("AdTextLinkText"))) & "')" 
			connBanManPro.Execute strSQL2
		    End If
			%>
			<p align="center"><font face="Arial" size="5">Successfully added new Banner.</font></p>
			<%  	 
		Case "Update"
			If Request.Form("AdTextLink") <> "-1" Then
				'error checks *************************
				If Trim(Request.Form("AdTargetURL"))="" Then
					Response.Write "Invalid Target URL"
					Response.End
				End If
				If Trim(Request.Form("AdImageURL"))="" Then
					Response.Write "Invalid Image URL"
					Response.End
				End If			
			Else
				If Trim(Request.Form("AdTextLinkText"))="" Then
					Response.Write "Invalid Link Text"
					Response.End
				End If		
			End If
			'end error checks *********************
		    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then
 			strSQL2="UPDATE Banners SET "
			strSQL2=strSQL2 & "AdDescription='" &  FixBlank(Request.Form("AdDescription")) & "',"  
			strSQL2=strSQL2 & "AdvertiserID=" &  FixBlank(Request.Form("AdvertiserID")) & ","
			strSQL2=strSQL2 & "AdTargetURL='" &  FixBlank(Request.Form("AdTargetURL")) & "',"
			strSQL2=strSQL2 & "AdAltText='" &  FixBlank(Request.Form("AdAltText")) & "'," 
			strSQL2=strSQL2 & "AdImageURL='" &  FixBlank(Request.Form("AdImageURL")) & "'," 
			strSQL2=strSQL2 & "AdBorder=" &  FixBlank(Request.Form("AdBorder")) & "," 
			strSQL2=strSQL2 & "AdWidth=" &  FixZero(Request.Form("AdWidth")) & "," 
			strSQL2=strSQL2 & "AdHeight=" &  FixZero(Request.Form("AdHeight")) & "," 
			strSQL2=strSQL2 & "AdAlign='" &  FixBlank(Request.Form("AdAlign")) & "'," 
			strSQL2=strSQL2 & "AdNewWindow=" &  FixBlank(Request.Form("AdNewWindow")) & "," 
			strSQL2=strSQL2 & "AdTextUnderneath='" &  FixBlank(Request.Form("AdTextUnderneath")) & "',"
			strSQL2=strSQL2 & "AdTextLink=" &  SetTrueFalse(Request.Form("AdTextLink")) & ","
			strSQL2=strSQL2 & "AdTextLinkText='" &  FixBlank(Request.Form("AdTextLinkText")) & "',"
			If Request.Form("RunOfNetwork")="ON" Then
				'Check if Advertiser is global
				strTemp="Select AdvertiserID From Advertisers Where AdvertiserID=" & Clng(Request.Form("AdvertiserID")) & " And UserID=0"
				set rsTemp=connBanManPro.Execute(strTEmp)
				If Not rsTemp.EOF Then
					strSQL2=strSQL2 & "UserID=0,"
				Else
					strSQL2=strSQL2 & "UserID=" & CLng(Session("BanManProSiteID"))  & ","
				End If
				Set rsTemp=Nothing
			Else
				strSQL2=strSQL2 & "UserID=" & CLng(Session("BanManProSiteID"))  & ","
			End If
			strSQL2=strSQL2 & "AdCode='" &  FixBlank(CreateAdCode(Request.Form("AdTargetURL"),Request.Form("AdImageURL"),Request.Form("AdWidth"),Request.Form("AdHeight"),Request.Form("AdAltText"),Request.Form("AdAlign"),Request.Form("AdBorder"),Request.Form("AdTextUnderneath"),Request.Form("AdTextLink"),Request.Form("AdTextLinkText"))) & "' "
			strSQL2=strSQL2 & "WHERE Banners.[BannerID]=" & strBannerID
			connBanManPro.Execute strSQL2
		    End If
			%>
			<p align="center"><font face="Arial" size="5">Record Updated.</font></p>
			<%  	 
		Case "Delete"
			'delete entry
			If Trim(strBannerID) <> "" Then
			    If UCase(Session("UserName")) <> "ADMIN" And UCase(Session("Password"))<> "ADMIN" Then
				strSQL2="DELETE FROM Banners WHERE Banners.[BannerID]=" & strBannerID 
				connBanManPro.Execute strSQL2
			    End If
				%>
				<p align="center"><font face="Arial" size="5">Record Deleted.</font></p>
				<%  	
			Else 	%>
				<p align="center"><font face="Arial" size="5">Nothing to Delete.</font></p>
				<%
			End If		
		Case "ViewAll", ""
			If strTask="" Then
				strSQL2="SELECT Top 10 Banners.AdDescription,Banners.BannerID,Banners.AdFragment,Advertisers.CompanyName,Banners.AdCode "
				strSQL2=strSQL2 & "FROM Advertisers RIGHT JOIN Banners ON Advertisers.AdvertiserID = Banners.AdvertiserID Where (Banners.UserID=" & CLng(Session("BanManProSiteID")) & " Or Banners.UserID=0)"
			Else
				strSQL2="SELECT Banners.AdDescription,Banners.BannerID,Banners.AdFragment,Advertisers.CompanyName,Banners.AdCode "
				strSQL2=strSQL2 & "FROM Advertisers RIGHT JOIN Banners ON Advertisers.AdvertiserID = Banners.AdvertiserID Where (Banners.UserID=" & CLng(Session("BanManProSiteID")) & " Or Banners.UserID=0)"
				If Request("AdvertiserID") <> "" Then
					strSQL2=strSQL2 & " AND Advertisers.AdvertiserID=" & CLng(Request("AdvertiserID"))
				End If
				
			End If
			Set rsb=connBanManPro.Execute(strSQL2)
			If Not rsb.EOF Then
				strMessage="Listing of all Banners in Database."	
				'call include file and create table of all data
				%>
				<!--#include file="showallBanners.asp"-->
				<%
			End If
			Set rsb=Nothing
		Case "Details"
			'show Banner entry
			If Trim(strBannerID) <> "" Then
				strSQL2="SELECT Banners.*, Advertisers.CompanyName  FROM Advertisers RIGHT JOIN Banners ON Advertisers.AdvertiserID = Banners.AdvertiserID  WHERE (((Banners.BannerID)=" & strBannerID & ")) AND UserID=" & CLng(Session("BanManProSiteID"))
				Set rsb=connBanManPro.Execute(strSQL2)
				If Not rsb.EOF Then
					strMessage="Banner information for: " & rsb("AdDescription")  	
					'call include file to create table of data
					%>
					<!--#include file="showBannerdetails.asp"-->
					<%
				End If
				Set rsb=Nothing

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

<% '''''''''''''''''''Create AdCode''''''''''''''''''''''''''''''''''''''''''''''''''''
Function CreateAdCode(strTargetURL,strImageURL,strWidth,strHeight,strAltText,strAlign,strBorder,strTextUnderneath,blnAdTextLink,strAdTextLinkText)
	If blnAdTextLink="-1" Then
		strAdCode="<a href=" & Chr(34) & strTargetURL & Chr(34) & ">" & strAdTextLinkText & "</a>"
		strAdCode=strAdCode & "<a href=" & Chr(34) & strTargetURL & Chr(34) & "><img src=" & Chr(34) & "blank.gif" & Chr(34) & " width=" & Chr(34) & "1" & Chr(34) & " height=" & Chr(34) & "1" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></a>"
	Else
		strAdCode="<a href=" & Chr(34) & strTargetURL & Chr(34) & "><img src=" & Chr(34) & strImageURL & Chr(34)
		strAdCode=strAdCode & "  width=" & Chr(34) & strWidth & Chr(34) & " height=" & Chr(34) & strHeight & Chr(34) & " alt=" & Chr(34) & Trim(strAltText) & Chr(34) & " align=" & Chr(34) &  strAlign & Chr(34) & " border=" & Chr(34) & strBorder & Chr(34) & "></a><br>"
		strAdCode=strAdCode & "  <a href="  & Chr(34) & strTargetURL & Chr(34) &  ">" & strTextUnderneath & "</a>"
	End If
	CreateAdCode=strAdCode
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
<% '''''''''''''''''''''change blank field to 0 '''''''''''''''''''''''''''''''''''''''''''''''''
Function FixZero(strData)
	If Trim(strData)="" Then
		FixZero=0
	Else
		FixZero=strData
	End If
End Function
%>