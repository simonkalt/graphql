<%@ Language=VBScript %>

<!--#INCLUDE file="checkuser.asp"-->
<!-- #INCLUDE FILE="Data\adovbs.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<body bgcolor=#F9D568 topmargin="0" leftmargin="0" marginwidth = "0" marginheight = "0"  link="blue" vlink="blue" alink="blue">

<!--#include file = "Header.inc" ---> 
<%
	Set cnMain = Server.CreateObject("ADODB.Connection")
    Set rs = Server.CreateObject("ADODB.Recordset")
    Dim strSQL
    
	cnMain.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")
	
	strSQL = "sp_LoginConfirm '" & Session("login") & "', '" & session("password") & "'"
	
	rs.Open strSQL,cnMain,adOpenKeyset,adLockReadOnly
	
%>
	

<%
  Dim intRows
  Dim intCols
  
  'Load up array

  rs.MoveFirst
  
  intImageCount = (rs.RecordCount)
  
  intCols = 2
  intRows = intImageCount \ intCols
  
  If intImageCount mod intCols > 0 then
    intRows = intRows + 1
  End If

  
  'Set up array
  'Define data array
  ReDim aData(intRows,intCols)
  ReDim aText(intRows,intCols)

		If rs.EOF Then
			Response.Write "No companies found." & vbcrlf
			Response.End
		Else
			For intCntRows = 1 to intRows
				For intCntCols = 1 to intCols
					If Not rs.EOF Then
						strCell = "<font face=tahoma size=2><a href=SelectCompanyConfirm.asp?ID=" & rs.Fields("CompanyID") & "><IMG border=1 width=100 SRC=images/Hotel.gif><br>" & rs.Fields("CompanyName") & "</a></font>"
						aData(intCntRows,intCntCols) = strCell
						rs.MoveNext
					End If
				Next
			Next
		End If

  
%>

<%

			Response.Write "<br><TABLE width=80% cellpadding=5 border=3 bgcolor=#c0c0c0 align=center>" ' style='BORDER-RIGHT: medium none; BORDER-TOP: medium none; BACKGROUND: #999999; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none; BORDER-COLLAPSE: collapse; mso-border-alt: solid windowtext .5pt; mso-padding-alt: 0in 5.4pt 0in 5.4pt' cellSpacing=0 cellPadding=0 bgColor=#999999 width=750 border=1>"
			For intCntRows = 1 to intRows
				Response.Write "<TR>" 'style='HEIGHT: 57.1pt'>"
				For intCntCols = 1 to intCols
					Response.Write "<TD valign=middle align=center>" ' style='BORDER-RIGHT: windowtext 0.5pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: windowtext 0.5pt solid; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 0.5pt solid; WIDTH: 1.7in; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 0.5pt solid; HEIGHT: 57.1pt' width=163>"
					Response.Write aData(intCntRows,intCntCols) & "<BR>"
					Response.Write "</TD>"
				Next 
				Response.Write "</TR>"
			Next 
			Response.Write "</Table>"
%>

</p>

</BODY>
</HTML>
