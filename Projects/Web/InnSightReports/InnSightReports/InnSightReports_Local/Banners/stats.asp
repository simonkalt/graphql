<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title></title>
</head>

<body>
<div align="center"><center>

<table border="1" cellpadding="3" cellspacing="0" background="images/tableback.gif" width="604" bordercolor="#000000">
  <tr>
    <td width="175" bgcolor="#7A74FA"><strong><font face="Arial" size="2">Campaign</font></strong></td>
    <td width="92" align="center" bgcolor="#7A74FA"><strong><font face="Arial" size="2">Start
    Date</font></strong></td>
    <td width="89" align="center" bgcolor="#7A74FA"><strong><font face="Arial" size="2">End
    Date</font></strong></td>
    <td width="75" align="center" bgcolor="#7A74FA"><strong><font face="Arial" size="2">Clicks</font></strong></td>
    <td width="37" align="center" bgcolor="#7A74FA"><strong><font face="Arial" size="2">Impressions</font></strong></td>
    <td width="64" align="center" bgcolor="#7A74FA"><strong><font face="Arial" size="2">Click
    Rate</font></strong></td>
  </tr>
<% 
intCntColor=0
Do While Not rs.EOF 
If intCntColor=0 Then
	strColor="#B6B6B6" 
Else
	strColor="#AFABFC"
End If %>
  <tr>
    <td align="left" width="175" bgcolor="<%=strColor%>"><font face="Arial" size="2"><%If Session("AdvertiserID")<=0 Then%><a href="campaigns.asp?Task=Edit&CampaignID=<%=rs("CampaignID")%>"><%=rs("CampaignName")%></a><%Else%><%=rs("CampaignName")%><%End If%>
      </font>
</td>
    <td align="center" width="92" bgcolor="<%=strColor%>"><font face="Arial" size="2"><%=FormatDateTime(rs("CampaignStartDate"),vbShortDate)%>
      </font>
</td>
    <td align="center" width="89" bgcolor="<%=strColor%>"><font face="Arial" size="2"><%=FormatDateTime(rs("CampaignEndDate"),vbShortDate)%>
      </font>
</td>
    <td align="center" width="75" bgcolor="<%=strColor%>"><font face="Arial" size="2"><%=rs("SumOfClicks")%>
      </font>
</td>
    <td align="center" width="37" bgcolor="<%=strColor%>"><font face="Arial" size="2"><%=rs("SumOfImpressionCount")%>
      </font>
<% If Clng(rs("SumOfClicks"))> 0 AND Clng(rs("SumOfImpressionCount")) Then
	varPercent=FormatPercent(rs("SumOfClicks")/rs("SumOfImpressionCount"))
Else
	varPercent="0.00%"
End If %>
</td>
    <td align="center" width="64" bgcolor="<%=strColor%>"><font face="Arial" size="2"><%=varPercent%>
      </font>
</td>
  </tr>
<% rs.MoveNext
If intCntColor=0 Then
	intCntColor=1
Else
	intCntColor=0
End If
Loop %>
</table>
</center></div>
</body>
</html>
