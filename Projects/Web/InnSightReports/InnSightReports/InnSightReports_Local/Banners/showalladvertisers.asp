<html>

<head>
<title></title>
</head>

<body>
<div align="center">
  <center>

<table border="0" cellpadding="0" cellspacing="0" width="590" bordercolor="#000080" background="images/tableback.gif">
  <tr>
    <td align="left"><a href="help/advertisers.htm" target="_new"><img border="0" src="images/ListingofAllAdvertisers.gif" alt="Click for more help on advertisers." WIDTH="590" HEIGHT="30"></a></td>
  </tr>
</table>
  </center>
</div>
<div align="center"><center>

<table border="2" cellpadding="4" cellspacing="0" width="590" bordercolor="#003063" background="images/tableback.gif">
  <tr>
    <td align="left" width="38">&nbsp;</td>
    <td align="left" width="212"><font face="Arial" size="2"><strong>Company Name </strong></font></td>
    <td align="left" width="127"><font face="Arial" size="2"><strong>Description</strong></font></td>
    <td align="left" width="123"><font face="Arial" size="2"><strong>Contact</strong></font></td>
    <td align="left" width="38">&nbsp;</td>
  </tr>
<% Do While rss.EOF <> True %>
  <tr>
    <td align="left" width="38"><a href="advertisers.asp?Task=Edit&amp;AdvertiserID=<%=rss("AdvertiserID")%>"><img src="images/Editsmall.gif" alt="Edit Advertiser" border="0" WIDTH="38" HEIGHT="18"></a></td>
    <td align="left" width="212"><font face="Arial" size="2"><%If Trim(rss("CompanyWebSite"))<>"" Then %><a href="<%=rss("CompanyWebSite")%>"><%=rss("CompanyName")%></a><%Else%><%=rss("CompanyName")%><%End If%></font></td>
    <td align="left" width="127"><font face="Arial" size="2"><%=rss("AdvertiserDesc")%></font></td>
    <td align="left" width="123"><font face="Arial" size="2"><a href="mailto:<%=rss("Email")%>"><%=rss("Contact")%></a></font></td>
    <td align="left" width="38"><a href="advertisers.asp?Task=Delete&amp;AdvertiserID=<%=rss("AdvertiserID")%>&amp;Confirm=True"><img src="images/delsmall.gif" alt="Delete Advertiser" border="0" WIDTH="38" HEIGHT="18"></a></td>
  </tr>
<% rss.MoveNext
Loop %>
</table>
</center></div>
<div align="center">
  <center>
<table border="0" cellpadding="0" cellspacing="0" width="590" bordercolor="#000080" style="border: medium none" background="images/tableback.gif">
  <tr>
    <td align="center" bgcolor="#808080"><img border="0" src="images/bottomblue.gif" WIDTH="590" HEIGHT="30"></td>
  </tr>
</table>
  </center>
</div>
</body>
</html>
