<html>

<head>
<title></title>
</head>

<body>
<div align="center">
<%If Request("Task")="" Then %>
<font face="Arial" size="3">Listing of First 10 Banners</font><p>
<% End If %>
<table border="0" cellpadding="0" cellspacing="0" width="590" bordercolor="#000080" style="border: medium none" background="images/tableback.gif">
  <tr>
    <td align="center" bgcolor="#808080"><a href="help/banners.htm" target="_new"><img border="0" src="images/ListingofAllbanners.gif" WIDTH="590" HEIGHT="30"></a></td>
  </tr>
</table>
<% 
intCnt=0
Do While rsb.EOF <> True 
intCnt=intCnt+1
%>

<div align="center"><center>

<table border="0" cellpadding="4" cellspacing="0" width="593" bordercolor="#000080" style="border-style: none; border-width: medium" background="images/tableback.gif">
  <tr>
    <td colspan="3" align="center" bgcolor="#808080" width="583"><font face="Arial" size="3">
<% If intCnt=1 Or Request("AdvertiserID")="" Then %><b>Advertiser:</b>
      <%=rsb("CompanyName")%> <% End If %>
      <br><b>Banner:</b> <%=rsb("AdDescription")%></font></td>
  </tr>
  <tr>
    <td width="39"><%If rsb("AdFragment")<>True Then%><a href="banners.asp?Task=Edit&amp;BannerID=<%=rsb("BannerID")%>"><img src="images/Editsmall.gif" alt="Edit Banner" border="0" WIDTH="38" HEIGHT="18"></a><br><%End If%>
<%If rsb("AdFragment")=True Then%>    <a href="banners.asp?Task=ViewCode&amp;BannerID=<%=rsb("BannerID")%>"><img src="images/code.gif" alt="View the Code for this Banner Ad" border="0" WIDTH="38" HEIGHT="18"></a><%End If%></td>
    <td width="478"><p align="center"><%=rsb("AdCode")%></td>
    <td width="46"><a href="banners.asp?Task=Delete&amp;Confirm=True&amp;BannerID=<%=rsb("BannerID")%>"><img src="images/delsmall.gif" alt="Delete Banner" border="0" WIDTH="38" HEIGHT="18"></a></td>
  </tr>
</table>
</center></div></div><% rsb.MoveNext
Loop %>

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
