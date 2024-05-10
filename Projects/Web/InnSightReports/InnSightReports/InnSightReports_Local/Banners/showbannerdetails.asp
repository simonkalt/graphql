<html>

<head>
<title></title>
</head>

<body>
<div align="center"><center>

<table border="2" cellpadding="5" cellspacing="0" width="500">
  <tr>
    <td width="180" align="right"><font face="Arial" size="2"><strong>Banner Ad Description:</strong></font></td>
    <td width="320"><font color="#0000A0" face="Arial" size="2"><%=rsb("AdDescription")%></font></td>
  </tr>
  <tr>
    <td width="180" align="right"><font face="Arial" size="2"><strong>Advertiser:</strong></font></td>
    <td width="320"><font color="#0000A0" face="Arial" size="2"><a href="<%=rsb("CompanyName")%>"><%=rsb("CompanyName")%></a></font></td>
  </tr>
  <tr>
    <td width="180" align="right"><font face="Arial" size="2"><strong>Target URL:</strong></font></td>
    <td width="320"><font color="#0000A0" face="Arial" size="2"><a
    href="<%=rsb("AdTargetURL")%>"><%=rsb("AdTargetURL")%></a></font></td>
  </tr>
  <tr>
    <td width="180" align="right"><font face="Arial" size="2"><strong>Image URL:</strong></font></td>
    <td width="320"><font color="#0000A0" face="Arial" size="2"><%=rsb("AdImageURL")%></font></td>
  </tr>
  <tr>
    <td width="180" align="right"><font face="Arial" size="2"><strong>Image Text:</strong></font></td>
    <td width="320"><font color="#0000A0" face="Arial" size="2"><%=rsb("AdAltText")%></font></td>
  </tr>
  <tr>
    <td width="180" align="right"><font face="Arial" size="2"><strong>Width, Height, Border:</strong></font></td>
    <td width="320"><font color="#0000A0" face="Arial" size="2"><%=rsb("AdWidth")%>, <%=rsb("AdHeight")%>, <%=rsb("AdBorder")%></font></td>
  </tr>
  <tr>
    <td width="180" align="right"><font face="Arial" size="2"><strong>Alignment:</strong></font></td>
    <td width="320"><font color="#0000A0" face="Arial" size="2"><%=rsb("AdAlign")%></font></td>
  </tr>
  <tr>
    <td width="180" align="right"><font face="Arial" size="2"><strong>Launch New Window:</strong></font></td>
    <td width="320"><font color="#0000A0" face="Arial" size="2"><%If rsb("AdNewWindow")=-1 Then%>Yes<%Else%>No<%End If%></font></td>
  </tr>
  <tr>
    <td width="180" align="right"><font face="Arial" size="2"><strong>Optional Text:</strong></font></td>
    <td width="320"><font color="#0000A0" face="Arial" size="2"><%=rsb("AdTextUnderneath")%></font></td>
  </tr>
</table>
</center></div>
</body>
</html>
