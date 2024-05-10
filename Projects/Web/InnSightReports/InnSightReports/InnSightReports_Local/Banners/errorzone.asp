<% 
strData=Chr(60) & "% Dim strZoneID"  & vbCRLF & "Dim strTask" &  vbCRLF 
strData=strData & "strZoneID=" & strZoneID & vbCRLF 
strData=strData & "lngBMPSiteID=" & Clng(varCampaigns(intCnt)) & vbCRLF 
strData=strData & "strTask=" & Chr(34) & "Get" & Chr(34) & vbCRLF & "%" & Chr(62) & vbCRLF  & Chr(60) 
strData=strData & "!--#include virtual=" & Chr(34) & getFilePath() & "banman.asp" & Chr(34) & "-->" %>


<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>An Error Occurred Writing Zone File</title>
</head>

<body>
<div align="center"><center>

<table border="0" cellpadding="5" cellspacing="0" width="510">
  <tr>
    <td width="510"><font face="Arial" size="2"><font color="#FF0000">An Error Occurred Writing
    The Zone File of:</font><br>
      <%=strPath & "zones\banmanzone" & strZoneID & ".asp"%> <br>
    <br>
    Typically the cause of the problem is insufficient privelages on the ZONES
    directory.&nbsp; Ban Man Pro uses the FileSystemObject to write files to the ZONES
    directory.&nbsp; Visit the Ban Man Pro <a
    href="http://www.banmanpro.com/support.asp">support page</a> for information on how to fix
    this problem. &nbsp; Another potential cause is you have not properly set the <strong>Server
    Path</strong> in your preferences.&nbsp; Lastly, if you are using the HTML
      mode then you have incorrectly selected the mode.</font></td>
  </tr>
  <tr>
    <td width="510"><strong><font face="Arial" size="2">As a work-around, you can do the following:</font></strong></td>
  </tr>
  <tr>
    <td width="510"><ol>
      <li><font face="Arial" size="2">Create A File in the Zones Directory Called: <%="banmanzone" & strZoneID & ".asp"%></font></li>
      <li><font face="Arial" size="2">Copy and paste the following code into that file.</font></li>
    </ol>
    <form method="POST" action>
      <p><font face="Arial" size="2"><textarea rows="5" name="S1" cols="50"><%=strData%></textarea></font></p>
    </form>
    </td>
  </tr>
</table>
</center></div>
</body>
</html>
