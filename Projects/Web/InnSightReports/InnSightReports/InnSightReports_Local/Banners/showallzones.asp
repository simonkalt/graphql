<html>

<head>
<title></title>
</head>

<body>
  <div align="center">
    <center>
    <table border="0" cellpadding="0" cellspacing="0" width="590">
      <tr>
        <td><a href="help/zones.htm" target="_new"><img border="0" src="images/ListingofAllZones.gif" WIDTH="590" HEIGHT="30"></a></td>
      </tr>
    </table>
    </center>
  </div>
  <% Do While Not rsz.EOF %>

<div align="center"><center>

<table border="1" cellpadding="4" cellspacing="0" width="590" bordercolor="#003063" background="images/tableback.gif">
  <tr>
    <td width="46"><p align="center"><a href="zones.asp?Task=Edit&amp;ZoneID=<%=rsz("ZoneID")%>"><img src="images/Editsmall.gif" alt="Edit Zone" border="0" WIDTH="38" HEIGHT="18"></a><br>
    <a href="zones.asp?Task=ViewCode&amp;ZoneID=<%=rsz("ZoneID")%>"><img src="images/code.gif" alt="View Code for this Zone" border="0" WIDTH="38" HEIGHT="18"></a></td>
    <td><div align="left"><table border="0" cellpadding="0" cellspacing="0" width="100%">
      <tr>
        <td align="center"><strong><font face="Arial" size="3"><%=rsz("ZoneDescription")%></font></strong></td>
      </tr>
      <tr>
        <td align="center"><%If Instr(UCase(rsz("ZonePageURL")),"HTTP")>0 Then%><a href="<%=rsz("ZonePageURL")%>"><%=rsz("ZonePageURL")%></a><%End If%></td>
      </tr>
      <tr>
        <td align="center"><%If IsNull(rsz("ZoneAverage")) Then%>0<%Else%><%=rsz("ZoneAverage")%><%End If%> Impressions per day based on <%=" " & Application("ZoneAverageDays") & " " %>days.</a></td>
      </tr>
    </table>
    </div></td>
    <td width="44"><a href="zones.asp?Task=Delete&amp;Confirm=True&amp;ZoneID=<%=rsz("ZoneID")%>"><img src="images/delsmall.gif" alt="Delete Zone" border="0" WIDTH="38" HEIGHT="18"></a></td>
  </tr>
</table>
</center></div><% rsz.MoveNext
Loop %>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="590">
    <tr>
      <td><img border="0" src="images/bottomblue.gif" WIDTH="590" HEIGHT="30"></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>

<% Function NullToZero(strField)
	If IsNull(strField) Then
		NullToZero=0
	Else
		NullToZero=Csng(strField)
	End If
End Function %>