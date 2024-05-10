      <div align="center">
        <table border="0" cellpadding="0" cellspacing="0" width="590">
          <tr>
            <td><a href="help/sites.htm" target="_new"><img border="0" src="images/banmanprosites.gif" WIDTH="590" HEIGHT="30"></a></td>
          </tr>
        </table>
      </div>
      <div align="center">
        <center>
        <table border="2" cellpadding="0" cellspacing="0" width="590" bordercolor="#003063">
          <tr>
            <td width="38"></td>
            <td width="82" align="center"><font face="Arial" size="2"><b>Site ID</b></font></td>
            <td width="420" align="center"><font face="Arial" size="2"><b>Site Name</b></font></td>
            <td width="38"></td>
          </tr>
<% Do While Not rss.EOF %>
          <tr>
            <td width="38"><a href="sites.asp?Task=Edit&amp;SiteID=<%=rss("SiteID")%>"><img border="0" src="images/Editsmall.gif" WIDTH="38" HEIGHT="18"></a></td>
            <td width="82" align="center"><font face="Arial" size="2"><%=rss("SiteID")%></font></td>
            <td width="420" align="center"><font face="Arial" size="2"><a href="<%=rss("SiteURL")%>"><%=rss("SiteName")%></a></font></td>
            <td width="38"><a href="sites.asp?Task=Delete&amp;SiteID=<%=rss("SiteID")%>&amp;Confirm=True"><img border="0" src="images/delsmall.gif" WIDTH="38" HEIGHT="18"></a></td>
          </tr>
<% rss.MoveNext
Loop %>
        </table>
        </center>
      </div>
            <div align="center">
        <table border="0" cellpadding="0" cellspacing="0" width="590">
          <tr>
            <td><img border="0" src="images/bottomblue.gif" WIDTH="590" HEIGHT="30"></td>
          </tr>
        </table>
      </div>