<html>

<head>
<title></title>
</head>

<body>
  <div align="center">
  <%If Request("Task")="" Then %>
<font face="Arial" size="3">Listing of First 10 VALID Campaigns</font><p>
<% End If %>
    <center>
    <table border="0" cellpadding="0" cellspacing="0" width="590">
      <tr>
        <td><a href="help/campaigns.htm" target="_new"><img border="0" src="images/ListingofAllCampaigns.gif" WIDTH="590" HEIGHT="30"></a></td>
      </tr>
    </table>
    </center>
  </div>
<div align="center"><% Do While Not rsc.EOF %>

<div align="center"><center>

<table border="1" cellpadding="2" cellspacing="0" width="590" bordercolor="#003063" background="images/tableback.gif">
  <tr>
    <td width="61" height="50" rowspan="3"><p align="center"><a href="campaigns.asp?Task=Edit&amp;CampaignID=<%=rsc("CampaignID")%>"><img src="images/Editsmall.gif" alt="Edit Campaign" border="0" WIDTH="38" HEIGHT="18"></a>
    <%If rsc("CampaignDistribution")="Text" Then %>
    <br><a href="campaigns.asp?Task=Link&amp;CampaignID=<%=rsc("CampaignID")%>"><img src="images/link.gif" alt="View Static Code" border="0" WIDTH="38" HEIGHT="18"></a>
    <%End If%>
    </td>
    <td width="286" height="17" colspan="2"><%If rsc("CampaignSiteDefault")=True Then%><font face="Arial"><font color="#FF0000">**Site Default Campaign</font><font size="2"><br><%End If%><font face="Arial" size="2"><strong>Name</strong>:
    <a href="campaigns.asp?Task=Edit&amp;CampaignID=<%=rsc("CampaignID")%>"><%=rsc("CampaignName")%></a></font></td>
    <td width="277" height="17" colspan="2"><font face="Arial" size="2"><strong>Company</strong>:
<%=rsc("CompanyName")%>    </font></td>
    <td width="57" height="50" rowspan="3"><font face="Arial" size="2"><p align="center"></font><a href="campaigns.asp?Task=Delete&amp;Confirm=True&amp;CampaignID=<%=rsc("CampaignID")%>"><img src="images/delsmall.gif" alt="Delete Campaign" border="0" WIDTH="38" HEIGHT="18"></a></td>
  </tr>
<% If rsc("CampaignSiteDefault") <> True Then %>
  <tr>
    <td width="109" height="17" align="left"><font face="Arial" size="2"><strong>Start Date</strong>:<br>
<%=rsc("CampaignStartDate")%>    </font></td>
    <td width="109" height="17" align="left"><font face="Arial" size="2"><strong>End Date</strong>:<br>
<%=FormatDateTime(rsc("CampaignEndDate"),vbShortDate)%>    </font></td>
    <td width="109" height="17" align="left"><font face="Arial" size="2"><strong>Start Time</strong>:<br>
<%=FormatDateTime(rsc("CampaignDailyStart"),vbLongTime)%>    </font></td>
    <td width="110" height="17" align="left"><font face="Arial" size="2"><strong>End Time</strong>:<br>
<%=FormatDateTime(rsc("CampaignDailyEnd"),vbLongTime)%>    </font></td>
  </tr>
  <tr>
    <td width="109" height="16" align="left"><font face="Arial" size="2"><strong>Type</strong>:<br>
<%=rsc("CampaignType")%>    </font></td>
    <td width="109" height="16" align="left"><font face="Arial" size="2"><strong>Quantity</strong>:<br>
<%=rsc("CampaignQuantitySold")%>    </font></td>
    <td width="109" height="16" align="left"><font face="Arial" size="2"><strong>Cost</strong>:<br>
<%=rsc("CampaignCost")%>    </font></td>
    <td width="110" height="16" align="left"><font face="Arial" size="2"><strong>Distribution</strong>:<br>
<%=rsc("CampaignDistribution")%>    </font></td>
  </tr>
<%End If %>
</table>
<% rsc.MoveNext
Loop %>
</div>

  <div align="center">
    <center>
    <table border="0" cellpadding="0" cellspacing="0" width="590">
      <tr>
        <td><img border="0" src="images/bottomblue.gif" WIDTH="590" HEIGHT="30"></td>
      </tr>
    </table>
    </center>
  </div>
  </center></div>
</body>
</html>
