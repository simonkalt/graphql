<%  	
	If Session("AdvertiserID")=0 Then
		strSQL="SELECT * FROM Campaigns Where UserID=" & CLng(Session("BanManProSiteID")) & " OR UserID=0 ORDER BY Campaigns.CampaignName ASC"
		Set rsCampaigns=connBanManPro.Execute(strSQL)    
	Else
		strSQL="SELECT * FROM Campaigns WHERE AdvertiserID=" & Session("AdvertiserID") &  " ORDER BY Campaigns.CampaignName ASC"
		Set rsCampaigns=connBanManPro.Execute(strSQL)  
	End If

		'retrieve reports that advertiser is allowed to view
		strSQL="Select * From BanManProReports"
		Set rsReportTypes=connBanManPro.Execute(strSQL)  
  
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Detailed Reports</title>
</head>

<body>

<form method="POST" action="createreport.asp">
  <div align="center">
    <center>
    <table border="0" cellpadding="0" cellspacing="0" width="498">
      <tr>
        <td><img border="0" src="images/CreateReporttop.gif" WIDTH="498" HEIGHT="30"></td>
      </tr>
    </table>
    </center>
  </div>
  <div align="center"><center><table border="0" cellpadding="5" cellspacing="0" width="498" background="images/tableback.gif">
    <tr>
      <td align="right"><font face="Arial"><strong>Campaign:</strong></font></td>
      <td><font face="Arial"><select name="Campaign" size="1">
        <%If Session("AdvertiserID")=0 Then %><option selected value="All Campaigns">All Campaigns</option><%End If%>
<% Do While NOT rsCampaigns.EOF %>        
	<option value="<%=rsCampaigns("CampaignID")%>"><%=rsCampaigns("CampaignName")%></option>
<%rsCampaigns.MoveNext
Loop
%>      </select></font></td>
    </tr>
    <tr>
      <td align="right"><font face="Arial"><strong>Report Type:</strong></font></td>
      <td><font face="Arial"><select name="ReportType" size="1">
<% If rsReportTypes("Reports_SummaryByDay")=True Or Session("AdvertiserID")=0 Then %>
        <option value="Summary By Day">Summary By Day</option>
<% End If %>
<% If rsReportTypes("Reports_SummaryByBanner")=True Or Session("AdvertiserID")=0 Then %>
        <option value="Summary By Banner">Summary By Banner</option>
<% End If %>
<% If rsReportTypes("Reports_SummaryByBannerByDay")=True Or Session("AdvertiserID")=0 Then %>
        <option value="Summary By Banner By Day">Summary By Banner By Day</option>
<% End If %>
<% If rsReportTypes("Reports_SummaryByZone")=True Or Session("AdvertiserID")=0 Then %>
        <option value="Summary By Zone">Summary By Zone</option>
<% End If %>
<% If rsReportTypes("Reports_SummaryByZoneByDay")=True Or Session("AdvertiserID")=0 Then %>
        <option value="Summary By Zone By Day">Summary By Zone By Day</option>
<% End If %>
<% If rsReportTypes("Reports_ClickDetail")=True Or Session("AdvertiserID")=0 Then %>
        <option value="Click Detail">Click Detail</option>
<% End If %>
<%If Session("AdvertiserID")=0 Then %>
        <option value="Executive">Executive</option>
	<option value="Billing">Billing Summary</option>
<option value="Expiration">Campaign Expiration</option>
<% If Application("BanManProMultiSite")=True Then %>
<option value="Cross Site Summary By Zone">All Sites Summary By Zone</option>
<option value="Cross Site Summary By Campaign">All Sites Summary By Campaign</option>
<% End If %>
<% End If %>
      </select></font></td>
    </tr>
    <tr>
      <td align="right"><font face="Arial"><strong>Format:</strong></font></td>
      <td><font face="Arial"><select name="ReportFormat" size="1">
          <option value="HTML">HTML</option>
          <option value="EXCEL">Excel</option>
      </select></font></td>
    </tr>
    <tr>
      <td align="right"><strong><font face="Arial">Start Date:</font></strong></td>
      <td><div align="left"><table border="0" cellpadding="0" cellspacing="0" width="307">
        <tr>
          <td width="29"><small><font face="Arial">Year</font></small></td>
          <td width="83"><font face="Arial" size="1"><select name="StartYear" size="1">
            <option <% If Year(Now)="1999" Then%>selected<% End If%> value="1999">1999</option>
            <option <% If Year(Now)="2000" Then%>selected<% End If%> value="2000">2000</option>
            <option <% If Year(Now)="2001" Then%>selected<% End If%> value="2001">2001</option>
            <option <% If Year(Now)="2002" Then%>selected<% End If%> value="2002">2002</option>
            <option <% If Year(Now)="2003" Then%>selected<% End If%> value="2003">2003</option>
            <option <% If Year(Now)="2004" Then%>selected<% End If%> value="2004">2004</option>
            <option <% If Year(Now)="2005" Then%>selected<% End If%> value="2005">2005</option>
            <option <% If Year(Now)="2006" Then%>selected<% End If%> value="2006">2006</option>
            <option <% If Year(Now)="2007" Then%>selected<% End If%> value="2007">2007</option>
            <option <% If Year(Now)="2008" Then%>selected<% End If%> value="2008">2008</option>
            <option <% If Year(Now)="2009" Then%>selected<% End If%> value="2009">2009</option>
            <option <% If Year(Now)="2010" Then%>selected<% End If%> value="2010">2010</option>
          </select></font></td>
          <td width="38"><small><font face="Arial">Month</font></small></td>
          <td width="78"><font face="Arial" size="1"><select name="StartMonth" size="1">
            <option <% If Month(Now)="1" Then%>selected<% End If%> value="01">Jan</option>
            <option <% If Month(Now)="2" Then%>selected<% End If%> value="02">Feb</option>
            <option <% If Month(Now)="3" Then%>selected<% End If%> value="03">Mar</option>
            <option <% If Month(Now)="4" Then%>selected<% End If%> value="04">Apr</option>
            <option <% If Month(Now)="5" Then%>selected<% End If%> value="05">May</option>
            <option <% If Month(Now)="6" Then%>selected<% End If%> value="06">Jun</option>
            <option <% If Month(Now)="7" Then%>selected<% End If%> value="07">Jul</option>
            <option <% If Month(Now)="8" Then%>selected<% End If%> value="08">Aug</option>
            <option <% If Month(Now)="9" Then%>selected<% End If%> value="09">Sep</option>
            <option <% If Month(Now)="10" Then%>selected<% End If%> value="10">Oct</option>
            <option <% If Month(Now)="11" Then%>selected<% End If%> value="11">Nov</option>
            <option <% If Month(Now)="12" Then%>selected<% End If%> value="12">Dec</option>
          </select></font></td>
          <td width="25"><small><font face="Arial">Day</font></small></td>
          <td width="54"><font face="Arial" size="1"><select name="StartDay" size="1">
            <option <% If Day(Now)="1" Then%>selected<% End If%> value="1">1</option>
            <option <% If Day(Now)="2" Then%>selected<% End If%> value="2">2</option>
            <option <% If Day(Now)="3" Then%>selected<% End If%> value="3">3</option>
            <option <% If Day(Now)="4" Then%>selected<% End If%> value="4">4</option>
            <option <% If Day(Now)="5" Then%>selected<% End If%> value="5">5</option>
            <option <% If Day(Now)="6" Then%>selected<% End If%> value="6">6</option>
            <option <% If Day(Now)="7" Then%>selected<% End If%> value="7">7</option>
            <option <% If Day(Now)="8" Then%>selected<% End If%> value="8">8</option>
            <option <% If Day(Now)="9" Then%>selected<% End If%> value="9">9</option>
            <option <% If Day(Now)="10" Then%>selected<% End If%> value="10">10</option>
            <option <% If Day(Now)="11" Then%>selected<% End If%> value="11">11</option>
            <option <% If Day(Now)="12" Then%>selected<% End If%> value="12">12</option>
            <option <% If Day(Now)="13" Then%>selected<% End If%> value="13">13</option>
            <option <% If Day(Now)="14" Then%>selected<% End If%> value="14">14</option>
            <option <% If Day(Now)="15" Then%>selected<% End If%> value="15">15</option>
            <option <% If Day(Now)="16" Then%>selected<% End If%> value="16">16</option>
            <option <% If Day(Now)="17" Then%>selected<% End If%> value="17">17</option>
            <option <% If Day(Now)="18" Then%>selected<% End If%> value="18">18</option>
            <option <% If Day(Now)="19" Then%>selected<% End If%> value="19">19</option>
            <option <% If Day(Now)="20" Then%>selected<% End If%> value="20">20</option>
            <option <% If Day(Now)="21" Then%>selected<% End If%> value="21">21</option>
            <option <% If Day(Now)="22" Then%>selected<% End If%> value="22">22</option>
            <option <% If Day(Now)="23" Then%>selected<% End If%> value="23">23</option>
            <option <% If Day(Now)="24" Then%>selected<% End If%> value="24">24</option>
            <option <% If Day(Now)="25" Then%>selected<% End If%> value="25">25</option>
            <option <% If Day(Now)="26" Then%>selected<% End If%> value="26">26</option>
            <option <% If Day(Now)="27" Then%>selected<% End If%> value="27">27</option>
            <option <% If Day(Now)="28" Then%>selected<% End If%> value="28">28</option>
            <option <% If Day(Now)="29" Then%>selected<% End If%> value="29">29</option>
            <option <% If Day(Now)="30" Then%>selected<% End If%> value="30">30</option>
            <option <% If Day(Now)="31" Then%>selected<% End If%> value="31">31</option>
          </select></font></td>
        </tr>
      </table>
      </div></td>
    </tr>
    <tr>
      <td align="right"><font face="Arial"><strong>End Date:</strong></font></td>
      <td><div align="left"><table border="0" cellpadding="0" cellspacing="0" width="307">
        <tr>
          <td width="29"><small><font face="Arial">Year</font></small></td>
          <td width="84"><font face="Arial" size="1"><select name="EndYear" size="1">
            <option <% If Year(Now)="1999" Then%>selected<% End If%> value="1999">1999</option>
            <option <% If Year(Now)="2000" Then%>selected<% End If%> value="2000">2000</option>
            <option <% If Year(Now)="2001" Then%>selected<% End If%> value="2001">2001</option>
            <option <% If Year(Now)="2002" Then%>selected<% End If%> value="2002">2002</option>
            <option <% If Year(Now)="2003" Then%>selected<% End If%> value="2003">2003</option>
            <option <% If Year(Now)="2004" Then%>selected<% End If%> value="2004">2004</option>
            <option <% If Year(Now)="2005" Then%>selected<% End If%> value="2005">2005</option>
            <option <% If Year(Now)="2006" Then%>selected<% End If%> value="2006">2006</option>
            <option <% If Year(Now)="2007" Then%>selected<% End If%> value="2007">2007</option>
            <option <% If Year(Now)="2008" Then%>selected<% End If%> value="2008">2008</option>
            <option <% If Year(Now)="2009" Then%>selected<% End If%> value="2009">2009</option>
            <option <% If Year(Now)="2010" Then%>selected<% End If%> value="2010">2010</option>
          </select></font></td>
          <td width="38"><small><font face="Arial">Month</font></small></td>
          <td width="77"><font face="Arial" size="1"><select name="EndMonth" size="1">
            <option <% If Month(Now)="1" Then%>selected<% End If%> value="01">Jan</option>
            <option <% If Month(Now)="2" Then%>selected<% End If%> value="02">Feb</option>
            <option <% If Month(Now)="3" Then%>selected<% End If%> value="03">Mar</option>
            <option <% If Month(Now)="4" Then%>selected<% End If%> value="04">Apr</option>
            <option <% If Month(Now)="5" Then%>selected<% End If%> value="05">May</option>
            <option <% If Month(Now)="6" Then%>selected<% End If%> value="06">Jun</option>
            <option <% If Month(Now)="7" Then%>selected<% End If%> value="07">Jul</option>
            <option <% If Month(Now)="8" Then%>selected<% End If%> value="08">Aug</option>
            <option <% If Month(Now)="9" Then%>selected<% End If%> value="09">Sep</option>
            <option <% If Month(Now)="10" Then%>selected<% End If%> value="10">Oct</option>
            <option <% If Month(Now)="11" Then%>selected<% End If%> value="11">Nov</option>
            <option <% If Month(Now)="12" Then%>selected<% End If%> value="12">Dec</option>
          </select></font></td>
          <td width="25"><small><font face="Arial">Day</font></small></td>
          <td width="54"><font face="Arial" size="1"><select name="EndDay" size="1">
            <option <% If Day(Now)="1" Then%>selected<% End If%> value="1">1</option>
            <option <% If Day(Now)="2" Then%>selected<% End If%> value="2">2</option>
            <option <% If Day(Now)="3" Then%>selected<% End If%> value="3">3</option>
            <option <% If Day(Now)="4" Then%>selected<% End If%> value="4">4</option>
            <option <% If Day(Now)="5" Then%>selected<% End If%> value="5">5</option>
            <option <% If Day(Now)="6" Then%>selected<% End If%> value="6">6</option>
            <option <% If Day(Now)="7" Then%>selected<% End If%> value="7">7</option>
            <option <% If Day(Now)="8" Then%>selected<% End If%> value="8">8</option>
            <option <% If Day(Now)="9" Then%>selected<% End If%> value="9">9</option>
            <option <% If Day(Now)="10" Then%>selected<% End If%> value="10">10</option>
            <option <% If Day(Now)="11" Then%>selected<% End If%> value="11">11</option>
            <option <% If Day(Now)="12" Then%>selected<% End If%> value="12">12</option>
            <option <% If Day(Now)="13" Then%>selected<% End If%> value="13">13</option>
            <option <% If Day(Now)="14" Then%>selected<% End If%> value="14">14</option>
            <option <% If Day(Now)="15" Then%>selected<% End If%> value="15">15</option>
            <option <% If Day(Now)="16" Then%>selected<% End If%> value="16">16</option>
            <option <% If Day(Now)="17" Then%>selected<% End If%> value="17">17</option>
            <option <% If Day(Now)="18" Then%>selected<% End If%> value="18">18</option>
            <option <% If Day(Now)="19" Then%>selected<% End If%> value="19">19</option>
            <option <% If Day(Now)="20" Then%>selected<% End If%> value="20">20</option>
            <option <% If Day(Now)="21" Then%>selected<% End If%> value="21">21</option>
            <option <% If Day(Now)="22" Then%>selected<% End If%> value="22">22</option>
            <option <% If Day(Now)="23" Then%>selected<% End If%> value="23">23</option>
            <option <% If Day(Now)="24" Then%>selected<% End If%> value="24">24</option>
            <option <% If Day(Now)="25" Then%>selected<% End If%> value="25">25</option>
            <option <% If Day(Now)="26" Then%>selected<% End If%> value="26">26</option>
            <option <% If Day(Now)="27" Then%>selected<% End If%> value="27">27</option>
            <option <% If Day(Now)="28" Then%>selected<% End If%> value="28">28</option>
            <option <% If Day(Now)="29" Then%>selected<% End If%> value="29">29</option>
            <option <% If Day(Now)="30" Then%>selected<% End If%> value="30">30</option>
            <option <% If Day(Now)="31" Then%>selected<% End If%> value="31">31</option>
          </select></font></td>
        </tr>
      </table>
      </div></td>
    </tr>
    <tr>
      <td align="right">&nbsp;</td>
      <td><input type="submit" value="Create Report" name="B1"></td>
    </tr>
  </table>
<table border="0" cellpadding="0" cellspacing="0" width="498">
  <tr>
    <td><img border="0" src="images/createreportbottom.gif" WIDTH="498" HEIGHT="30"></td>
  </tr>
</table>
  </center></div>
</form>

</body>
</html>
