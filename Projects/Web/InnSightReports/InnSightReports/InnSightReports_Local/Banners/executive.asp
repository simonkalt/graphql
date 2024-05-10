<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Product:  Banner Manager Pro
'   Author:   Joe Rohrbach of Brookfield Consultants
'   Notes:    None
'                  
'
'                         COPYRIGHT NOTICE
'
'   The contents of this file are protected under the United States
'   copyright laws as an unpublished work, and are confidential and
'   proprietary to Brookfield Consultants.  Its use or disclosure in 
'   whole or in part without the expressed written permission of 
'   Brookfield Consultants is prohibited.
'
'   (c) Copyright 1999 by Brookfield Consultants.  All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
NumDays=DateDiff("d",strStartDate,strEndDate)
If NumDays=0 Then
	NumDays=1
End If
On Error Resume Next
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Ban Man Pro Executive Report</title>
</head>

<body>

<div align="center">
  <center>
  <table border="1" cellpadding="10" cellspacing="0" width="533" bordercolor="#000000" bgcolor="#C0C0C0">
    <tr>
      <td width="529">
        <p align="center"><font size="5" face="Verdana Ref"> Executive
        Advertising Report</font></td>
    </tr>
  </table>
  </center>
</div>
<p align="center"><font size="4" face="Verdana">Overall Summary for Report
Period<br>
</font><font face="Arial" size="2">(All Campaigns)</font></p>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" width="527">
    <tr>
      <td width="75" rowspan="2" align="center"><b><font face="Arial" size="2">Report
        Start Date</font></b></td>
      <td width="75" rowspan="2" align="center"><b><font face="Arial" size="2">Report
        End Date</font></b></td>
      <td width="150" colspan="2" align="center">
        <p align="center"><b><font face="Arial" size="2">Impressions</font></b></td>
      <td width="150" align="center" colspan="2"><b><font face="Arial" size="2">Clicks</font></b></td>
      <td width="75" align="center" rowspan="2"><b><font face="Arial" size="2">Overall
        Click Rate</font></b></td>
    </tr>
    <tr>
      <td width="75" align="center"><b><font face="Arial" size="2">Total</font></b></td>
      <td width="75" align="center"><b><font face="Arial" size="2">Avg/Day</font></b></td>
      <td width="75" align="center"><b><font face="Arial" size="2">Total</font></b></td>
      <td width="75" align="center"><b><font face="Arial" size="2">Avg/Day</font></b></td>
    </tr>
<% If Not rsReports.EOF Then %>
    <tr>
      <td width="75" align="center"><font face="Arial" size="2"><%=strStartDate%></font></td>
      <td width="75" align="center"><font face="Arial" size="2"><%=strEndDate%></font></td>
      <td width="75" align="center"><font face="Arial" size="2"><%=rsReports("SumOfImpressionCount")%></font></td>
      <td width="75" align="center"><font face="Arial" size="2"><%=CInt(rsReports("SumOfImpressionCount")/NumDays)%></font></td>
      <td width="75" align="center"><font face="Arial" size="2"><%=rsReports("SumOfClicks")%></font></td>
      <td width="75" align="center"><font face="Arial" size="2"><%=CInt(rsReports("SumOfClicks")/NumDays)%></font></td>
      <td width="75" align="center"><font face="Arial" size="2"><%=FormatPercent(rsReports("SumOfClicks")/rsReports("SumOfImpressionCount"))%></font></td>
    </tr>
<% End If %>
  </table>
  </center>
</div>
<p align="center"><font face="Verdana" size="4">Campaign Summary
</font><b><font face="Arial" size="4"> <br>
</font></b><font face="Arial" size="2">(All Campaigns Expiring after report
start period)<br>
(Actual Stats for Entire Campaign Period, not report period)</font></p>
<div align="center">
  <center>
  <table border="1" cellpadding="2" cellspacing="0" width="526">
    <tr>
      <td width="58" align="center"><font size="2" face="Arial"><b>Campaign</b></font></td>
      <td width="58" align="center"><font size="2" face="Arial"><b>Advertiser</b></font></td>
      <td width="58" align="center"><font size="2" face="Arial"><b>Start Date</b></font></td>
      <td width="58" align="center"><font size="2" face="Arial"><b>End Date</b></font></td>
      <td width="58" align="center"><font size="2" face="Arial"><b>Campaign Type</b></font></td>
      <td width="58" align="center"><font size="2" face="Arial"><b>Quantity<br>
        Sold</b></font></td>
      <td width="58" align="center"><font face="Arial" size="2"><b>Impressions</b></font></td>
      <td width="59" align="center"><font face="Arial" size="2"><b>Clicks</b></font></td>
      <td width="59" align="center"><font face="Arial" size="2"><b>Click Rate</b></font></td>
    </tr>
<% Do While Not rsReportsEx.EOF %>
    <tr>
      <td width="58" align="center"><font face="Arial" size="2"><%=rsReportsEx("CampaignName")%></font></td>
      <td width="58" align="center"><font face="Arial" size="2"><%=rsReportsEx("CompanyName")%></font></td>
      <td width="58" align="center"><font face="Arial" size="2"><%If rsReportsEx("CampaignStartDate")> Date() Then%><font color="#FF0000"><%End If%><%=rsReportsEx("CampaignStartDate")%><%If rsReportsEx("CampaignStartDate")> Date() Then%></font><%End If%></font></td>
      <td width="58" align="center"><font face="Arial" size="2"><%=rsReportsEx("CampaignEndDate")%></font></td>
      <td width="58" align="center"><font face="Arial" size="2"><%=rsReportsEx("CampaignType")%></font></td>
      <td width="58" align="center"><font face="Arial" size="2"><%=rsReportsEx("CampaignQuantitySold")%></font></td>
      <td width="58" align="center"><font face="Arial" size="2"><%=rsReportsEx("SumOfImpressionCount")%></font></td>
      <td width="59" align="center"><font face="Arial" size="2"><%=rsReportsEx("SumOfClicks")%></font></td>
      <td width="59" align="center"><font face="Arial" size="2"><%=FormatPercent(rsReportsEx("SumOfClicks")/rsReportsEx("SumOfImpressionCount"))%></font></td>
    </tr>
<% rsReportsEx.MoveNext
Loop %>
  </table>
  </center>
</div>

<p align="center"><font face="Verdana" size="4">Future Campaigns<br>
</font><font face="Arial" size="2">(All Campaigns Starting After Today)</font></p>
<div align="center">
  <center>
  <table border="1" cellpadding="2" cellspacing="0" width="527">
    <tr>
      <td align="center"><font size="2" face="Arial"><b>Campaign</b></font></td>
      <td align="center"><font size="2" face="Arial"><b>Advertiser</b></font></td>
      <td align="center"><font size="2" face="Arial"><b>Start Date</b></font></td>
      <td align="center"><font size="2" face="Arial"><b>End Date</b></font></td>
      <td align="center"><font size="2" face="Arial"><b>Campaign&nbsp;<br>
 Type</b></font></td>
      <td align="center"><font size="2" face="Arial"><b>Quantity<br>
        Sold</b></font></td>
    </tr>
    <% Do While Not rsReportsEx2.EOF %>
    <tr>
      <td align="center"><font face="Arial" size="2"><%=rsReportsEx2("CampaignName")%></font></td>
      <td align="center"><font face="Arial" size="2"><%=rsReportsEx2("CompanyName")%></font></td>
      <td align="center"><font face="Arial" size="2"><%=rsReportsEx2("CampaignStartDate")%></font></td>
      <td align="center"><font face="Arial" size="2"><%=rsReportsEx2("CampaignEndDate")%></font></td>
      <td align="center"><font face="Arial" size="2"><%=rsReportsEx2("CampaignType")%></font></td>
      <td align="center"><font face="Arial" size="2"><%=rsReportsEx2("CampaignQuantitySold")%></font></td>
    </tr>
    <% rsReportsEx2.MoveNext
Loop %>
  </table>
  </center>
</div>
<p align="center">&nbsp;</p>

</body>

</html>
