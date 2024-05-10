<html>

<head>
<title></title>
</head>

<body>

<form method="POST" action="campaigns.asp?Task=AddNew">
  <div align="center">
    <center>
    <table border="0" cellpadding="0" cellspacing="0" width="590">
      <tr>
        <td align="center"><a href="help/campaigns.htm" target="_new"><img border="0" src="images/ListingofAllCampaigns.gif" WIDTH="590" HEIGHT="30"></a></td>
      </tr>
    </table>
    </center>
  </div>
  <div align="center"><center><table border="0" cellpadding="4" cellspacing="0" width="590" background="images/tableback.gif">
    <tr>
      <td align="right"><strong><font face="Arial" size="2"><div align="left"><p align="center"></font><font face="Arial" color="#0000A0" size="4">Select An Advertiser</font></strong>
        </div>
      </td>
    </tr>
    <tr>
      <td align="center"><font face="Arial"><select name="AdvertiserID" size="1">
<% 	Do While Not rsa.EOF  %>        <option value="<%=rsa("AdvertiserID")%>"><%=rsa("CompanyName")%></option>
<% 		rsa.MoveNext
	Loop
%>      </select></font></td>
    </tr>
    <tr>
      <td align="center">
        <div align="center">
          <center>
          <table border="0" cellpadding="0" cellspacing="0" width="180">
            <tr>
              <td width="229" colspan="2" align="center"><font face="Arial" size="3"><b>Campaign
                Type:</b></font></td>
            </tr>
            <tr>
              <td width="179" align="center"><font face="Arial" size="2">
                Banners</font></td>
              <td width="50"><font face="Arial" size="2"><input type="radio" value="Banners" name="CampaignType" checked></font></td>
            </tr>
            <tr>
              <td width="179" align="center"><font face="Arial" size="2">Static
                Text</font></td>
              <td width="50"><font face="Arial" size="2"><input type="radio" value="StaticText" name="CampaignType"></font></td>
            </tr>
          </table>
          </center>
        </div>
      </td>
    </tr>
    <tr>
      <td align="center">&nbsp;<font face="Arial"><input type="submit" value="Next &gt;&gt;" name="btnAddAdvertiser"></font></td>
    </tr>
  </table>
  </center></div>
  <div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="590">
    <tr>
      <td><img border="0" src="images/bottomblue.gif" WIDTH="590" HEIGHT="30"></td>
    </tr>
  </table>
  </center>
</div>
</form>

</body>
</html>
