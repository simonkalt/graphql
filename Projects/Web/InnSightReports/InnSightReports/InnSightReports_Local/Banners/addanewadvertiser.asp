<% 	If strTask="Edit" Then
		strTask2="Update"
		strButtonText="Update Advertiser"
	Else
		strTask2="Insert"
		strButtonText="Submit New Advertiser"
	End If
%>
<html>

<head>
<title></title>
</head>

<body>

<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.CompanyName.value == "")
  {
    alert("Please enter a value for the \"Company/Advertiser Name\" field.");
    theForm.CompanyName.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="advertisers.asp?Task=<%=strTask2%>&amp;AdvertiserID=<%=strAdvertiserID%>" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1">
  <div align="center">
    <center>
    <table border="0" cellpadding="0" cellspacing="0" width="590" bordercolor="#000080" background="images/tableback.gif">
      <tr>
        <td align="left"><a href="help/advertisers.htm" target="_new"><img border="0" src="images/ListingofAllAdvertisers.gif" alt="Click for more help on advertisers." WIDTH="590" HEIGHT="30"></a></td>
      </tr>
    </table>
    </center>
  </div>
  <div align="center"><center><table border="0" cellpadding="2" cellspacing="0" width="590" background="images/tableback.gif">
    <tr>
      <td width="206" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Company
      Name:</strong></font></td>
      <td width="370"><!--webbot bot="Validation" S-Display-Name="Company/Advertiser Name" B-Value-Required="TRUE" --><input type="text" name="CompanyName" size="35" <% If strTask="Edit" Then %>value="<%=rss("CompanyName")%>" <%End If%>></td>
    </tr>
    <!--Multi-Site option only -->
    <% If Application("BanManProMultiSite")=True Then %>
    <!--Multi-Site option only -->
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Run of
        Network:</strong></font></td>
      <td width="370"><input type="checkbox" name="RunOfNetwork" value="ON" <% If strTask="Edit" Then %><%If rss("UserID")=0 Then%>checked<%End If%><%End If%>><font face="Arial" size="2">
        </font><font face="Arial" size="1">(Available to all sites if checked)</font></td>
    </tr>
    <!--Multi-Site option only -->
    <% End If%>
    <!--Multi-Site option only -->
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Description:</strong></font></td>
      <td width="370"><input type="text" name="AdvertiserDesc" size="35" <% If strTask="Edit" Then %>value="<%=rss("AdvertiserDesc")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Website:</strong></font></td>
      <td width="370"><input type="text" name="CompanyWebSite" size="35" <% If strTask="Edit" Then %>value="<%=rss("CompanyWebSite")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Contact:</strong></font></td>
      <td width="370"><input type="text" name="Contact" size="35" <% If strTask="Edit" Then %>value="<%=rss("Contact")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Email:</strong></font></td>
      <td width="370"><input type="text" name="Email" size="35" <% If strTask="Edit" Then %>value="<%=rss("Email")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Login Name (For
      Reports):</strong></font></td>
      <td width="370"><input type="text" name="LoginName" size="35" <% If strTask="Edit" Then %>value="<%=rss("LoginName")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Login Password:</strong></font></td>
      <td width="370"><input type="text" name="LoginPassword" size="35" <% If strTask="Edit" Then %>value="<%=rss("LoginPassword")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Address 1:</strong></font></td>
      <td width="370"><input type="text" name="CompanyAddress1" size="35" <% If strTask="Edit" Then %>value="<%=rss("CompanyAddress1")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Address 2:</strong></font></td>
      <td width="370"><input type="text" name="CompanyAddress2" size="35" <% If strTask="Edit" Then %>value="<%=rss("CompanyAddress2")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>City:</strong></font></td>
      <td width="370"><input type="text" name="City" size="35" <% If strTask="Edit" Then %>value="<%=rss("City")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>State:</strong></font></td>
      <td width="370"><input type="text" name="State" size="35" <% If strTask="Edit" Then %>value="<%=rss("State")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Country:</strong></font></td>
      <td width="370"><input type="text" name="Country" size="35" <% If strTask="Edit" Then %>value="<%=rss("Country")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Zip:</strong></font></td>
      <td width="370"><input type="text" name="Zip" size="35" <% If strTask="Edit" Then %>value="<%=rss("Zip")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Telephone:</strong></font></td>
      <td width="370"><input type="text" name="Telephone" size="35" <% If strTask="Edit" Then %>value="<%=rss("Telephone")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Fax:</strong></font></td>
      <td width="370"><input type="text" name="Fax" size="35" <% If strTask="Edit" Then %>value="<%=rss("Fax")%>" <%End If%>></td>
    </tr>
  </center>
    <tr>
      <td width="576" align="right" colspan="2">
        <p align="left"><a href="help/advertisers.htm#Advertisers_EmailREports" target="_blank"><img border="0" src="images/emailreportstoadvertisers.gif" WIDTH="586" HEIGHT="25"></a></td>
    </tr>
    <center>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Email
        Daily Report:</strong></font></td>
      <td width="370"><input type="checkbox" name="DailyReport" value="-1" <% If strTask="Edit" Then %><%If rss("DailyReport")=True Then%>checked<%End If%><%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right"><font face="Arial" size="2"><strong>Email
        Weekly Report (Sun-Sat):</strong></font></td>
      <td width="370"><input type="checkbox" name="WeeklyReport" value="-1" <% If strTask="Edit" Then %><%If rss("WeeklyReport")=True Then%>checked<%End If%><%End If%>></td>
    </tr>
    <tr>
      <td width="206" align="right">&nbsp;</td>
      <td width="370"><input type="submit" value="<%=strButtonText%>" name="B1"></td>
    </tr>
    <tr>
      <td width="206" align="right">&nbsp;</td>
      <td width="370"><font face="Arial" size="2"><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Indicates Required Fields</font></td>
    </tr>
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
</form>

</div>
</body>
</html>