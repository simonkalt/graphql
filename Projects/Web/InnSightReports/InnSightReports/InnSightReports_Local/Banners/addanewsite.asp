<% 	If strTask="Edit" Then
		strTask2="Update"
		strButtonText="Update Site Name"
	Else
		strTask2="Insert"
		strButtonText="Submit New Site"
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

  if (theForm.SiteName.value == "")
  {
    alert("Please enter a value for the \"Web Site Name\" field.");
    theForm.SiteName.focus();
    return (false);
  }

  if (theForm.SiteName.value.length > 50)
  {
    alert("Please enter at most 50 characters in the \"Web Site Name\" field.");
    theForm.SiteName.focus();
    return (false);
  }

  if (theForm.SiteURL.value == "")
  {
    alert("Please enter a value for the \"Web Site URL\" field.");
    theForm.SiteURL.focus();
    return (false);
  }

  if (theForm.SiteURL.value.length > 255)
  {
    alert("Please enter at most 255 characters in the \"Web Site URL\" field.");
    theForm.SiteURL.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="sites.asp?Task=<%=strTask2%>&amp;SiteID=<%=strSiteID%>" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1">
  <div align="center">
    <center>
    <table border="0" cellpadding="0" cellspacing="0" width="590" bordercolor="#000080" background="images/tableback.gif">
      <tr>
        <td align="left"><a href="help/sites.htm" target="_blank"><img border="0" src="images/banmanprosites.gif" WIDTH="590" HEIGHT="30"></a></td>
      </tr>
    </table>
    </center>
  </div>
  <div align="center"><center><table border="0" cellpadding="2" cellspacing="1" width="590" background="images/tableback.gif">
    <tr>
      <td width="135" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Web
        Site Name:</strong></font></td>
      <td width="441"><!--webbot bot="Validation" S-Display-Name="Web Site Name" B-Value-Required="TRUE" I-Maximum-Length="50" --><input type="text" name="SiteName" size="50" <% If strTask="Edit" Then %>value="<%=rss("SiteName")%>" <%End If%> maxlength="50"></td>
    </tr>
    <tr>
      <td width="135" align="right"><img src="images/required2.gif" WIDTH="14" HEIGHT="12"><font face="Arial" size="2"><strong>Web
        Site URL:</strong></font></td>
      <td width="441"><!--webbot bot="Validation" S-Display-Name="Web Site URL" B-Value-Required="TRUE" I-Maximum-Length="255" --><input type="text" name="SiteURL" size="50" <% If strTask="Edit" Then %>value="<%=rss("SiteURL")%>" <%End If%> maxlength="255"></td>
    </tr>
    <tr>
      <td width="135" align="right">&nbsp;</td>
      <td width="441"><input type="submit" value="<%=strButtonText%>" name="B1"></td>
    </tr>
    <tr>
      <td width="135" align="right">&nbsp;</td>
      <td width="441"><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Indicates Required Fields</td>
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