<% 	If strTask="ViewCode" Then
		strTask2="UpdateAdvanced"
		strButtonText="Update Banner"
	Else
		strTask2="InsertAdvanced"
		strButtonText="Submit New Banner Using Code Above"
	End If
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title></title>
</head>

<body>

<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript"><!--
function FrontPage_Form1_Validator(theForm)
{

  var checkOK = "0123456789-.,";
  var checkStr = theForm.AdWidth.value;
  var allValid = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch == ".")
    {
      allNum += ".";
      decPoints++;
    }
    else if (ch != ",")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Banner Width\" field.");
    theForm.AdWidth.focus();
    return (false);
  }

  if (decPoints > 1)
  {
    alert("Please enter a valid number in the \"AdWidth\" field.");
    theForm.AdWidth.focus();
    return (false);
  }

  var checkOK = "0123456789-.,";
  var checkStr = theForm.AdHeight.value;
  var allValid = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch == ".")
    {
      allNum += ".";
      decPoints++;
    }
    else if (ch != ",")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Banner Height\" field.");
    theForm.AdHeight.focus();
    return (false);
  }

  if (decPoints > 1)
  {
    alert("Please enter a valid number in the \"AdHeight\" field.");
    theForm.AdHeight.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="banners.asp?Task=<%=strTask2%>&amp;BannerID=<%=strBannerID%>" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1">
  <div align="center"><center>
    <table border="0" cellpadding="0" cellspacing="0" width="590" bordercolor="#000080" style="border: medium none" background="images/tableback.gif">
      <tr>
        <td align="center" bgcolor="#808080"><a href="help/banners.htm" target="_new"><img border="0" src="images/ListingofAllbanners.gif" WIDTH="590" HEIGHT="30"></a></td>
      </tr>
    </table>
    <table border="0" cellpadding="5" cellspacing="0" width="590" background="images/tableback.gif">
    <tr>
      <td width="578" colspan="2"><font face="Arial" size="2">Use this option to add a new
      banner using your own Ad Code.&nbsp; This feature is necessary for using third party ad
      code through various advertising agencies such as BurstMedia, ValueClick, FlyCast, etc.&nbsp;
        <a href="http://www.banmanpro.com/support/3rdpartycode.asp" target="_new">Click
        here for more information on 3rd Party Ad Code.</a>&nbsp; If using SSI or an
        ASP function call for serving ads, only the first three parameters are
        required.<br>
        </font>
        <hr>
      </td>
    </tr>
    <tr>
      <td width="135"><font face="Arial" size="2"><strong><div align="right"><p><img border="0" src="images/req.gif" WIDTH="14" HEIGHT="12">Advertiser:</strong></font>
        </div>
      </td>
      <td width="431"><font face="Arial"><select name="AdvertiserID" size="1">
<% 	If strTask="ViewCode" Then %>        <option selected value="<%=rsb("AdvertiserID")%>"><%=rsb("CompanyName")%></option>
<%		Do While Not rsa.EOF   
		    If (rsa("AdvertiserID") <> rsb("AdvertiserID")) And (rsa("CompanyName") <> rsb("CompanyName")) Then %>        <option value="<%=rsa("AdvertiserID")%>"><%=rsa("CompanyName")%></option>
<%		    End If
		rsa.MoveNext
		Loop
	Else
		Do While Not rsa.EOF  %>        <option value="<%=rsa("AdvertiserID")%>"><%=rsa("CompanyName")%></option>
<% 		rsa.MoveNext
		Loop
	End If
%>      </select></font></td>
    </tr>
    <tr>
      <td width="135"><font face="Arial" size="2"><strong><div align="right"><p><img border="0" src="images/req.gif" WIDTH="14" HEIGHT="12">Ad Description:</strong></font>
        </div>
      </td>
      <td width="431"><input type="text" name="AdDescription" size="40" <% If strTask="ViewCode" Then %>value="<%=rsb("AdDescription")%>" <%End If%>></td>
    </tr>
    <!--Multi-Site option only -->
    <% If Application("BanManProMultiSite")=True Then %>
    <!--Multi-Site option only -->
    <tr>
      <td width="135" align="right"><font face="Arial" size="2"><strong>Run of
        Network:</strong></font></td>
      <td width="431" height="25"><input type="checkbox" name="RunOfNetwork" value="ON" <% If strTask="ViewCode" Then %><%If rsb("UserID")=0 Then%>checked<%End If%><%End If%>><font face="Arial" size="2">
        </font><font face="Arial" size="1">(Available to all sites if checked)</font></td>
    </tr>
    <!--Multi-Site option only -->
    <% End If %>
    <!--Multi-Site option only -->
    <tr>
      <td width="135"><div align="right"><p><font face="Arial" size="2"><strong><img border="0" src="images/req.gif" WIDTH="14" HEIGHT="12">Ad Code:</strong></font>
        </div>
      </td>
      <td width="431"><textarea rows="10" name="AdCode" cols="40"><%If strTask="ViewCode" Then %><%=rsb("AdCode")%><%Else%>Paste Your Ad Code Here<%End If%></textarea><br>
        <font face="Arial" size="2"><a href="http://www.banmanpro.com/support/trackclicks.asp" target="_blank">Click
        here for list of Advanced Banner Parameters.</a></font></td>
    </tr>
</center>
    <tr>
      <td width="135">
        <p align="right"><font face="Arial" size="2"><strong>Target URL:</strong></font></td>
      <td width="431"><input type="text" name="TargetURL" size="40" value="<% If strTask="ViewCode" Then %><%=rsb("AdTargetURL")%><%End If%>"></td>
    </tr>
    <tr>
      <td width="566" colspan="2">
        <hr>
        <p><b><font face="Arial" size="3">Parameters for Non SSI ad serving
        code.&nbsp;</font></b><font face="Arial" size="2"> All parameters below
        are required only if using the Advanced Javascript or Non
        Cache-Defeating code.&nbsp; You must specify
        an image source and target URL for any non-Java script compliant
        browsers to prevent a broken link on older browsers.&nbsp;&nbsp;</font></td>
    </tr>
    <tr>
      <td width="135">
        <p align="right"><font face="Arial" size="2"><strong>Use Code above in Netscape:</strong></font></td>
      <center>
      <td width="431">
        <table border="0" cellpadding="0" cellspacing="0" width="315">
          <tr>
            <td width="41"><input type="checkbox" name="Netscape4" value="-1" <%If strTask="ViewCode" Then %><%If rsb("AdNewWindow")<>0 Then%>checked<%End If%><%End If%>></td>
            <td><font face="Arial" size="2">Note: When using Burst Media and
              Flycast advanced code do not check.&nbsp;</font></td>
          </tr>
        </table>
      </td>
    </tr>
</center>
    <tr>
      <td width="135">
        <p align="right"><font face="Arial" size="2"><strong>OR&nbsp;</strong></font></td>
      <td width="431">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
    <tr>
      <td width="135">
        <p align="right"><font face="Arial" size="2"><strong>Use this code in
        Netscape browsers : </strong></font></td>
      <td width="431"><textarea rows="10" name="AdCodeNetscape" cols="40"><%If strTask="ViewCode" Then %><%=rsb("AdCodeNetscape")%><%Else%><%End If%></textarea></td>
    </tr>
    <tr>
      <td width="135">
        <p align="right"><font face="Arial" size="2"><strong>Image Source:</strong></font></td>
      <center>
      <td width="431"><input type="text" name="ImageSource" size="40" value="<% If strTask="ViewCode" Then %><%=rsb("AdImageURL")%><%End If%>"></td>
    </tr>
</center>
      <center>
</center>
  <tr>
      <td width="135">
        <p align="right"><font face="Arial" size="2"><strong>Width:</strong></font></td>
      <td width="431"><!--webbot bot="Validation" S-Display-Name="Banner Width" S-Data-Type="Number" S-Number-Separators=",." --><input type="text" name="AdWidth" size="10" value="<% If strTask="ViewCode" Then %><%=rsb("AdWidth")%><%End If%>"></td>
  </tr>
    <tr>
      <td width="135">
        <p align="right"><font face="Arial" size="2"><strong>Height:</strong></font></td>
      <center>
      <td width="431"><!--webbot bot="Validation" S-Display-Name="Banner Height" S-Data-Type="Number" S-Number-Separators=",." --><input type="text" name="AdHeight" size="10" value="<% If strTask="ViewCode" Then %><%=rsb("AdHeight")%><%End If%>"></td>
    </tr>
</center>
      <center>
    <tr>
      <td width="135">&nbsp;&nbsp; </td>
      <td width="431"><input type="submit" value="<%=strButtonText%>" name="B1"></td>
    </tr>
  </table>
</center>
    <div align="center">
      <center>
      <table border="0" cellpadding="0" cellspacing="0" width="590" bordercolor="#000080" style="border: medium none" background="images/tableback.gif">
        <tr>
          <td align="center" bgcolor="#808080"><img border="0" src="images/bottomblue.gif" WIDTH="590" HEIGHT="30"></td>
        </tr>
      </table>
      </center>
    </div>
</form>

</body>
</html>