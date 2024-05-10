<% 	If strTask="Edit" Then
		strTask2="Update"
		strButtonText="Update Banner"
	Else
		strTask2="Insert"
		strButtonText="Submit New Banner"
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

  if (theForm.AdDescription.value == "")
  {
    alert("Please enter a value for the \"AdDescription\" field.");
    theForm.AdDescription.focus();
    return (false);
  }

  if (theForm.AdTargetURL.value == "")
  {
    alert("Please enter a value for the \"AdTargetURL\" field.");
    theForm.AdTargetURL.focus();
    return (false);
  }

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
    alert("Please enter only digit characters in the \"Banner Height\" field.");
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
    alert("Please enter only digit characters in the \"Banner Width\" field.");
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
  <div align="center">
    <center>
    <table border="0" cellpadding="0" cellspacing="0" width="590" bordercolor="#000080" style="border: medium none" background="images/tableback1.gif">
      <tr>
        <td align="center" bgcolor="#808080"><a href="help/banners.htm" target="_new"><img border="0" src="images/ListingofAllbanners.gif" WIDTH="590" HEIGHT="30"></a></td>
      </tr>
    </table>
    </center>
  </div>
  <div align="center"><center><table border="0" cellpadding="2" cellspacing="0" width="591" background="images/tableback1.gif">
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><b><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Advertiser:</b></font></td>
      <td width="403" height="24"><select name="AdvertiserID" size="1">
<% 	If strTask="Edit" Then %>        <option selected value="<%=rsb("AdvertiserID")%>"><%=rsb("CompanyName")%></option>
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
%>      </select></td>
    </tr>
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><b><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Banner Ad
        Desc:</b></font></td>
      <td width="403" height="18"><!--webbot bot="Validation" S-Display-Name="AdDescription" B-Value-Required="TRUE" --><input type="text" name="AdDescription" size="40" <% If strTask="Edit" Then %>value="<%=rsb("AdDescription")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><b><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Target URL:</b></font></td>
      <td width="403" height="25"><!--webbot bot="Validation" S-Display-Name="AdTargetURL" B-Value-Required="TRUE" --><input type="text" name="AdTargetURL" size="40" <% If strTask="Edit" Then %>value="<%=rsb("AdTargetURL")%>" <%End If%>></td>
    </tr>
    <!--Multi-Site option only -->
    <% If Application("BanManProMultiSite")=True Then %>
    <!--Multi-Site option only -->
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><strong>Run of
        Network:</strong></font></td>
      <td width="403" height="25"><input type="checkbox" name="RunOfNetwork" value="ON" <% If strTask="Edit" Then %><%If rsb("UserID")=0 Then%>checked<%End If%><%End If%>><font face="Arial" size="2">
        </font><font face="Arial" size="1">(Available to all sites if checked)</font></td>
    </tr>
    <!--Multi-Site option only -->
    <% End If %>
    <!--Multi-Site option only -->
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><b><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Type:</b></font></td>
      <td width="403" height="25"><font face="Arial" size="2"><input type="radio" value="0" <% If strTask="Edit" Then %><%If rsb("AdTextLink")=True Then%><%Else%>checked<%End If%><%Else%>checked<%End If%> name="AdTextLink" style="background-image: url('images/tableback1.gif')">Image
      Ad&nbsp;&nbsp; <input type="radio" name="AdTextLink" value="-1" <% If strTask="Edit" Then %><%If rsb("AdTextLink")=True Then%>checked<%Else%><%End If%><%Else%><%End If%> style="background-image: url('images/tableback1.gif')"> Text Link Ad</font></td>
    </tr>
  </center>
    <tr>
      <td align="right" colspan="2">
        <p align="left"><font face="Arial" size="2"><b><a href="help/banners.htm#Banners_InformationForImageAds" target="_blank"><img border="0" src="images/informationforimageads.gif" WIDTH="586" HEIGHT="26"></a></b></font></td>
    </tr>
    <center>
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><b><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Image URL:</b></font></td>
      <td width="403" height="25"><input type="text" name="AdImageURL" size="40" <% If strTask="Edit" Then %>value="<%=rsb("AdImageURL")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><b>Alt Text:</b></font></td>
      <td width="403" height="25"><input type="text" name="AdAltText" size="40" <% If strTask="Edit" Then %>value="<%=rsb("AdAltText")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><b><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Width:</b></font></td>
      <td width="403" height="25"><!--webbot bot="Validation" S-Display-Name="Banner Height" S-Data-Type="Number" S-Number-Separators=",." --><input type="text" name="AdWidth" size="10" <% If strTask="Edit" Then %>value="<%=rsb("AdWidth")%>" <%Else%>value="468" <%End If%>></td>
    </tr>
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><b><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Height:</b></font></td>
      <td width="403" height="25"><!--webbot bot="Validation" S-Display-Name="Banner Width" S-Data-Type="Number" S-Number-Separators=",." --><input type="text" name="AdHeight" size="10" <% If strTask="Edit" Then %>value="<%=rsb("AdHeight")%>" <%Else%>value="60" <%End If%>></td>
    </tr>
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><b>Border:</b></font></td>
      <td width="403" height="25"><select name="AdBorder" size="1">
<% If strTask="Edit" Then %>        <option <% If rsb("AdBorder")=0 Then%>selected<%End If%> value="0">0</option>
        <option <% If rsb("AdBorder")=1 Then%>selected<%End If%> value="1">1</option>
        <option <% If rsb("AdBorder")=2 Then%>selected<%End If%> value="2">2</option>
        <option <% If rsb("AdBorder")=3 Then%>selected<%End If%> value="3">3</option>
        <option <% If rsb("AdBorder")=4 Then%>selected<%End If%> value="4">4</option>
        <option <% If rsb("AdBorder")=5 Then%>selected<%End If%> value="5">5</option>
<% Else %>        <option selected value="0">0</option>
        <option value="1">1</option>
        <option value="2">2</option>
        <option value="3">3</option>
        <option value="4">4</option>
        <option value="5">5</option>
<% End If %>      </select></td>
    </tr>
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><b>Alignment:</b></font></td>
      <td width="403" height="25"><select name="AdAlign" size="1">
<% If strTask="Edit" Then %>        <option <% If rsb("AdAlign")="Center" Then%>selected<%End If%> value="Center">Center</option>
        <option <% If rsb("AdAlign")="Left" Then%>selected<%End If%> value="Left">Left</option>
        <option <% If rsb("AdAlign")="Right" Then%>selected<%End If%> value="Right">Right</option>
<% Else %>        <option selected value="Center">Center</option>
        <option value="Left">Left</option>
        <option value="Right">Right</option>
<% End If %>      </select></td>
    </tr>
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><b>Optional
      Text Underneath:</b></font></td>
      <td width="403" height="25"><input type="text" name="AdTextUnderneath" size="40" <% If strTask="Edit" Then %>value="<%=rsb("AdTextUnderneath")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><b>Launch In
      New Window:</b></font></td>
      <td width="403" height="25"><select name="AdNewWindow" size="1">
<% If strTask="Edit" Then %>        <option <% If rsb("AdNewWindow")=-1 Then%>selected<%End If%> value="-1">Yes</option>
        <option <% If rsb("AdNewWindow")= 0 Then%>selected<%End If%> value="0">No</option>
<% Else %>        <option selected value="-1">Yes</option>
        <option value="0">No</option>
<% End If %>      </select></td>
    </tr>
    <tr>
      <td width="579" align="left" colspan="2"><font face="Arial" size="2"><b><a href="help/banners.htm#Banners_TextLinks" target="_blank"><img border="0" src="images/informationforstatictextlinks.gif" WIDTH="586" HEIGHT="25"> </a> </b></font> </td>
    </tr>
    <tr>
      <td width="176" align="right"><font face="Arial" size="2"><b><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Link Text:</b></font> </td>
      <td width="403" height="25"><input type="text" name="AdTextLinkText" size="40" <% If strTask="Edit" Then %>value="<%=rsb("AdTextLinkText")%>" <%End If%>></td>
    </tr>
    <tr>
      <td width="176" align="right">&nbsp; </td>
      <td width="403" height="25"><font face="Arial"><input type="submit" value="<%=strButtonText%>" name="B1"></font></td>
    </tr>
    <tr>
      <td width="176" align="right">&nbsp; </td>
      <td width="403" height="25"><font face="Arial" size="2"><img src="images/required2.gif" WIDTH="14" HEIGHT="12">Indicates Required Fields</font></td>
    </tr>
  </table>
  </center></div>
  <div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="590" bordercolor="#000080" style="border: medium none" background="images/tableback1.gif">
    <tr>
      <td align="center" bgcolor="#808080"><img border="0" src="images/bottomblue.gif" WIDTH="590" HEIGHT="30"></td>
    </tr>
  </table>
  </center>
</form>
</body>
</html>