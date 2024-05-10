<% If Request("Task")="Expired" Then
	strTempTask="Expired"
Else
	strTempTask="ViewAll"
End If
%>
<% Set rsAdvertiser=connBanManPro.Execute("Select CompanyName,AdvertiserID From Advertisers Where (UserID=" & CLng(Session("BanManProSiteID")) & " Or UserID=0) Order By CompanyName ASC") %>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.AdvertiserID.selectedIndex == 0)
  {
    alert("The first \"Select An Advertiser\" option is not a valid selection.  Please choose one of the other options.");
    theForm.AdvertiserID.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?Task=<%=strTempTask%>" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="454">
            <tr>
              <td width="448" align="center">
                <p align="center">       
        <font face="Arial" size="3" color="#000080">Advertiser:</font>    <!--webbot bot="Validation" S-Display-Name="Select An Advertiser" B-Disallow-First-Item="TRUE" --><select size="1" name="AdvertiserID">
        <option value>Select An Advertiser</option>
         <% Do While Not rsAdvertiser.EOF %>    
          <option <% If Request.Form("AdvertiserID") <> "" Then%><%If Clng(Request.Form("AdvertiserID"))=Clng(rsAdvertiser("AdvertiserID")) Then%>selected<%end If%><%End If%> value="<%=rsAdvertiser("AdvertiserID")%>"><%=rsAdvertiser("CompanyName")%></option>
          <% rsAdvertiser.MoveNext
          Loop %>
        </select> <input type="image" SRC="images/goround.gif" BORDER="0" id="image1" name="image1" align="absmiddle" WIDTH="28" HEIGHT="28"></td>
            </tr>
          </table>
        </div>
      </form>
    <p align="center">
<% Set rsAdvertiser=Nothing %>