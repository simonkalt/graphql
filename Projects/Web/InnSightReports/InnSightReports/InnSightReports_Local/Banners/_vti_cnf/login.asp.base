
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Home Page</title>
</head>

<body>

<form method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?Login=True">
  <div align="center"><center><table border="0" cellpadding="1" cellspacing="0" width="354"
  bgcolor="#000000" background="images/blackback.gif">
    <tr>
      <td><table border="0" cellpadding="7" cellspacing="0" width="354" background="images/tableback.gif" >
        <tr>
          <td align="right" colspan="2" width="337"><font face="Arial" size="3"><strong><div
          align="left"><p><font color="#0000A0">Enter User Name and Password to Login</font></strong></font></td>
        </tr>
        <tr>
          <td align="right" width="102"><font face="Arial" size="2"><strong>User Name:</strong></font></td>
          <td width="251"><input type="text" name="UserName" size="20" value="<%If blnAdvertiserLogin=True Then%><%=Request.Cookies ("BanManPro")("AdvertiserName")%><%Else%><%=Request.Cookies("BanManPro")("UserName")%><%End If%>"></td>
        </tr>
        <tr>
          <td align="right" width="102"><font face="Arial" size="2"><strong>Password:</strong></font></td>
          <td width="251"><input type="password" name="Password" size="20" value="<%If blnAdvertiserLogin=True Then%><%=Request.Cookies ("BanManPro")("AdvertiserPassword")%><%Else%><%=Request.Cookies("BanManPro")("Password")%><%End If%>"></td>
        </tr>
               <% If Request.QueryString("Advertiser")<>"True" Then %>
       
       	<!--#include file="dbconnect.asp"-->
	
   <%
	If Application("BanManProMultiSite")=True Then 
	'connect to database
	Set connBanManPro=Server.CreateObject("ADODB.Connection") 
	connBanManPro.Mode = 3      '3 = adModeReadWrite
	connBanManPro.Open Application("BannerManagerConnectString")
	Set rs=connBanManPro.Execute("Select * From BanManProWebSites")
	
	%>
	 <tr>
          <td width="102">
            <p align="right"><font face="Arial" size="2"><strong>Web Site:</strong></font></td>
          <center>
          <td width="251"><select size="1" name="lstWebSites">
          <%Do While Not rs.EOF %>
              <option value="<%=rs("SiteID")%>"><%=rs("SiteName")%></option>
              <% rs.MoveNext
              Loop %>
            </select></td>
        </tr>
	<% End If
        End If%>
        <tr>
          <td colspan="2" width="337">
            <p align="center"><font face="Arial" size="2"><input type="checkbox" name="StoreCookie" value="ON" <%If blnAdvertiserLogin=True Then%><%If Request.Cookies ("BanManPro")("AdvertiserName")<> "" Then%>checked<%End If%><%Else%><%If Request.Cookies ("BanManPro")("UserName")<> "" Then%>checked<%End If%><%End If%>>
            Remember Username/Password in Cookie</font></td>
        </tr>
      </table>
  </center><table border="0" cellpadding="7" cellspacing="0" width="354" background="images/tableback.gif" >
        <tr background="images/tableback.gif">
          <td width="105">&nbsp;&nbsp; </td>
          <td width="217"><input type="submit" value="Login" name="btnLogin"></td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
  </center></div>
</form>
</body>
</html>


