<% If Trim(Session("UserName"))<>"" then %>    
      <div align="center">
        <table border="0" cellpadding="1" cellspacing="0" width="500" background="images/blackback.gif">
          <tr>
            <td>
              <div align="center">
                <table border="0" cellpadding="3" cellspacing="0" width="500" background="images/tableback.gif">
                  <tr>
                    <td>
                      <p align="center"><font face="Arial" size="2">Current Server Date/Time: <%=Now%></font></p>
                      <% If Application("BanManProMultiSite")=True And Session("AdvertiserID")=0 Then 
                      		Set rsdt=connBanManPro.Execute("Select * From BanManProWebSites Order By SiteName ASC")
                      		If Trim(Request.Form("lstSite"))<>"" Then
                      				Session("BanManProSiteID")=Clng(Request.Form("lstSite"))
                     		 End If %>
                      <div align="center">
                        <center>
                        <table border="0" cellpadding="0" cellspacing="0" width="425" background="images/tableback.gif">
                          <tr>
                            <td align="center">
                                                    			<form method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME") %>">
                        		<p align="center"><font color="#000080" face="Arial" size="3">Web
                                </font><font face="Arial" size="3"><font color="#000080">Site</font>:
                      			</font><select size="1" name="lstSite">
                   				<%Do While Not rsdt.EOF %>
                       			<option value="<%=rsdt("SiteID")%>" <% If CLng(Session("BanManProSiteID"))=Clng(rsdt("SiteID")) Then%>selected<%End If%>><%=rsdt("SiteName")%></option>
                       			<% rsdt.MoveNext
                       			Loop %>
                        		&nbsp;
                        		</select> <input type="image" SRC="images/goround.gif" BORDER="0" id="image1" name="image1" align="absmiddle" WIDTH="28" HEIGHT="28"></p>
                      			</form>
                      			</td>
                          </tr>
                        </table>
                        </center>
                      </div>

				<% Set rsdt=Nothing%>
                      <%End If %>
			<% ' Advertisers Page 
                      If Trim(Cstr(strAdvertiserID))<>"" THen %>
                      		<p align="center">
                     		<font face="Arial" size="3">
                     		<font color="#000080">Advertisers:</font>
                      		</font><font face="Arial" size="2">
<%				Set rsLetters=connBanManPro.Execute("Select CompanyName From Advertisers Where (UserID=" & CLng(Session("BanManProSiteID")) & ") AND CompanyName Like '[a-z]%' Order By CompanyName ASC")
				strLetter="Z"
				Do While Not rsLetters.EOF
					If UCase(Left(rsLetters("CompanyName"),1))<>strLetter Then
						strLetter=UCase(Left(rsLetters("CompanyName"),1))
						Response.Write "<a href=" & Chr(34) & "advertisers.asp?Letter=" & strLetter & Chr(34) & ">" & strLetter & "</a>,"
        				End If
					rsLetters.MoveNext
				Loop
				Set rsLetters=Nothing
				Set rsLetters=connBanManPro.Execute("Select CompanyName From Advertisers Where (UserID=" & CLng(Session("BanManProSiteID")) & " ) And (CompanyName < 'a%' Or CompanyName > 'z%') Order By CompanyName ASC")
				If Not rsLetters.EOF Then 
					%>
					<a href="advertisers.asp?Letter=Other">Other</a>,
					<%
				End If
%>
				<a href="advertisers.asp?Task=ViewAll">
                      		ALL</a></font></p>
                     <% ElseIf Trim(Cstr(strZoneID))<>"" THen %>
                      		<p align="center">
                     		<font face="Arial" size="3">
                     		<font color="#000080">Zones:</font>
                      		</font><font face="Arial" size="2">
<%				Set rsLetters=connBanManPro.Execute("Select ZoneDescription From Zones Where (UserID=" & CLng(Session("BanManProSiteID")) & ") AND ZoneDescription Like '[a-z]%' Order By ZoneDescription ASC")
				strLetter="Z"
				Do While Not rsLetters.EOF
					If UCase(Left(rsLetters("ZoneDescription"),1))<>strLetter Then
						strLetter=UCase(Left(rsLetters("ZoneDescription"),1))
						Response.Write "<a href=" & Chr(34) & "zones.asp?Letter=" & strLetter & Chr(34) & ">" & strLetter & "</a>,"
        				End If
					rsLetters.MoveNext
				Loop
				Set rsLetters=Nothing
				Set rsLetters=connBanManPro.Execute("Select ZoneDescription From Zones Where (UserID=" & CLng(Session("BanManProSiteID")) & " ) And (ZoneDescription < 'a%' Or ZoneDescription > 'z%') Order By ZoneDescription ASC")
				If Not rsLetters.EOF Then 
					%>
					<a href="zones.asp?Letter=Other">Other</a>,
					<%
				End If
%>
				<a href="zones.asp?Task=ViewAll">
                      		ALL</a></font></p>
                      <% ElseIf Trim(Cstr(strBannerID))<>""  Or Trim(Cstr(strCampaignID))<>"" Then %>
                     		 <!--Start of list of advertisers-->
    							<% 	If UCase(Session("UserName"))=UCase(Application("AdministratorName")) And UCase(Session("Password"))=UCase(Application("AdministratorPassword")) Then %>
								<% If Request("Task")="ViewAll" Or Request("Task")="" Or Request("Task")="Expired" Or Request("Task")="Update" Or Request("Task")="UpdateAdvanced" Then %>
								<!--#Include File="listadvertisers.asp"-->
								<% End If %>
   						 		<%End If %>
								<!--End of list of advertisers-->
                      <% End If %>
                   
                    </td>
                  </tr>
                </table>
              </div>
            </td>
          </tr>
        </table>
      </div>
          <p align="center">
	  <font face="Arial" size="3"><br>
<% End If %>