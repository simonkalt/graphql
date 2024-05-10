<%
	'Set cnHeaderSQL = Server.CreateObject("ADODB.Connection")
	'Set rsHeader = Server.CreateObject("ADODB.Recordset")
  
	'cnHeaderSQL.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")

	'Set rsHeader = cnHeaderSQL.Execute("sp_MarqueeLookup " & Session("CompanyID"))
	'intID= rsHeader.Fields(0).Value
	'Set rsHeader = cnHeaderSQL.Execute("Select * from tblMarqueeMessages Where ID=" & intID)
%>

<!--
TABLE style="WIDTH: 780px; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-TOP: windowtext 0.5pt solid; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 0.5pt solid; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 0.5pt solid" 
cellSpacing=0 borderColorDark=black cellPadding=0 background="images/bluebg.gif">
<tr>
<td>
<marquee width="750" height="20">
<font face="Tahoma" size="2" color="white"><%'=rsHeader.Fields("MessageText")%></font>
</marquee>
</td>
</tr>
</table-->



<!--TABLE WIDTH="750" cellSpacing=0 cellPadding=0>
<tr>
<TD valign=top style='BORDER-RIGHT: windowtext 0.5pt solid; BORDER-TOP: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid' bgcolor=#0a246a>
<marquee width="750" height="20" bgcolor="#0a246a">
<font face="Tahoma" size="2" color="white">InnSight Concierge Banner Text Here</font>
</td>
</tr>
</table-->


<table cellspacing=0 cellspacing=0 width=755 border=0 bgcolor=silver> <!--background=images/HeaderBlend.jpg-->
	<tr>
		<!--td align="left">
		</td-->
		<td align=center> <%' width=200%>
			<!--applet code="fprotate.class" codebase="Applets/" width="280" height="34">
              <param name="rotatoreffect" value="disolve">
			  <param name="time" value="7">
			  <param name="url" valuetype="ref" value="#">
			  <param name="image1" valuetype="ref" value="images/Banners/TestBanner.gif">
			  <param name="image2" valuetype="ref" value="images/Banners/FordBanner.gif">
			  <param name="image3" valuetype="ref" value="images/Banners/ApStumpsBanner.gif">
			</applet-->
			
			<!-- Begin Ban Man Pro Banner Code -  Zone: All Hotels -->
			<SCRIPT LANGUAGE="JAVASCRIPT">
			<!--
			var browName = navigator.appName;
			var browVersion = parseInt(navigator.appVersion);
			var ua=navigator.userAgent.toLowerCase();
			var adcode='';
			if (browName=='Netscape'){
			     if ((browVersion>=4)&&(ua.indexOf("mac")==-1))
			          { document.write('<S'+'CRIPT src="http://www.wbwd.net/innsight/banners/banman.asp?ZoneID=1&Task=Get&Browser=NETSCAPE4">');
			          document.write('</'+'scr'+'ipt>');
			          document.write(adcode); }
			     else if (browVersion>=3) 
			          { document.write('<A HREF="http://www.wbwd.net/innsight/banners/banman.asp?ZoneID=1&Task=Click&Mode=HTML" target="_new"><IMG SRC="http://www.wbwd.net/innsight/banners/banman.asp?ZoneID=1&Task=Get&Mode=HTML&fightcache=' + (new Date()).getTime() + '" width="230" height="33" border="0"></A>'); } }
			if (browName=='Microsoft Internet Explorer')
			     { document.write('<ifr'+'ame src="http://www.wbwd.net/innsight/banners/banman.asp?ZoneID=1&Task=Get" width=230 height=33 Marginwidth=0 Marginheight=0 Hspace=0 Vspace=0 Frameborder=0 Scrolling=No></ifr'+'ame>'); }
			// --> 
			</script>
			<noscript>
			    <a href="http://www.wbwd.net/innsight/banners/banman.asp?ZoneID=1&Task=Click&Mode=HTML&PageID=84416" target="_new">
			    <img src="http://www.wbwd.net/innsight/banners/banman.asp?ZoneID=1&Task=Get&Mode=HTML&PageID=84416" width="230" height="33" border="0"></a>
			</noscript>
			<!-- End Ban Man Pro Banner Code -  Zone: All Hotels -->		
		
		</td>
		<!--td align="right">
		</td-->
	</tr>
</table>