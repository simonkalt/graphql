<%@ Language=VBScript %>

<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>

<HTML>
<HEAD>
<STYLE>
	.button	{ font-family:tahoma;font-size:11px;width=100px }
</STYLE>
</HEAD>
<BODY bottommargin=0 rightmargin=4 leftmargin=0 valign=bottom language=javascript topmargin=1 bgcolor=#F9D568>

<p align=right><input onclick=window.top.frames("frmGPFrame").doSave(); type=button id=cmdSave value="Save & Close" class=button>&nbsp;<input onclick=window.close(); type=button id=cmdCancel value=Cancel class=button>

</BODY>
</HTML>
