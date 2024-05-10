<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

dim table, idFieldID, idFieldName

table = Request.QueryString("table")
set cn = server.CreateObject("adodb.connection")
set rs = server.CreateObject("adodb.recordset")
cn.Open Application("sqlInnSight_ConnectionString")
set rs = cn.Execute("select * from " & table & " order by 2")

%>

<HTML>
<HEAD>
	<script>
		function md( o )
		{
		
		}
		
		function mdblclick( o )
		{
			var id = o.id.substr(o.id.indexOf("_")+1);
			var str = "EditDescription.asp?Mode=Edit&table=<%=table%>&idFieldName="+idFieldID+"&id="+id+"&DescriptionFieldName="+idFieldName;
			var result = window.showModalDialog(str,"","center: yes; dialogheight: 175px; dialogwidth: 440px; status: no; scroll: no;")
			window.document.location = window.document.location;
		}
	</script>
</HEAD>
<BODY onload="return window_onload()">
	<table border=1>
	<%do until rs.EOF
		Response.Write "<tr ondblclick=mdblclick(this) onmousedown=md(this) id=tr_" & rs.Fields(0).Value & ">"
		idFieldID = rs.Fields(0).Name
		idFieldName = rs.Fields(1).Name
		for i = 0 to rs.Fields.count-1
			Response.Write "<td>" & rs.Fields(i).Value & "</td>"
		next
		Response.Write "</tr>"
		rs.MoveNext
	loop
	
	rs.Close
	set rs = nothing
	cn.Close
	set cn = nothing%>
	</table>
</BODY>
</HTML>

<script>
		function window_onload()
		{
			idFieldID = "<%=idFieldID%>"
			idFieldName = "<%=idFieldName%>"
		}
</script>
