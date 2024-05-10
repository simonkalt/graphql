<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P>
<%
Dim objLastASPError

'Obtain an instance of the latest error
Set objLastASPError = Server.CreateObject("server.GetLastError")

'Output some of the properties
%>

An error occurred:<BR>
Description: <%=objLastASPError.Description%><BR>
Category: <%=objLastASPError.Category%><BR>
File: <%=objLastASPError.File%><BR>
Number: <%=objLastASPError.Number%><BR>

<br><br>
err: <%=err%><br>
my description: <%=err.description%><br>
my category: <%=err.category%><br>
my file: <%=err.file%><br>
my number: <%=err.number%><br>
</P>
</BODY>
</HTML>
