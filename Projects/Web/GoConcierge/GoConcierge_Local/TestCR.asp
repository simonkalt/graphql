<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<TEXTAREA wrap=hard width=100px rows=12 cols=20 id=textarea1 name=textarea1></TEXTAREA>
<INPUT type="button" value="Button" id=button1 name=button1 language=javascript onclick="doCR()">
</BODY>
</HTML>

<script language=javascript>
function doCR()
{
//var tr = window.textarea1.createTextRange()
var obj = window.textarea1.getClientRects()
for(var i=0;i<obj.length;i++)
	alert(i);
//window.textarea1.innerHTML = window.textarea1.innerHTML.replace(/\n/g,'zzzzzzzzz');
}
</script>

<script language=vbscript>
	x = alert(right("0" & Month(now()),2) & right("0" & Day(now()),2) & right(Year(now()),2) & right("0" & Hour(now()),2) & right("0" & minute(now()),2))
	
	set oFSO = CreateObject("Scripting.FileSystemObject")
	'oFSO.CopyFile "c:\temp\*.xls", "c:\temp\abc.xls", true
	set oFSO = nothing
</script>