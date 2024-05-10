<script language="vbscript" runat="server">
	
	Public Function CheckString(s, endchar)
	    pos = InStr(s, "'")
	    While pos > 0
	       s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
	       pos = InStr(pos + 2, s, "'")
	    Wend
	    CheckString = "'" & s & "'" & endchar
	End Function

</script>