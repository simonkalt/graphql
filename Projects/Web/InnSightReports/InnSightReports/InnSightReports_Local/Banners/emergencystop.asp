<% 
	If Application("BMP_Emergency_Stop")=True Then
		If Request("Task")="Get" Or strTask="Get" Then
			If Request("Mode")="HTML" Then
				Response.Redirect Join(strBMParPath, "/") & "blank.gif"
			Else
				If Request.QueryString("Browser")="NETSCAPE4" Then
					Response.Buffer=True
					Response.ContentType="application/x-javascript"
					Response.Write "document.write(' '); "
				Else
					Response.Write " "
				End If
			End If
			Response.End
		Else
			Response.Redirect "down.htm"
		End If
	End If
%>