<br>
<div align="center">
  <center>
  <table border="1" cellpadding="15" cellspacing="0" background="images/tableback.gif" bordercolor="#000000">
    <tr>
      <td>

<p align="center"><font face="Arial" size="4">Impressions In Past 7 Days</font></p>

<div align="center">
  <center>

        <table background="images/tableback.gif">
<tr>
<td align="left" valign="bottom">
<p align="center">
<font face="Arial" size="2">
Date&nbsp;<br>
        
Imp.</font></p>
</td>
<%
lngMax=0
Do While Not rsReport1.EOF
	lngCount=rsReport1("SumOfImpressionCount")
	If lngCount > lngMax Then
		lngMax=lngCount
	End If
	rsReport1.MoveNext
Loop

rsReport1.MoveFirst

Do While Not rsReport1.EOF
	lngCount=rsReport1("SumOfImpressionCount")
	If lngMax > 100 Then
		lngCount=lngCount/(lngMax/100)
	ElseIf lngMax <10 Then 
		lngCount=lngCount*10
	End If
%>
<td align="center" valign="bottom">
<p align="center">
<font face="Arial" size="2">
<img src="images/blue.jpg" width="25" height="<%= lngCount %>"><br>
<%= rsReport1("ImpressionDay") & "<br>" & rsReport1("SumOfImpressionCount") %></font></p>
</td>
<%
rsReport1.MoveNext
Loop
%>
</tr></table>
  </center>
</div>
</td>
    </tr>
  </table>
  </center>
</div>