<% 
strTargetURL="http://www.vbcode.com"
strImageURL="images/test.gif"
strWidth="468"
strHeight="60"
strAltText="Visit our Sponsors"
strAlign="Center"
strBorder="1"
strTextUnderneath="Nothing"
strAdCode="<a href=" & Chr(34) & strTargetURL & Chr(34) & "><img src=" & Chr(34) & strImageURL & Chr(34)
strAdCode=strAdCode & "  width=" & Chr(34) & strWidth & Chr(34) & " height=" & Chr(34) & strHeight & Chr(34) & " alt=" & Chr(34) & strAltText & Chr(34) & " align=" & Chr(34) &  strAlign & Chr(34) & " border=" & Chr(34) & strBorder & Chr(34) & "></a><br>"
strAdCode=strAdCode & "  <a href="  & Chr(34) & strTargetURL & Chr(34) &  ">" & strTextUnderneath & "</a>"
%>