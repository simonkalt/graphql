<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<%
' Declare our variables:
Dim objCDO  ' Our CDO object

Dim strTo   ' Strings to hold our email fields
Dim strFrom
Dim strSubject
Dim strBody

Dim intFontSize
Dim intLabelWidth

intFontSize = 2
intLabelWidth = 135

' First we'll read in the values entered and set by
' hand the ones we don't let you enter for our demo.
strTo = Request.Form("to")
'strTo = "simon@innsightreports.com"

' These could read the message subject and body in
' from a form just like the "to" field if we let you
' enter them.
'
'strSubject = Request.Form("subject")
'strBody    = Request.Form("body")
'
' We instead hard code them below just so people
' don't abuse this page.

'***********************************************************
' PLEASE CHANGE THESE SO WE DON'T APPEAR TO BE SENDING YOUR
' EMAIL. WE ALSO DON'T WANT THE EMAILS TO GET SENT TO US
' WHEN SOMETHING GOES WRONG WITH YOUR SCRIPT... THANKS
'***********************************************************
strFrom = "Simon the testing master <simon@wolffdinapoli.com>"

strSubject = Request.Form("Subject")

' This is multi-lined simply for readability.
' Notice that it is a properly formatted HTML
' message and not just plain text like most email.
' A lot of people have asked how to use form data
' in the emails so I added an example of including
' the entered address in the body of the email.
strBody = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbCrLf _
		& "<html>" & vbCrLf _
		& "<head>" & vbCrLf _
		& " <title>Test e-mail from Simon in HTML format</title>" & vbCrLf _
		& " <meta http-equiv=Content-Type content=""text/html; charset=iso-8859-1"">" & vbCrLf _
		& " <meta http-equiv=Content-Type inline=""image/jpeg; 1_logo.jpg"">" & vbCrLf _
		& "</head>" & vbCrLf _
		& "<body bgcolor=""#FAD667"">" & vbCrLf _
		& " <center><img border=0 src=1_logo.jpg></center>" & vbCrLf & vbCrLf & vbCrLf _
		& " <br><br><table width=100% bgcolor=black><tr valign=middle><td><font face=Tahoma size=3 color=white><strong>Sea World</strong></font></td></tr></table>" & vbCrLf _
		& " <p>" & vbCrLf _
		& "  <table width=100% border=0 cellpadding=1>" _
		& "  <tr valign=top><td width=" & intLabelWidth & "><font face=Tahoma size=" & intFontSize & "><strong>Phone:</strong></td><td><font face=Tahoma size=2>(555) 555-1212</td></tr>" _
		& "  <tr valign=top><td width=" & intLabelWidth & "><font face=Tahoma size=" & intFontSize & "><strong>Fax:</strong></td><td><font face=Tahoma size=2>(555) 555-1212</td></tr>" _
		& "  <tr valign=top><td width=" & intLabelWidth & "><font face=Tahoma size=" & intFontSize & "><strong>Street:</strong></td><td><font face=Tahoma size=2>500 Sea World Drive</td></tr>" _
		& "  <tr valign=top><td width=" & intLabelWidth & "><font face=Tahoma size=" & intFontSize & "><strong>City, State  Zip:</strong></td><td><font face=Tahoma size=2>San Diego, CA  93223</td></tr>" _
		& "  <tr valign=top><td width=" & intLabelWidth & "><font face=Tahoma size=" & intFontSize & "><strong>Website:</strong></td><td><font face=Tahoma size=2>http://www.seaworld.com</td></tr>" _
		& "  </table><hr>" _

		& "  <table width=100% border=0 cellpadding=1>" _
		& "  <tr valign=top><td width=" & intLabelWidth & "><font face=Tahoma size=" & intFontSize & "><strong>Transportation:</strong></td><td><font face=Tahoma size=2>Bus from hotel</td></tr>" _
		& "  <tr valign=top><td width=" & intLabelWidth & "><font face=Tahoma size=" & intFontSize & "><strong>Price:</strong></td><td><font face=Tahoma size=2>$43 Adults<br>$12 Kids</td></tr>" _
		& "  <tr valign=top><td width=" & intLabelWidth & "><font face=Tahoma size=" & intFontSize & "><strong>Hours:</strong></td><td><font face=Tahoma size=2>9am-11pm</td></tr>" _
		& "  <tr valign=top><td width=" & intLabelWidth & "><font face=Tahoma size=" & intFontSize & "><strong>Parking:</strong></td><td><font face=Tahoma size=2>$12 per day</td></tr>" _
		& "  <tr valign=top><td width=" & intLabelWidth & "><font face=Tahoma size=" & intFontSize & "><strong>Notes:</strong></td><td><font face=Tahoma size=2>Family fun and entertainment.  Play cards with sharks and crabs.  Shamu rides with every can of dolphin-safe tuna at gate.</td></tr>" _
		& "  </table><hr>" _

		& "  <table width=100% border=0 cellpadding=1>" _
		& "  <tr valign=top><td width=" & intLabelWidth & "><font face=Tahoma size=" & intFontSize & "><strong>Directions:<br>(from hotel)</strong></td><td><font face=Tahoma size=2>Depart Four Seasons Resort Aviara, North San Diego on Four Seasons Pt (North-West) 0.2 mile(s)" & "<br>" _
		& "1: Continue (North) on Local road(s)" & "<br>" _ 
		& "2: Turn LEFT (West) onto Aviara Pky 0.8 mile(s)" & "<br>" _
		& "3: Turn LEFT (West) onto Poinsettia Ln 1.1 mile(s)" & "<br>" _
		& "4: Turn LEFT (South) onto Ramp 0.2 mile(s)" & "<br>" _
		& "5: Merge onto I-5 (San Diego Fwy) (South) 24.4 mile(s)" & "<br>" _
		& "6: Turn off onto Ramp 0.2 mile(s)" & "<br>" _
		& "7: Bear RIGHT (West) onto Sea World Dr 1.8 mile(s)" & "<br>" _
		& "Arrive Sea World" & "<br>" _
		& "Total Route 28.7 mi    36 mins" _
		& "  </td></tr>" _
		& "  <tr valign=top><td width=" & intLabelWidth & "><font face=Tahoma size=" & intFontSize & "><strong>Directions:<br>(to hotel)</strong></td><td><font face=Tahoma size=2>Just go backwards</td></tr>" _
		& "  </table><hr>" _
		
		& "  <table width=100% border=0 cellpadding=1>" _
		& "  <tr><td><center><img border=0 src=http://www.wbwd.net/innsight/maps/1_to_2179.jpg></center></td></tr>" _
		& "  </table><hr>" _
		
		& " </p>" & vbCrLf _
		& " <font size=""-1"">" & vbCrLf _
		& "  <p>Please address all concerns to simon@wolffdinapoli.com.</p>" _
		& "  <p>This message was sent to: " & strTo & "</p>" & vbCrLf _
		& " </font>" & vbCrLf _
		& "</body>" & vbCrLf _
		& "</html>" & vbCrLf

' Some lines to help you check the formatting of your
' email before you actually start sending it to people.
'Response.Write "<pre>"
'Response.Write Server.HTMLEncode(strbody)
'Response.Write "</pre>"
'Response.End


' Ok... we've got all our values so let's get emailing:

' We just check to see if someone has entered anything into the to field.
' If it's equal to nothing we show the form, otherwise we send the message.
' If you were doing this for real you might want to check other fields too
' and do a little entry validation like checking for valid syntax etc.

' Note: I was getting so many bad addresses being entered and bounced
' back to me by mailservers that I've added a quick validation routine.
If strTo = "" Or Not IsValidEmail(strTo) Then
	%>
	<form action="<%= Request.ServerVariables("URL") %>" METHOD="post">
		Enter your e-mail address:<br />
		<input type="text" name="to" size="30" />
		<input type="submit" value="Send Mail!" />
	</form>
	<%
Else
	' Send our message:
	' Note that I'm using the Win2000 CDO and not CDONTS!
	' As such it will only work on Win2000.
	'Set objCDO = Server.CreateObject("CDO.Message")
	'With objCDO
	'	.To       = strTo
	'	.From     = strFrom
	'	.Subject  = strSubject
	'	.HtmlBody = strBody
	'	.Send
	'End With
	'Set objCDO = Nothing

	'==============================================================
	' You'd normally use the above, but I thought I should include
	' the CDONTS version for those of you still running NT4.
	'==============================================================
	Set objCDO = Server.CreateObject("CDONTS.NewMail")
	objCDO.From    = strFrom
	objCDO.To      = strTo
	objCDO.Subject = strSubject
	objCDO.Body    = strBody
	
	objCDO.BodyFormat = 0 ' CdoBodyFormatHTML
	objCDO.MailFormat = 0 ' CdoMailFormatMime
	objCDO.AttachFile "d:\inetpub\wwwroot\wbwd\innsight\ClientUploads\1_logo.jpg"
	
	if trim(Request.Form("cc")) <> "" then
		objCDO.Cc  = Request.Form("cc")
	end if
	'objCDO.Bcc = "user@domain.com;user@domain.com"
	'
	' Send the message!
	objCDO.Send
	Set objCDO = Nothing

	Response.Write "Message sent to " & strTo & "!"
End If
%>

<% ' Only functions and subs follow!

' A quick email syntax checker.  It's not perfect,
' but it's quick and easy and will catch most of
' the bad addresses than people type in.
Function IsValidEmail(strEmail)
	Dim bIsValid
	bIsValid = True
	
	If Len(strEmail) < 5 Then
		bIsValid = False
	Else
		If Instr(1, strEmail, " ") <> 0 Then
			bIsValid = False
		Else
			If InStr(1, strEmail, "@", 1) < 2 Then
				bIsValid = False
			Else
				If InStrRev(strEmail, ".") < InStr(1, strEmail, "@", 1) + 2 Then
					bIsValid = False
				End If
			End If
		End If
	End If

	IsValidEmail = bIsValid
End Function
%>


</BODY>
</HTML>
