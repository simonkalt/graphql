<%@ Language=VBScript %>

<!--#INCLUDE file="checkuser.asp"-->
<%
If Session("ScreenHeight") < 750 Then
  GridHeight = 250
Else
  GridHeight = 450
End If
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<%

	Set cnSQL = Server.CreateObject("ADODB.Connection")
	Set rsSQL = Server.CreateObject("ADODB.Recordset")
	Set rsSQLCount = Server.CreateObject("ADODB.Recordset")
  
	cnSQL.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")


	Select Case Request.QueryString("Sort")
		Case "Name"
			Set rsSQL = cnSQL.Execute("sp_GetLocations " & Session("CompanyID") & ", 'Name'")
		Case "Stars"
			Set rsSQL = cnSQL.Execute("sp_GetLocations " & Session("CompanyID") & ", 'Stars'")
		Case "Cost"
			Set rsSQL = cnSQL.Execute("sp_GetLocations " & Session("CompanyID") & ", 'Cost'")
		Case Else
			Set rsSQL = cnSQL.Execute("sp_GetLocations " & Session("CompanyID") & ", 'Name'")
	End Select
	Set rsSQLCount = cnSQL.Execute("sp_GetLocationsCount " & Session("CompanyID"))
  
%>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Dim MaxCol          ' Number of columns
Dim MaxRow          ' Number of rows
Dim GridArray()		' Array to store the data




    
    MaxCol = 13
    MaxRow = <%=rsSQLCount.Fields(0).Value%>
    ID_Col = MaxCol - 1


    If MaxRow > 0 Then

        ' If MaxRow = 0, then (MaxRow - 1) equals -1. This

        ' causes an error in the statement below, so we

        ' handle this special case in the Else clause.

        ReDim GridArray(MaxCol - 1, MaxRow - 1)

    Else

        ReDim GridArray(MaxCol - 1, 0)

    End If

    

<%
	J = 0
    Do While Not rsSQL.EOF
      If J Mod 2 = 0 Then
      %>
        GridArray(0, <%=J%>) = False
      <%
      Else
      %>
        GridArray(0, <%=J%>) = False
      <%
      End If
      %>
      GridArray(0, <%=J%>) = False
      GridArray(1, <%=J%>) = "<%=rsSQL.Fields("CompanyName")%>"
      GridArray(2, <%=J%>) = "<%=rsSQL.Fields("Street")%>"
      GridArray(3, <%=J%>) = "<%=rsSQL.Fields("City")%>"
      GridArray(4, <%=J%>) = "<%=rsSQL.Fields("State")%>"
      GridArray(5, <%=J%>) = "<%=rsSQL.Fields("Phone")%>"
      GridArray(6, <%=J%>) = "<%=rsSQL.Fields("DirectionsMap")%>"
      GridArray(7, <%=J%>) = "<%=rsSQL.Fields("CouponText")%>"
      GridArray(8, <%=J%>) = "<%=rsSQL.Fields("MenuWebsite")%>"
      GridArray(9, <%=J%>) = "<%=rsSQL.Fields("Website")%>"
      GridArray(10, <%=J%>) = "<%=rsSQL.Fields("HotelRating")%>"
      GridArray(11, <%=J%>) = "<%=rsSQL.Fields("CostRating")%>"
      GridArray(ID_COL, <%=J%>) = "<%=rsSQL.Fields("LocationID")%>"
      <%
      J = J + 1
      rsSQL.MoveNext
    Loop
%>


Sub txtSelect_onkeyup

	Do While window.TDBGrid1.SelBookmarks.Count > 0
		window.TDBGrid1.SelBookmarks.Remove (window.TDBGrid1.SelBookmarks.Count - 1)
	Loop
	
	intLen = Len(window.txtSelect.value)
	strText = UCase(Trim(window.txtSelect.value))

	For intCnt = 0 to UBound(GridArray, 2)
		If strText = UCase(Left(GridArray(1,intCnt), intLen)) Then
			window.TDBGrid1.SelBookmarks.Add CStr(intCnt)
			window.TDBGrid1.MoveRelative intCnt, CStr(0)
			Exit For
		End If
	Next	

End Sub

Sub cmdAddItem_onclick
	Set objOpt = document.createElement("option")
	objOpt.Value = window.lstLocations.options(window.lstLocations.selectedIndex).value
	objOpt.Text = window.lstLocations.options(window.lstLocations.selectedIndex).text
	window.lstSelected.options.add(objOpt)
End Sub

Sub cmdRemoveItem_onclick
  window.lstSelected.remove(window.lstSelected.selectedIndex)
End Sub

Sub button5_onclick
    Set C = window.TDBGrid1.Columns.Add (2)
    C.Visible = True
    
    Set C = window.TDBGrid1.Columns.Add(3)
    C.Visible = True

    Set Col2 = window.TDBGrid1.Columns(2)
    Set Col3 = window.TDBGrid1.Columns(3)

    

    ' Set column heading text

    Col2.Caption = "Column 2"
    Col3.Caption = "Column 3"

End Sub

Sub TDBGrid1_UnboundGetRelativeBookmark(StartLocation, ByVal offset, NewLocation, ApproximatePosition)
' TDBGrid1 calls this routine each time it needs to

' reposition itself. StartLocation is a bookmark

' supplied by the grid to indicate the "current"

' position -- the row we are moving from. Offset is

' the number of rows we must move from StartLocation

' in order to arrive at the desired destination row.

' A positive offset means the desired record is after

' the StartLocation, and a negative offset means the

' desired record is before StartLocation.

' If StartLocation is NULL, then we are positioning

' from either BOF or EOF. Once we determine the

' correct index for StartLocation, we will simply add

' the offset to get the correct destination row.

' GetRelativeBookmark already does all of this, so we

' just call it here.

    NewLocation = GetRelativeBookmark(StartLocation, offset)

' If we are on a valid data row (i.e., not at BOF or

' EOF), then set the ApproximatePosition (the ordinal

' row number) to improve scroll bar accuracy. We can

' call IndexFromBookmark to do this.

    If Not IsNull(NewLocation) Then

       ApproximatePosition = IndexFromBookmark(NewLocation, 0)

    End If




End Sub

Sub TDBGrid1_UnboundReadData(ByVal RowBuf, StartLocation, ByVal ReadPriorRows)
' UnboundReadData is fired by an unbound grid whenever

' it requires data for display. This event will fire

' when the grid is first shown, when Refresh or ReBind

' is used, when the grid is scrolled, and after a

' record in the grid is modified and the user commits

' the change by moving off of the current row. The

' grid fetches data in "chunks", and the number of rows

' the grid is asking for is given by RowBuf.RowCount.

' RowBuf is the row buffer where you place the data and

' the bookmarks for the rows that the grid is requesting

' to display. It will also hold the number of rows that

' were successfully supplied to the grid.

' StartLocation is a bookmark which specifies the row

' before or after the desired data, depending on the

' value of ReadPriorRows. If we are reading rows after

' StartLocation (ReadPriorRows = False), then the first

' row of data the grid wants is the row after

' StartLocation, and if we are reading rows before

' StartLocation (ReadPriorRows = True) then the first

' row of data the grid wants is the row before

' StartLocation.

' ReadPriorRows is a boolean value indicating whether

' we are reading rows before (ReadPriorRows = True) or

' after (ReadPriorRows = False) StartLocation.


    Dim Bookmark

    Dim I, RelPos

    Dim J, RowsFetched

    

' Get a bookmark for the start location

    Bookmark = StartLocation

        

    If ReadPriorRows Then

        RelPos = -1 ' Requesting data in rows prior to

                    ' StartLocation

    Else

        RelPos = 1  ' Requesting data in rows after

                    ' StartLocation

    End If

    

    RowsFetched = 0

    For I = 0 To RowBuf.RowCount - 1

        ' Get the bookmark of the next available row

        Bookmark = GetRelativeBookmark(Bookmark, RelPos)

    

        ' If the next row is BOF or EOF, then stop

        ' fetching and return any rows fetched up to this

        ' point.

        If IsNull(Bookmark) Then Exit For

    

        ' Place the record data into the row buffer

        For J = 0 To RowBuf.ColumnCount - 1
            'Debug.Print I, J, GetUserData(Bookmark, J)
            RowBuf.Value(I, J) = GetUserData(Bookmark, J)

        Next

    

        ' Set the bookmark for the row

        RowBuf.Bookmark(I) = Bookmark

    

        ' Increment the count of fetched rows

        RowsFetched = RowsFetched + 1

    Next

    

' Tell the grid how many rows we fetched

    RowBuf.RowCount = RowsFetched



End Sub

Sub TDBGrid1_UnboundWriteData(ByVal RowBuf, WriteLocation)
  GridArray(0, WriteLocation) = Not GridArray(0, WriteLocation)
End Sub

Sub TDBGrid1_PostEvent(ByVal MsgId)
Select Case MsgId

    Case 0

        Exit Sub

    Case 1

        'Data1.Refresh

    Case 2

        ' Handle Mouseclick Event
        If window.TDBGrid1.Col = 8 Then
          'MsgBox "handle go to web: " & GridArray(TDBGrid1.Col,window.TDBGrid1.Bookmark)
          If Len(GridArray(TDBGrid1.Col,window.TDBGrid1.Bookmark)) > 0 Then
            window.location.href = "http://" & GridArray(TDBGrid1.Col,window.TDBGrid1.Bookmark)
          End If
        End If
        If window.TDBGrid1.Col = 9 Then
          'MsgBox "handle go to web: " & GridArray(TDBGrid1.Col,window.TDBGrid1.Bookmark)
          If Len(GridArray(TDBGrid1.Col,window.TDBGrid1.Bookmark)) > 0 Then
            window.location.href = "http://" & GridArray(TDBGrid1.Col,window.TDBGrid1.Bookmark)
          End If
        End If

        If window.TDBGrid1.Col = 0 Then
			window.TDBGrid1.Update
        End If

    Case 3
        'Debug.Print "Update "
        window.TDBGrid1.Update
    Case 4

End Select

End Sub

Sub TDBGrid1_Click
  TDBGrid1.PostMsg 2
End Sub

Sub TDBGrid1_BeforeUpdate(Cancel)
  window.TDBGrid1.PostMsg 3
End Sub

Sub button6_onclick
Dim intRnd
  Do While window.TDBGrid1.SelBookmarks.Count > 0
    window.TDBGrid1.SelBookmarks.Remove (window.TDBGrid1.SelBookmarks.Count - 1)
  Loop
  intRnd = UBound(GridArray, 2)
  Randomize
  intRnd = cInt((intRnd - 0) * Rnd + 0)


  window.TDBGrid1.SelBookmarks.Add CStr(intRnd)
  window.TDBGrid1.MoveRelative intRnd, CStr(0)
End Sub

Sub TDBGrid1_HeadClick(ByVal ColIndex)
  Select Case colIndex
    Case 10
		'MsgBox "10"
		window.location.href = "SearchByLocation3.asp?Sort=Stars"
	Case 11
		'MsgBox "11"
		window.location.href = "SearchByLocation3.asp?Sort=Cost"
  End Select
End Sub

Sub cmdViewLocation_onclick
	Dim strLocationList
	Dim intSelCnt 
	
	intSelCnt = 0

	For intCnt = Lbound(GridArray,2) to UBound(GridArray,2)
		If GridArray(0, intCnt) <> False Then
			If intSelCnt = 0 Then
				strLocationList = strLocationList & GridArray(ID_Col,intCnt)
			Else
				strLocationList = strLocationList & "," & GridArray(ID_Col,intCnt)
			End If
			intSelCnt = intSelCnt + 1
		End If
	Next
	
	If intSelCnt = 0 Then
		MsgBox "You must select at least one record to view a report."
	Else
		window.frmViewLocation.txtViewLocationList.value = strLocationList
		window.frmViewLocation.submit
		'window.location.href = "ReportLocation.asp?LocationIDList=" & strLocationList
	End If
End Sub

Sub cmdPrintLocation_onclick
	Dim strLocationList
	Dim intSelCnt 
	
	intSelCnt = 0

	For intCnt = Lbound(GridArray,2) to UBound(GridArray,2)
		If GridArray(0, intCnt) <> False Then
			If intSelCnt = 0 Then
				strLocationList = strLocationList & GridArray(ID_Col,intCnt)
			Else
				strLocationList = strLocationList & "," & GridArray(ID_Col,intCnt)
			End If
			intSelCnt = intSelCnt + 1
		End If
	Next
	
	If intSelCnt = 0 Then
		MsgBox "You must select at least one record to view a report."
	Else
		window.frmPrintLocation.txtPrintLocationList.value = strLocationList
		window.frmPrintLocation.submit
		'window.location.href = "ReportLocation.asp?LocationIDList=" & strLocationList
	End If
End Sub

Sub cmdViewSummary_onclick
	Dim strLocationList
	Dim intSelCnt 
	
	intSelCnt = 0

	For intCnt = Lbound(GridArray,2) to UBound(GridArray,2)
		If GridArray(0, intCnt) <> False Then
			If intSelCnt = 0 Then
				strLocationList = strLocationList & GridArray(ID_Col,intCnt)
			Else
				strLocationList = strLocationList & "," & GridArray(ID_Col,intCnt)
			End If
			intSelCnt = intSelCnt + 1
		End If
	Next
	
	If intSelCnt = 0 Then
		MsgBox "You must select at least one record to view a report."
	Else
		window.frmViewSummary.txtViewSummaryList.value = strLocationList
		window.frmViewSummary.submit
	End If

End Sub

Sub cmdPrintSummary_onclick
	Dim strLocationList
	Dim intSelCnt 
	
	intSelCnt = 0

	For intCnt = Lbound(GridArray,2) to UBound(GridArray,2)
		If GridArray(0, intCnt) <> False Then
			If intSelCnt = 0 Then
				strLocationList = strLocationList & GridArray(ID_Col,intCnt)
			Else
				strLocationList = strLocationList & "," & GridArray(ID_Col,intCnt)
			End If
			intSelCnt = intSelCnt + 1
		End If
	Next
	
	If intSelCnt = 0 Then
		MsgBox "You must select at least one record to view a report."
	Else
		window.frmPrintSummary.txtPrintSummaryList.value = strLocationList
		window.frmPrintSummary.submit
	End If

End Sub


-->
</SCRIPT>

<TITLE>Search By Location</TITLE>
</HEAD>

<body vLink=black aLink=black link=black bgColor=#F9D568 leftMargin=0 topMargin=0 marginwidth="0" marginheight="0"><!--#include file = "Header.inc" ---> 

<table width=600 border=0>
  <tr>
    <td>Search: <INPUT id=txtSelect 
      style="WIDTH: 275px; HEIGHT: 25px" name=txtSelect> 
	</td>
    <TD align=right></TD></tr>
</table>

<table width=750 border=0>
  <OBJECT id=TDBGrid1 style="LEFT: 0px; WIDTH: 750px; TOP: 0px; HEIGHT: <%=GridHeight%>px" 
  codeBase=../../../tdbg6.cab classid=clsid:00028CD1-0000-0000-0000-000000000046 
  data=data:application/x-oleobject;base64,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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////8AAAAABgAAAAEAAAAAAAAAAAAAAAQAAAANAACADgAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////AAAAAAcAAAABAAAAAAAAAAAAAAAEAAAABQAAgAgAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//////////wAAAAAIAAAAAQAAAAAAAAAAAAAABAAAAAgAAIAFAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////8AAAAACQAAAAEAAAAAAAAAAAAAAAQAAAAA//8ACAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////AAAAAAoAAAABAAAAAAAAAAAAAAAEAAAABQAAgAgAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//////////wAAAAALAAAAAQAAAAAAAAAAAAAABAAAAAUAAIAIAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////8AAAAADAAAAAIAAAAAAAAAAAAAABQAAAAPAACAEgAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////AAAAAA0AAAADAAAAAAAAAAAAAAAUAAAADwAAgBIAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//////////wAAAAAOAAAABQAAAAAAAAAAAAAAFAAAAA8AAIASAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////8AAAAADwAAAAcAAAAAAAAAAAAAAAQAAAAFAACACAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////AAAAABAAAAAGAAAAAAAAAAAAAAAEAAAADQAAgA4AAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//////////wAAAAARAAAACAAAAAAAAAAAAAAABAAAAAgAAIAFAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////8AAAAAEgAAAAkAAAAAAAAAAAAAAAQAAAAA//8ACAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////AAAAABMAAAAKAAAAAAAAAAAAAAAEAAAABQAAgAgAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//////////wAAAAAUAAAABAAAAAAAAAAAAAAAEQAAAA8AAIASAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////8AAAAAFQAAAAwAAAAAAAAAAAAAABQAAAAPAACAEgAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////AAAAABYAAAANAAAAAAAAAAAAAAAUAAAADwAAgBIAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//////////wAAAAAXAAAADwAAAAAAAAAAAAAABAAAAAUAAIAIAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////8AAAAAGAAAAAsAAAAAAAAAAAAAAAQAAAAFAACACAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////AAAAABkAAAAMAAAAAAAAAAAAAAAUAAAADwAAgBIAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//////////wAAAAAaAAAADQAAAAAAAAAAAAAAFAAAAA8AAIASAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////8AAAAAGwAAAA8AAAAAAAAAAAAAAAQAAAAFAACACAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////AAAAABwAAAALAAAAAAAAAAAAAAAEAAAABQAAgAgAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//////////wAAAAAdAAAAAAAAAD8IAP8AAAAABAAAAAUAAIAIAACAOQMAAFRhaG9tYQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////8AAAAAHgAAAB0AAADAAgEAAAIAABQAAAAPAACAEgAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////AAAAAB8AAAAdAAAAwAABAAAAAAAUAAAADwAAgBIAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//////////wAAAAAgAAAAHQAAAMAAAAAAAAAABAAAAA0AAIAOAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////8AAAAAIQAAAB4AAAAAAQAAAAAAABEAAAAPAACAEgAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////AAAAACIAAAAdAAAAwAAAAAAAAAAEAAAACAAAgAUAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//////////wAAAAAjAAAAHQAAAIAAAAAAAAAABAAAAAD//wAIAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////8AAAAAJAAAAB0AAAAAAAAAAAAAAAQAAAAFAACACAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////AAAAAB0AAAAAAAAAAAAAAAAAAAAAAAAA//////////8BAAAAAAAAAAAAAAABAAAA/v///wQAAAAAAAAAAAAAAAUAAAD9////AgAAAAAAAAAAAAAAAgAAAOr///8DAAAAAAAAAAAAAAADAAAA/P///wUAAAAAAAAAAAAAAP/////7////BgAAAAAAAAAAAAAABAAAAPr///8HAAAAAAAAAAAAAAD/////8f///wgAAAAAAAAAAAAAAAYAAADv////CQAAAAAAAAAAAAAABwAAAO7///8KAAAAAAAAAAAAAAAIAAAA+f///wsAAAABAAAAAAAAAP/////r////FAAAAAEAAAAAAAAA//////j///8MAAAAAQAAAAAAAAD/////6f///w0AAAABAAAAAAAAAP/////3////DgAAAAEAAAAAAAAA//////b///8QAAAAAQAAAAAAAAD/////9f///w8AAAABAAAAAAAAAP/////w////EQAAAAEAAAAAAAAA/////+3///8SAAAAAQAAAAAAAAD/////7P///xMAAAABAAAAAAAAAP/////0////GAAAAAEAAAABAAAA//////P///8VAAAAAQAAAAEAAAD/////6P///xYAAAABAAAAAQAAAP/////y////FwAAAAEAAAABAAAA//////T///8cAAAAAQAAAAIAAAD/////8////xkAAAABAAAAAgAAAP/////o////GgAAAAEAAAACAAAA//////L///8bAAAAAQAAAAIAAAD/////CAAAAE5vcm1hbAAFBgAAABAAAAAAAAAA/////0AAAAAAAAAAHQAAAEhlYWRpbmcAAAAAAOBzQQYAAHsF0HNBBmxpdEhlYWRlHgAAAEZvb3RpbmcA4HB9BQAAAABgdH0FwB19BTEAAABgBAAAHwAAAFNlbGVjdGVkAAAAAGEAAAAwdH0FEHR9BUAAAAAxAAAAIAAAAENhcHRpb24A8GZ8BdBmfAWwZnwFYGB8BUBgfAUgYHwFIQAAAEhpZ2hsaWdodFJvdwDxfAXg8HwFwPB8BaDwfAWA8HwFIgAAAEV2ZW5Sb3cAkKt8BXCrfAVQq3wFsOp8BRCcfAXg9XwFIwAAAE9kZFJvdwAFEAF9BeAAfQWAAH0FUAB9BbBufQWAbn0FJAAAAAsAAAD//wAAAwAAAAAAAAALAAAAAAAAAAMAAAAAAAAACwAAAAAAAAADAAAAAAAAAAMAAAAAAAAAAwAAAAAAAAADAAAAAAAAAAMAAAAAAAAAAwAAAAAAAAADAAAAAAAAAAMAAAAAAAAAHgAAAAEAAAAAAAAAHgAAAAEAAAAAAAAAHgAAAAEAAAAAAAAAAwAAAAAAAAALAAAAAAAAAAMAAAAAAAAABAAAAAAAAAADAAAA6AMAAAsAAAD//wAACwAAAAAAAAADAAAAAQAAAAMAAAAAAAAAAwAAAAAAAAADAAAAAAAAAAMAAAAAAAAAAwAAAMgAAAADAAAAAAAAAAMAAADAwMAAAwAAAJDQAwA6AAAAAAAAAAkAAABUREJHcmlkMQACAAAADAAAAEFsbG93QWRkTmV3AC8AAAAMAAAAQWxsb3dBcnJvd3MAAQAAAAwAAABBbGxvd0RlbGV0ZQAEAAAADAAAAEFsbG93VXBkYXRlAL0AAAAOAAAAQW5pbWF0ZVdpbmRvdwDAAAAAEwAAAEFuaW1hdGVXaW5kb3dDbG9zZQC+AAAAFwAAAEFuaW1hdGVXaW5kb3dEaXJlY3Rpb24AvwAAABIAAABBbmltYXRlV2luZG93VGltZQD4/f//CwAAAEFwcGVhcmFuY2UACP7//wwAAABCb3JkZXJTdHlsZQD6/f//CAAAAENhcHRpb24AYAAAAAkAAABDZWxsVGlwcwB+AAAADgAAAENlbGxUaXBzRGVsYXkAfQAAAA4AAABDZWxsVGlwc1dpZHRoAI8AAAAOAAAAQ29sdW1uRm9vdGVycwAHAAAADgAAAENvbHVtbkhlYWRlcnMACAAAAAgAAABDb2x1bW5zACUAAAAJAAAARGF0YU1vZGUAxAAAABIAAABEZWFkQXJlYUJhY2tDb2xvcgAKAAAADAAAAERlZkNvbFdpZHRoAFAAAAANAAAARWRpdERyb3BEb3duAF8AAAAKAAAARW1wdHlSb3dzAP79//8IAAAARW5hYmxlZAAwAAAADwAAAEV4cG9zZUNlbGxNb2RlAJEAAAAKAAAARm9vdExpbmVzAAwAAAAKAAAASGVhZExpbmVzAJgAAAALAAAASW5zZXJ0TW9kZQBdAAAADwAAAExheW91dEZpbGVOYW1lAFwAAAALAAAATGF5b3V0TmFtZQCxAAAACgAAAExheW91dFVSTABKAAAADgAAAE1hcnF1ZWVVbmlxdWUAzgAAAAgAAABNYXhSb3dzAKMAAAAKAAAATW91c2VJY29uAKQAAAANAAAATW91c2VQb2ludGVyAIQAAAAMAAAATXVsdGlTZWxlY3QAYQAAAA4AAABNdWx0aXBsZUxpbmVzAJ8AAAAMAAAAT0xFRHJhZ01vZGUAoAAAAAwAAABPTEVEcm9wTW9kZQCXAAAAEQAAAFBpY3R1cmVBZGRuZXdSb3cAlQAAABIAAABQaWN0dXJlQ3VycmVudFJvdwCzAAAAEQAAAFBpY3R1cmVGb290ZXJSb3cAsgAAABEAAABQaWN0dXJlSGVhZGVyUm93AJYAAAATAAAAUGljdHVyZU1vZGlmaWVkUm93ALAAAAATAAAAUGljdHVyZVN0YW5kYXJkUm93ALQAAAALAAAAUHJpbnRJbmZvcwAPAAAAEAAAAFJvd0RpdmlkZXJTdHlsZQAjAAAABwAAAFNwbGl0cwAxAAAAEAAAAFRhYkFjcm9zc1NwbGl0cwAyAAAACgAAAFRhYkFjdGlvbgCZAAAAFwAAAFRyYW5zcGFyZW50Um93UGljdHVyZXMAMwAAABAAAABXcmFwQ2VsbFBvaW50ZXIA0wcAAAkAAABfRXh0ZW50WADUBwAACQAAAF9FeHRlbnRZAAACAAAMAAAAX0xheW91dFR5cGUAEAAAAAsAAABfUm93SGVpZ2h0AAECAAALAAAAX1N0eWxlRGVmcwAEAgAAFgAAAF9XYXNQZXJzaXN0ZWRBc1BpeGVscwA= 
  width=750 height=<%=GridHeight%> VIEWASTEXT></OBJECT>
</table>


<table>
<tr>
<td width=750 align=center valign="center">
<table border=0>
	<tr>
		<td width=150 valign="center">
		<form action=http://www.innsightreports.com/reportlocation.asp?mode=view method=post id=frmViewLocation name=frmViewLocation>
		<INPUT type=HIDDEN id=txtViewLocationList name=txtViewLocationList>
		<INPUT id=cmdViewLocation style="FONT-SIZE: xx-small; WIDTH: 150px; HEIGHT: 25px" type=button value="View Selected Location(s)" name=cmdViewLocation width="200">
		</form>

		<form action=http://www.innsightreports.com/reportlocation.asp?mode=print method=post id=frmPrintLocation name=frmPrintLocation>
		<INPUT type=HIDDEN id=txtPrintLocationList name=txtPrintLocationList>
		<INPUT id=cmdPrintLocation style="FONT-SIZE: xx-small; WIDTH: 150px; HEIGHT: 25px" type=button value="Print Selected Location(s)" name=cmdPrintLocation width="200"> 
		</form>
		</td>
		<td valign="center">
		<IMG SRC="http://www.innsightreports.com/images/PrintDetail.gif" width=75%>
		</td>

		<td width=150 valign="center">

		<form action=http://www.innsightreports.com/reportlocationsummary.asp?mode=view method=post id=frmViewSummary name=frmViewSummary>
		<INPUT type=HIDDEN id=txtViewSummaryList name=txtViewSummaryList>
		<INPUT id=cmdViewSummary style="FONT-SIZE: xx-small; WIDTH: 150px; HEIGHT: 25px" type=button value="View Location Summary" name=cmdViewSummary width="200"> 
		</form>

		<form action=http://www.innsightreports.com/reportlocationsummary.asp?mode=print method=post id=frmPrintSummary name=frmPrintSummary>
		<INPUT type=HIDDEN id=txtPrintSummaryList name=txtPrintSummaryList>
		<INPUT id=cmdPrintSummary style="FONT-SIZE: xx-small; WIDTH: 150px; HEIGHT: 25px" type=button value="Print Location Summary" name=cmdPrintSummary width="200"> 
		</form>
		</td>
		<td valign="center">
		<IMG SRC="http://www.innsightreports.com/images/PrintSummary.gif" width=75%>
		</td>
	</tr>
</table>
</td>
</tr>
</table>



</body>
<SCRIPT LANGUAGE=vbscript>
<!--
    Set C = window.TDBGrid1.Columns.Add (2)
    C.Visible = True
    
    Set C = window.TDBGrid1.Columns.Add(3)
    C.Visible = True

    Set C = window.TDBGrid1.Columns.Add (4)
    C.Visible = True
    
    Set C = window.TDBGrid1.Columns.Add(5)
    C.Visible = True
    Set C = window.TDBGrid1.Columns.Add (6)
    C.Visible = True
    
    Set C = window.TDBGrid1.Columns.Add(7)
    C.Visible = True
    Set C = window.TDBGrid1.Columns.Add (8)
    C.Visible = True
    
    Set C = window.TDBGrid1.Columns.Add(9)
    C.Visible = True
    
    Set C = window.TDBGrid1.Columns.Add (10)
    C.Visible = True
    

    Set C = window.TDBGrid1.Columns.Add (11)
    C.Visible = True


    Set col0 = window.TDBGrid1.Columns(0)
    Set col1 = window.TDBGrid1.Columns(1)
    Set Col2 = window.TDBGrid1.Columns(2)
    Set Col3 = window.TDBGrid1.Columns(3)
    Set Col4 = window.TDBGrid1.Columns(4)
    Set Col5 = window.TDBGrid1.Columns(5)
    Set Col6 = window.TDBGrid1.Columns(6)
    Set Col7 = window.TDBGrid1.Columns(7)
    Set Col8 = window.TDBGrid1.Columns(8)
    Set Col9 = window.TDBGrid1.Columns(9)
    Set Col10 = window.TDBGrid1.Columns(10)
    Set Col11 = window.TDBGrid1.Columns(11)
    

    ' Set column heading text

    'Formatting
    'Selected
    Col0.Caption = "Selected"
    Col0.Width = 25
    
    Col1.Caption = "Name"
    Col2.Caption = "Address"
    Col3.Caption = "City"
    Col4.Caption = "State"
    Col5.Caption = "Phone"
    Col6.Caption = "Map"
    Col7.Caption = "Coupon"
    Col8.Caption = "Menu"
    Col9.Caption = "Web"
    Col10.Caption = "Stars"
    Col11.Caption = "Cost"

	'Formatting
	'Selected
    window.TDBGrid1.Columns(0).Locked = False
    window.TDBGrid1.Columns(0).Width = 25
    window.TDBGrid1.Columns(0).Caption = "Selected"
    window.TDBGrid1.Columns(0).Backcolor = 4223793
    
    
    window.TDBGrid1.Columns(1).Locked = False
    window.TDBGrid1.Columns(1).Width = 250
    window.TDBGrid1.Columns(1).Caption = "Name"
    window.TDBGrid1.Columns(1).Backcolor = 15333886

    window.TDBGrid1.Columns(2).Locked = False
    window.TDBGrid1.Columns(2).Width = 75
    window.TDBGrid1.Columns(2).Caption = "Address"
    window.TDBGrid1.Columns(2).Backcolor = 15333886

    window.TDBGrid1.Columns(3).Locked = False
    window.TDBGrid1.Columns(3).Width = 75
    window.TDBGrid1.Columns(3).Caption = "City"
    window.TDBGrid1.Columns(3).Backcolor = 14612478

    window.TDBGrid1.Columns(4).Locked = False
    window.TDBGrid1.Columns(4).Width = 37.5
    window.TDBGrid1.Columns(4).Caption = "State"
    window.TDBGrid1.Columns(4).Backcolor = 14612478
    
    window.TDBGrid1.Columns(5).Locked = False
    window.TDBGrid1.Columns(5).Width = 37.5
    window.TDBGrid1.Columns(5).Caption = "Phone"
    window.TDBGrid1.Columns(5).Backcolor = 14087422

    window.TDBGrid1.Columns(6).Locked = False
    window.TDBGrid1.Columns(6).Width = 37.5
    window.TDBGrid1.Columns(6).Caption = "Map"
    window.TDBGrid1.Columns(6).Backcolor = 13628158

    window.TDBGrid1.Columns(7).Locked = False
    window.TDBGrid1.Columns(7).Width = 37.5
    window.TDBGrid1.Columns(7).Caption = "Cpn"
    window.TDBGrid1.Columns(7).Backcolor = 12709629

    window.TDBGrid1.Columns(8).Locked = False
    window.TDBGrid1.Columns(8).Width = 37.5
    window.TDBGrid1.Columns(8).Caption = "Menu"
    window.TDBGrid1.Columns(8).Backcolor = 10741500
    window.TDBGrid1.Columns(8).font.underline = true

    window.TDBGrid1.Columns(9).Locked = False
    window.TDBGrid1.Columns(9).Width = 37.5
    window.TDBGrid1.Columns(9).Caption = "Web"
    window.TDBGrid1.Columns(9).Backcolor = 9822715
    window.TDBGrid1.Columns(9).font.underline = true

    window.TDBGrid1.Columns(10).Locked = False
    window.TDBGrid1.Columns(10).Width = 37.5
    window.TDBGrid1.Columns(10).Caption = "Stars"
    window.TDBGrid1.Columns(10).Backcolor = 8641786

    window.TDBGrid1.Columns(11).Locked = False
    window.TDBGrid1.Columns(11).Width = 37.5
    window.TDBGrid1.Columns(11).Caption = "Cost"
    window.TDBGrid1.Columns(11).Backcolor = 6870521

    window.TDBGrid1.ApproxCount = MaxRow


Private Function MakeBookmark(Index)

' This support function is used only by the remaining

' support functions. It is not used directly by the

' unbound events. It is a good idea to create a

' MakeBookmark function such that all bookmarks can be

' created in the same way. Thus the method by which

' bookmarks are created is consistent and easy to

' modify. This function creates a bookmark when given

' an array row index.

' Since we have data stored in an array, we will just

' use the array index as our bookmark. We will convert

' it to a string first, using the CStr function.

    MakeBookmark = CStr(Index)

End Function

Private Function IndexFromBookmark(Bookmark, ReadPriorRows)

' This support function is used only by the remaining

' support functions. It is not used directly by the

' unbound events.

    

' This function is the inverse of MakeBookmark. Given

' a bookmark, IndexFromBookmark returns the row index

' that the given bookmark refers to. If the given

' bookmark is Null, then it refers to BOF or EOF. In

' such a case, we need to use ReadPriorRows to

' distinguish between the two. If ReadPriorRows = True,

' the grid is requesting rows before the current

' location, so we must be at EOF, because no rows exist

' before BOF. Conversely, if ReadPriorRows = False,

' we must be at BOF.

    

    Dim Index

      

    If IsNull(Bookmark) Then

        If ReadPriorRows Then

            ' Bookmark refers to EOF. Since (MaxRow - 1)

            ' is the index of the last record, we can use

            ' an index of (MaxRow) to represent EOF.

            IndexFromBookmark = MaxRow

        Else

            ' Bookmark refers to BOF. Since 0 is the

            ' index of the first record, we can use an

            ' index of -1 to represent BOF.

            IndexFromBookmark = -1

        End If

    Else

        ' Convert string to long integer

        Index = clng(Bookmark)

        

        ' Check to see if the row index is valid:

        '  (0 <= Index < MaxRow).

        ' If not, set it to a large negative number to

        ' indicate that the bookmark is invalid.

        If Index < 0 Or Index >= MaxRow Then Index = -9999

        IndexFromBookmark = Index

    End If

End Function

Private Function GetRelativeBookmark(Bookmark, RelPos)

' GetRelativeBookmark is used to get a bookmark for a

' row that is a given number of rows away from the given

' row. This specific example will always use either -1

' or +1 for a relative position, since we will always be

' retrieving either the row previous to the current one,

' or the row following the current one.

' IndexFromBookmark expects a Bookmark and a Boolean

' value: this Boolean value is True if the next row to

' read is before the current one [in this case,

' (RelPos < 0) is True], or False if the next row to

' read is after the current one [(RelPos < 0) is False].

' This is necessary to distinguish between BOF and EOF

' in the IndexFromBookmark function if our bookmark is

' Null. Once we get the correct row index from

' IndexFromBookmark, we simply add RelPos to it to get

' the desired row index and create a bookmark using

' that index.

    Dim Index

    

    Index = IndexFromBookmark(Bookmark, RelPos < 0) + RelPos

    If Index < 0 Or Index >= MaxRow Then

        ' Index refers to a row before the first or after

        ' the last, so just return Null.

        GetRelativeBookmark = Null

    Else

        GetRelativeBookmark = MakeBookmark(Index)

    End If

End Function

Private Function GetUserData(Bookmark, Col)

' In this example, GetUserData is called by

' UnboundReadData to ask the user what data should be

' displayed in a specific cell in the grid. The grid

' row the cell is in is the one referred to by the

' Bookmark parameter, and the column it is in it given

' by the Col parameter. GetUserData is called on a

' cell-by-cell basis.

    Dim Index

' Figure out which row the bookmark refers to

    Index = IndexFromBookmark(Bookmark, False)

    

    If Index < 0 Or Index >= MaxRow Or Col < 0 Or Col >= MaxCol Then

        ' Cell position is invalid, so just return null

        ' to indicate failure

        GetUserData = Null

    Else

        GetUserData = GridArray(Col, Index)

    End If

End Function

-->
</SCRIPT>

</HTML>
